using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.IO;
using System.Data.OleDb;

namespace ISBL
{
    public class Admin
    {
        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        private SqlDataReader myDataReader;
        #endregion

        #region Constructors
        public Admin()
        {
            conn.setConnection("ocrsConnection");
        }
        #endregion

        #region public function

        //Closes db connection.
        public void DisposeConnection()
        {
            conn.Dispose();
        }

        public DataSet UserList(String strUserID, string strUserType, string strLoginID, string strLoginUserType, string strClientCode)
        {
            string strSQL;
            int intCompare;
            int intIsBDM;
            Boolean blAnd = false;
            ISBL.General oISBLGen = new ISBL.General();

            strSQL = "SELECT lUser.LoginID As LoginID";
            strSQL = strSQL + ",uType.UserTypeDescription As UserTypeDesc";
            strSQL = strSQL + ",uType.UserType As UserType";
            strSQL = strSQL + ",uRole.RoleDescription As [Role]";
            strSQL = strSQL + ",lUser.Active As Active ";
            strSQL = strSQL + ",c.ClientName As ClientName ";
            strSQL = strSQL + ",c.ClientCode As ClientCode ";
            strSQL = strSQL + "FROM ocrsLoginUser lUser ";
            strSQL = strSQL + "INNER JOIN ocrsUserTypeMaster uType ";
            strSQL = strSQL + "ON lUser.UserType = uType.UserType ";
            strSQL = strSQL + "INNER JOIN ocrsRoleMaster uRole ";
            strSQL = strSQL + "ON lUser.Role = uRole.Role ";
            strSQL = strSQL + "LEFT OUTER JOIN ocrsClientUser cu ON lUser.LoginID = cu.LoginID ";
            strSQL = strSQL + "LEFT OUTER JOIN ocrsClientMaster c ON cu.ClientCode = c.ClientCode ";

            intIsBDM = string.Compare(strLoginUserType, "BDM");
            if (intIsBDM == 0)
            {
                string strBDMClient;
                DataSet ds;
                strBDMClient = "";
                int i = 1;
                ds = oISBLGen.GetBDMClient(strLoginID);
                foreach (DataRow dr in ds.Tables[0].Rows)
                {
                    if (i < ds.Tables[0].Rows.Count)
                    {
                        strBDMClient = strBDMClient + "'" + dr["ClientCode"] + "',";
                    }
                    else
                    {
                        strBDMClient = strBDMClient + "'" + dr["ClientCode"] + "'";
                    }
                    i = i + 1;
                }
                strSQL = strSQL + "WHERE cu.ClientCode IN (" + strBDMClient + ") ";
                blAnd = true;
            }

            if (strUserID.Length > 0)
            {
                if (blAnd)
                {
                    strSQL = strSQL + "AND ";
                }
                else
                {
                    strSQL = strSQL + "WHERE ";
                    blAnd = true;
                }
                strSQL = strSQL + " lUser.LoginID LIKE '%" + oISBLGen.SafeSqlLiteral(strUserID) + "%' ";
            }

            intCompare = string.Compare(strUserType, "All");
            if (intCompare != 0)
            {
                if (blAnd)
                {
                    strSQL = strSQL + "AND ";
                }
                else
                {
                    strSQL = strSQL + "WHERE ";
                }
                strSQL = strSQL + "lUser.UserType='" + oISBLGen.SafeSqlLiteral(strUserType) + "' ";

                //added 5-Feb-09 Adam: If client is selected
                if (strClientCode.Length > 0 && strClientCode != "0") 
                    strSQL = strSQL + " AND cu.ClientCode='" + oISBLGen.SafeSqlLiteral(strClientCode) + "' ";
            }

            DataSet dsUser = new DataSet();
            SqlDataAdapter sda = new SqlDataAdapter();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandType = CommandType.Text;
            myCmd.CommandText = strSQL;
            sda.SelectCommand = myCmd;
            conn.Open();
            dsUser = conn.FillDataSet(sda);
            conn.Close();

            sda.Dispose();
            myCmd.Dispose();
            oISBLGen.DisposeConnection();
            oISBLGen = null;

            return dsUser;
        }
        #endregion

#region BlockEmail
        //To insert into ocrsBlockEmail table
        //Adam 9 Dec 2008
        public Boolean InsertBlockEmail(string strCRN, Boolean blnBlockBudget, Boolean blnBlockCancel, Boolean blnBlockDownload)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_InsertBlockEmail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters.Add("@BlockBudget", SqlDbType.Bit);
            myCmd.Parameters["@BlockBudget"].Value = blnBlockBudget;
            myCmd.Parameters.Add("@BlockCancel", SqlDbType.Bit);
            myCmd.Parameters["@BlockCancel"].Value = blnBlockCancel;
            myCmd.Parameters.Add("@BlockDownload", SqlDbType.Bit);
            myCmd.Parameters["@BlockDownload"].Value = blnBlockDownload;
            conn.Open();
            conn.callingMethod = "ISBL.General.InsertBlockEmail";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        //*** Start Adam 30 Jan 2009 View and Update Block Email
        public DataSet GetBlockEmail(string strCRN)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBlockEmail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters["@CRN"].Value = strCRN;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.ViewBlockEmail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

#endregion

        //Start for Upload Alert Cases - Adam June 2008

        public string readAlertCase(String strAttachmentFile)
        {

            ISBL.General oISBLGen = new ISBL.General();
            ISDL.Connect connAlert = new ISDL.Connect();
            connAlert.setConnection("ocrsConnection");


            ISDL.Connect connSaveClosedCases = new ISDL.Connect();
            connSaveClosedCases.setConnection("ocrsConnection");
            Boolean blnStatus;
            Guid CRNId;
            string strSubjectName;
            DateTime dtReportDate;
            string strCountry;
            string strReportType;
            string strIndustryType;
            string strNewsLink;
            Boolean R001;
            Boolean R002;
            Boolean R003;
            Boolean R004;
            Boolean R005;
            Boolean R006;
            Boolean R007;
            Boolean R008;
            Boolean R009;
            Boolean R010;
            Boolean R011;
            Boolean R012;
            Boolean blnRisk;
            int intUserId;
            int intStatusId;
            DateTime dtCreationDate = System.DateTime.Now;

            int intSubjectId;
            Boolean blnPrimary;
            int intSubjectTypeId;

            //Error Messages
            //String strMsgInvalidFile = "The excel file you uploaded is invalid. " +
            //                         "Please ensure that the sheet format is intact.";

            String strExcelConn, strQuery, strSheetName;
            //Connection String to Excel Workbook
            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter ExcelAdp = new OleDbDataAdapter();

            //DataTables
            DataTable dtExcelSchema = new DataTable();
            //DataTable dtExcelTbl = new DataTable();
            //DataSet dsExcel = new DataSet();

            //valid excel file
            try
            {
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strAttachmentFile +
                                ";Extended Properties=Excel 8.0";
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();

                //Find out the Sheet & Range names
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                strSheetName = dtExcelSchema.Rows[1]["TABLE_NAME"].ToString();
            }
            catch (Exception ex)
            {
                //throw new System.InvalidOperationException(strMsgInvalidFile);
                return ex.Message.ToString();
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
            }


            //Read the Complete Block A2:R65536 in Excel File 
            try
            {

                strQuery = "Select *  FROM [" + strSheetName + "A2:R65536]";
                strQuery = strQuery.Replace("'", "");
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);
                dtExcelSchema = new DataTable();
                ExcelAdp.Fill(dtExcelSchema);
            }
            catch (Exception ex)
            {
                //throw new System.InvalidOperationException(strMsgInvalidFile);
                return ex.Message.ToString();
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            if (dtExcelSchema.Rows.Count == 0)
            {
                //throw new System.InvalidOperationException(strMsgInvalidFile);
                return "Excel does not have records";
            }

            int intRowNumber = 0;
            int i = 0;
            try
            {
                connSaveClosedCases.Open();
                connSaveClosedCases.BeginTransaction();
                while (i < dtExcelSchema.Rows.Count)
                {
                    if (dtExcelSchema.Rows[i].IsNull(0))
                    {
                        dtExcelSchema.Rows[i].Delete();
                    }
                    else
                    {

                        CRNId = Guid.NewGuid();
                        intUserId = 4;
                        intStatusId = 4;
                        //strReportType = dtExcelSchema.Rows[i][4].ToString().Trim(); 
                        strReportType = "IEDD1"; //19 Aug 2008 Adam - Change of new template
                        dtReportDate = Convert.ToDateTime(dtExcelSchema.Rows[i][1].ToString().Trim());
                        strNewsLink = dtExcelSchema.Rows[i][16].ToString().Trim();
                        //Insert to dbClosedCases Table

                        blnStatus = SaveClosedCases(ref connSaveClosedCases, CRNId, intUserId, intStatusId, strReportType, dtReportDate, dtCreationDate, strNewsLink);
                        if (!blnStatus)
                        {
                            connSaveClosedCases.RollBackTransaction();
                            connSaveClosedCases.Dispose();
                            connSaveClosedCases = null;

                            //New Addition Start
                            if (File.Exists(strAttachmentFile))
                                File.Delete(strAttachmentFile);
                            //New Addition End

                            return "Saving dbClosedCases failed";
                        }


                        intSubjectId = 1;
                        strSubjectName = dtExcelSchema.Rows[i][0].ToString().Trim();
                        blnPrimary = true;
                        blnRisk = Convert.ToBoolean(dtExcelSchema.Rows[i][17].ToString().Trim());
                        intSubjectTypeId = 2;

                        strCountry = dtExcelSchema.Rows[i][2].ToString().Trim();
                        strIndustryType = dtExcelSchema.Rows[i][3].ToString().Trim();

                        //Then Insert to dbClosedSubjects Table

                        blnStatus = SaveClosedSubjects(ref connSaveClosedCases, CRNId, intSubjectId, strSubjectName, blnPrimary, blnRisk, intSubjectTypeId, strCountry, strIndustryType, dtReportDate, dtCreationDate);
                        if (!blnStatus)
                        {
                            connSaveClosedCases.RollBackTransaction();
                            connSaveClosedCases.Dispose();
                            connSaveClosedCases = null;


                            //New Addition Start
                            if (File.Exists(strAttachmentFile))
                                File.Delete(strAttachmentFile);
                            //New Addition End

                            return "Saving dbClosedSubjects failed";
                        }

                        R001 = ConvertYesToBoolean(dtExcelSchema.Rows[i][4].ToString().Trim());
                        R002 = ConvertYesToBoolean(dtExcelSchema.Rows[i][5].ToString().Trim());
                        R003 = ConvertYesToBoolean(dtExcelSchema.Rows[i][6].ToString().Trim());
                        R005 = ConvertYesToBoolean(dtExcelSchema.Rows[i][7].ToString().Trim());
                        R004 = ConvertYesToBoolean(dtExcelSchema.Rows[i][8].ToString().Trim());
                        R007 = ConvertYesToBoolean(dtExcelSchema.Rows[i][9].ToString().Trim());
                        R006 = ConvertYesToBoolean(dtExcelSchema.Rows[i][10].ToString().Trim());
                        R009 = ConvertYesToBoolean(dtExcelSchema.Rows[i][11].ToString().Trim());
                        R008 = ConvertYesToBoolean(dtExcelSchema.Rows[i][12].ToString().Trim());
                        R010 = ConvertYesToBoolean(dtExcelSchema.Rows[i][13].ToString().Trim());
                        R011 = ConvertYesToBoolean(dtExcelSchema.Rows[i][14].ToString().Trim());
                        R012 = ConvertYesToBoolean(dtExcelSchema.Rows[i][15].ToString().Trim());

                        //Then Insert into dbSubjectRiskList table
                        blnStatus = SaveSubjectRiskList(ref connSaveClosedCases, CRNId, intSubjectId, R001, R002, R003, R004, R005, R006, R007, R008, R009, R010, R011, R012);
                        if (!blnStatus)
                        {
                            connSaveClosedCases.RollBackTransaction();
                            connSaveClosedCases.Dispose();
                            connSaveClosedCases = null;


                            //New Addition Start
                            if (File.Exists(strAttachmentFile))
                                File.Delete(strAttachmentFile);
                            //New Addition End

                            return "Saving dbSubjectRiskList failed";
                        }


                        intRowNumber++;
                    }
                    i++;
                }
            }
            catch (Exception ex)
            {
                //throw new System.InvalidOperationException(strMsgInvalidFile);
                return ex.Message.ToString();
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            //New Addition Start
            if (File.Exists(strAttachmentFile))
                File.Delete(strAttachmentFile);
            //New Addition End

            connSaveClosedCases.CommitTransaction();
            connSaveClosedCases.Dispose();
            connSaveClosedCases = null;

            return "True";
        }

        public Boolean ConvertYesToBoolean(string strValue)
        {
            if (strValue == "Yes")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        private Boolean SaveClosedCases(ref ISDL.Connect connSaveClosedCases, Guid CRNId, int intUserID, int intStatusId, string strReportType, DateTime dtReportDate, DateTime dtCreationDate, string strNewsLink)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connSaveClosedCases.Connection;
            cmd.CommandText = "sp_SaveClosedCases";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connSaveClosedCases.currentTransaction;
            cmd.Parameters.Add("@CRNId", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@intUserID", SqlDbType.Int);
            cmd.Parameters.Add("@intStatusId", SqlDbType.Int);
            cmd.Parameters.Add("@strReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@dtReportDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@dtCreationDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@strNewsLink", SqlDbType.Text );
            cmd.Parameters["@CRNId"].Value = CRNId;
            cmd.Parameters["@intUserID"].Value = intUserID;
            cmd.Parameters["@intStatusId"].Value = intStatusId;
            cmd.Parameters["@strReportType"].Value = strReportType;
            cmd.Parameters["@dtReportDate"].Value = dtReportDate;
            cmd.Parameters["@dtCreationDate"].Value = dtCreationDate;
            cmd.Parameters["@strNewsLink"].Value = strNewsLink;

            strStatus = connSaveClosedCases.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    cmd.Dispose();
                    return true;
                }
                else
                {
                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                cmd.Dispose();
                return false;
            }
        }

        //SaveClosedSubjects Already Obsolete - Jul 2013 Adam
        private Boolean SaveClosedSubjects(ref ISDL.Connect connSaveClosedCases, Guid CRNId, int intSubjectId, string strSubjectName, Boolean blnPrimary, Boolean blnRisk, int intSubjectTypeId, string strCountry, string strIndustryType, DateTime dtReportDate, DateTime dtCreationDate)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connSaveClosedCases.Connection;
            cmd.CommandText = "sp_SaveClosedSubjects";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connSaveClosedCases.currentTransaction;
            cmd.Parameters.Add("@CRNId", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@intSubjectId", SqlDbType.Int);
            cmd.Parameters.Add("@strSubjectName", SqlDbType.NVarChar, 250);
            cmd.Parameters.Add("@blnPrimary", SqlDbType.Bit);
            cmd.Parameters.Add("@blnRisk", SqlDbType.Bit);
            cmd.Parameters.Add("@intSubjectTypeId", SqlDbType.Int);
            cmd.Parameters.Add("@strCountry", SqlDbType.VarChar, 100);
            cmd.Parameters.Add("@strIndustryType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@dtReportDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@dtCreationDate", SqlDbType.DateTime);

            cmd.Parameters["@CRNId"].Value = CRNId;
            cmd.Parameters["@intSubjectId"].Value = intSubjectId;
            cmd.Parameters["@strSubjectName"].Value = strSubjectName;
            cmd.Parameters["@blnPrimary"].Value = blnPrimary;
            cmd.Parameters["@blnRisk"].Value = blnRisk;
            cmd.Parameters["@intSubjectTypeId"].Value = intSubjectTypeId;
            cmd.Parameters["@strCountry"].Value = strCountry;
            cmd.Parameters["@strIndustryType"].Value = strIndustryType;
            cmd.Parameters["@dtReportDate"].Value = dtReportDate;
            cmd.Parameters["@dtCreationDate"].Value = dtCreationDate;

            strStatus = connSaveClosedCases.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    cmd.Dispose();
                    return true;
                }
                else
                {
                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                cmd.Dispose();
                return false;
            }
        }

        private Boolean SaveSubjectRiskList(ref ISDL.Connect connSaveClosedCases, Guid CRNId, int intSubjectId, Boolean R001, Boolean R002, Boolean R003, Boolean R004, Boolean R005, Boolean R006, Boolean R007, Boolean R008, Boolean R009, Boolean R010, Boolean R011, Boolean R012)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connSaveClosedCases.Connection;
            cmd.CommandText = "sp_SaveSubjectRiskList";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connSaveClosedCases.currentTransaction;
            cmd.Parameters.Add("@CRNId", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@intSubjectId", SqlDbType.Int);
            cmd.Parameters.Add("@R001", SqlDbType.Bit);
            cmd.Parameters.Add("@R002", SqlDbType.Bit);
            cmd.Parameters.Add("@R003", SqlDbType.Bit);
            cmd.Parameters.Add("@R004", SqlDbType.Bit);
            cmd.Parameters.Add("@R005", SqlDbType.Bit);
            cmd.Parameters.Add("@R006", SqlDbType.Bit);
            cmd.Parameters.Add("@R007", SqlDbType.Bit);
            cmd.Parameters.Add("@R008", SqlDbType.Bit);
            cmd.Parameters.Add("@R009", SqlDbType.Bit);
            cmd.Parameters.Add("@R010", SqlDbType.Bit);
            cmd.Parameters.Add("@R011", SqlDbType.Bit);
            cmd.Parameters.Add("@R012", SqlDbType.Bit);
            cmd.Parameters["@CRNId"].Value = CRNId;
            cmd.Parameters["@intSubjectId"].Value = intSubjectId;
            cmd.Parameters["@R001"].Value = R001;
            cmd.Parameters["@R002"].Value = R002;
            cmd.Parameters["@R003"].Value = R003;
            cmd.Parameters["@R004"].Value = R004;
            cmd.Parameters["@R005"].Value = R005;
            cmd.Parameters["@R006"].Value = R006;
            cmd.Parameters["@R007"].Value = R007;
            cmd.Parameters["@R008"].Value = R008;
            cmd.Parameters["@R009"].Value = R009;
            cmd.Parameters["@R010"].Value = R010;
            cmd.Parameters["@R011"].Value = R011;
            cmd.Parameters["@R012"].Value = R012;
            strStatus = connSaveClosedCases.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    cmd.Dispose();
                    return true;
                }
                else
                {
                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                cmd.Dispose();
                return false;
            }
        }

        //End for Upload Alert Cases - Adam June 2008

        public DataSet GetNonBDMClient(string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetNonBDMClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetNonBDMClient";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetUserType(Boolean blIncludeAll)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetUserType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@IncludeAll", SqlDbType.Bit);
            myCmd.Parameters["@IncludeAll"].Value = blIncludeAll;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetUserType";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetClient()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClient";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRole(string strUserType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRole";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 10);
            myCmd.Parameters["@UserType"].Value = strUserType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetRole";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
     
 public Boolean InsertNewLogin(string strClientCode, string strLoginID, string strLoginName, string strEmail, String strPassword, string strRole, string strUserType, string strCreatedBy, Boolean blLegacy, Boolean blDashboard, string strGOC, Boolean blOrderManagement, Array arrClientList ,
     string strUPN, string strADFSEmail, string strCName, string strFName, string strLName, string strGUnit, string strLocation, string strTelNo, string strWorkDesc, Boolean blMercuryID)
        {
            string strEncryptPassword;
            string strBody;
            string strSuccess;
            ISBL.General oISBLGen = new ISBL.General();
            try
            {
                if (oISBLGen.ISADFSClient(strClientCode)) // if is ADFS Client, update ADFS Claims Info to ocrsADFSIDMapping table
                {
                    strEncryptPassword = DataEncryption.Encrypt(strPassword);
                    SqlCommand myCmd = new SqlCommand();
                    myCmd.Connection = conn.Connection;
                    myCmd.CommandText = "sp_InsertNewLoginADFS";
                    myCmd.CommandType = CommandType.StoredProcedure;
                    myCmd.Parameters.Add("@ClientID", SqlDbType.VarChar, 30);
                    myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                    myCmd.Parameters.Add("@LoginName", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@Email", SqlDbType.VarChar, 200);
                    myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
                    myCmd.Parameters.Add("@Role", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@Legacy", SqlDbType.Bit);
                    myCmd.Parameters.Add("@Dashboard", SqlDbType.Bit);
                    myCmd.Parameters.Add("@GOC", SqlDbType.VarChar, 5);
                    myCmd.Parameters.Add("@MercuryID", SqlDbType.Bit); //Mercury Enhancement - Jan 2014
                    myCmd.Parameters.Add("@OrderManagement", SqlDbType.Bit);


                    myCmd.Parameters.Add("@UPN", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@ADFSEmail", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@CName", SqlDbType.NVarChar, 250);
                    myCmd.Parameters.Add("@FName", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@LName", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@GUnit", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@Location", SqlDbType.VarChar, 250);
                    myCmd.Parameters.Add("@TelNo", SqlDbType.VarChar, 50);
                    myCmd.Parameters.Add("@WorkDesc", SqlDbType.VarChar, 250);


                    myCmd.Parameters["@ClientID"].Value = strClientCode;
                    myCmd.Parameters["@LoginID"].Value = strLoginID;
                    myCmd.Parameters["@LoginName"].Value = strLoginName;
                    myCmd.Parameters["@Email"].Value = strEmail;
                    myCmd.Parameters["@Password"].Value = strEncryptPassword;
                    myCmd.Parameters["@Role"].Value = strRole;
                    myCmd.Parameters["@UserType"].Value = strUserType;
                    myCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                    myCmd.Parameters["@Legacy"].Value = blLegacy;
                    myCmd.Parameters["@Dashboard"].Value = blDashboard;
                    myCmd.Parameters["@GOC"].Value = strGOC;
                    myCmd.Parameters["@MercuryID"].Value = blMercuryID; //Mercury Enhancement - Jan 2014
                    myCmd.Parameters["@OrderManagement"].Value = blOrderManagement;


                    myCmd.Parameters["@UPN"].Value = strUPN;
                    myCmd.Parameters["@ADFSEmail"].Value = strADFSEmail;
                    myCmd.Parameters["@CName"].Value = strCName;
                    myCmd.Parameters["@FName"].Value = strFName;
                    myCmd.Parameters["@LName"].Value = strLName;
                    myCmd.Parameters["@GUnit"].Value = strGUnit;
                    myCmd.Parameters["@Location"].Value = strLocation;
                    myCmd.Parameters["@TelNo"].Value = strTelNo;
                    myCmd.Parameters["@WorkDesc"].Value = strWorkDesc;


                    conn.Open();
                    conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertNewLoginADFS";
                    strSuccess = conn.cmdScalarStoredProc(myCmd);
                    conn.Close();
                    myCmd.Dispose();
                }
                else
                {
                    strEncryptPassword = DataEncryption.Encrypt(strPassword);
                    SqlCommand myCmd = new SqlCommand();
                    myCmd.Connection = conn.Connection;
                    myCmd.CommandText = "sp_InsertNewLogin";
                    myCmd.CommandType = CommandType.StoredProcedure;
                    myCmd.Parameters.Add("@ClientID", SqlDbType.VarChar, 30);
                    myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                    myCmd.Parameters.Add("@LoginName", SqlDbType.VarChar, 100);
                    myCmd.Parameters.Add("@Email", SqlDbType.VarChar, 200);
                    myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
                    myCmd.Parameters.Add("@Role", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                    myCmd.Parameters.Add("@Legacy", SqlDbType.Bit);
                    myCmd.Parameters.Add("@Dashboard", SqlDbType.Bit);
                    myCmd.Parameters.Add("@GOC", SqlDbType.VarChar, 5);
                    myCmd.Parameters.Add("@MercuryID", SqlDbType.Bit); //Mercury Enhancement - Jan 2014
                    myCmd.Parameters.Add("@OrderManagement", SqlDbType.Bit);
                    myCmd.Parameters["@ClientID"].Value = strClientCode;
                    myCmd.Parameters["@LoginID"].Value = strLoginID;
                    myCmd.Parameters["@LoginName"].Value = strLoginName;
                    myCmd.Parameters["@Email"].Value = strEmail;
                    myCmd.Parameters["@Password"].Value = strEncryptPassword;
                    myCmd.Parameters["@Role"].Value = strRole;
                    myCmd.Parameters["@UserType"].Value = strUserType;
                    myCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                    myCmd.Parameters["@Legacy"].Value = blLegacy;
                    myCmd.Parameters["@Dashboard"].Value = blDashboard;
                    myCmd.Parameters["@GOC"].Value = strGOC;
                    myCmd.Parameters["@MercuryID"].Value = blMercuryID; //Mercury Enhancement - Jan 2014
                    myCmd.Parameters["@OrderManagement"].Value = blOrderManagement;
                    conn.Open();
                    conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertNewLogin";
                    strSuccess = conn.cmdScalarStoredProc(myCmd);
                    conn.Close();
                    myCmd.Dispose();
                }
                int compare = string.Compare(strSuccess, "1");
                if (compare == 0)
                {
                    if (strUserType == "BDM")
                    {
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.sp_InsertClientCaseManager.sp_InsertClientCaseManager";
                        foreach (string str in arrClientList)
                        {
                            SqlCommand cmd = new SqlCommand();
                            cmd.Connection = conn.Connection;
                            cmd.CommandText = "sp_InsertClientCaseManager";
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@ClientID", SqlDbType.VarChar, 30);
                            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                            cmd.Parameters["@ClientID"].Value = str;
                            cmd.Parameters["@LoginID"].Value = strLoginID;
                            cmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                            conn.cmdScalarStoredProc(cmd);
                            cmd.Dispose();
                        }
                        conn.Close();
                    }
                    else if (strRole == "C4")
                    {   // Added By Stev, 20 Oct 08. Insert multiple Client that associated to C4 Login ID.
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertC4Client";
                        foreach (string sClCode in arrClientList)
                        {
                            SqlCommand sqlCmd = new SqlCommand();
                            sqlCmd.Connection = conn.Connection;
                            sqlCmd.CommandText = "sp_InsertC4Client";
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                            sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                            sqlCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                            sqlCmd.Parameters["@LoginId"].Value = strLoginID;
                            sqlCmd.Parameters["@ClCode"].Value = sClCode;
                            sqlCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                            conn.cmdScalarStoredProc(sqlCmd);
                            sqlCmd.Dispose();
                        }
                        conn.Close();
                    }
                    else if (strRole == "C5")
                    {
                        // Added By Adam, 1 Jul 10. Insert multiple Client that associated to C5 Login ID.
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertC5Client";
                        foreach (string sClCode in arrClientList)
                        {
                            SqlCommand sqlCmd = new SqlCommand();
                            sqlCmd.Connection = conn.Connection;
                            sqlCmd.CommandText = "sp_InsertC5Client";
                            sqlCmd.CommandType = CommandType.StoredProcedure;
                            sqlCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                            sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                            sqlCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                            sqlCmd.Parameters["@LoginId"].Value = strLoginID;
                            sqlCmd.Parameters["@ClCode"].Value = sClCode;
                            sqlCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                            conn.cmdScalarStoredProc(sqlCmd);
                            sqlCmd.Dispose();
                        }
                        conn.Close();
                    }

                    strBody = "<Table><Tr><Td>Your New Login has been created.</Td></Tr>";
                    strBody += "<Tr><Td>Your new login information as below:</Td></Tr>";
                    strBody += "<Tr><Td>Your LoginID is: " + strLoginID + "</Td></Tr>";
                    strBody += "<Tr><Td>Your new Password is: " + strPassword + "</Td></Tr>";
                    strBody += "<Tr><Td height='50'>" + System.Configuration.ConfigurationManager.AppSettings["ISISURL"] + "</Td></Tr>";
                    strBody += "<Tr><Td height='50'>Thank you.</Td></Tr>";
                    //Start Change Email Signed - Feb 2013 Adam
                    //strBody += "<Tr><Td height='30'>Global World-Check</Td></Tr>";
                    strBody += "<Tr><Td height='30'>Thomson Reuters Risk</Td></Tr>";
                    strBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></Td></Tr>";
                    strBody += "<Tr><Td height='20'></Td></Tr>";
                    //End Change Email Signed - Feb 2013 Adam
                    strBody += "</Table>";

                    int strCompare = string.Compare(strUserType, "CLIENT", true);
                    if (strCompare == 0)
                    {
                        Boolean blnSendEmail;
                       
                        if (oISBLGen.ISADFSClient(strClientCode)) // if is ADFS Client, do not send out email
                            blnSendEmail = true;
                        else{
                            if (blMercuryID) //Mercury Enhancement - Jan 2014 - no email notification to be sent
                                blnSendEmail = true; //Mercury Enhancement - Jan 2014 - no email notification to be sent
                            else
                                blnSendEmail = oISBLGen.sendemail(strEmail, System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"], "New Login Created", strBody, true);
                            }
                        oISBLGen = null;
                        return true;
                    }
                    return true;
                }
                else
                {
                    return false;
                }

            }
            catch (Exception)
            {
                return false;
            }
        }

        public Boolean CheckUserIDExist(string strLoginID)
        {
            // Amented by Stev, 23 Oct 08. Return true if UserId founded, else false.
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckUserIDExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckUserIDExist";
            string count = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (count == "0")
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        //Mitul - 13May2013 - Capture Client Browser and OS details and log them - START
        public DataSet ValidateLogin(string strLoginID, string strPassword, string strIPAddress, string strClientBrowser, string strClientBrowserVersion, string strUserAgent, String strJSEnabled, Boolean blnRemovePreviousInstance)
        {
            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_ValidateLoginBrowserInfo", conn.Connection);
            myCmd.SelectCommand.CommandType = CommandType.StoredProcedure;
            myCmd.SelectCommand.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.SelectCommand.Parameters.Add("@Password", SqlDbType.VarChar, 100);
            myCmd.SelectCommand.Parameters.Add("@InstanceID", SqlDbType.UniqueIdentifier);
            myCmd.SelectCommand.Parameters.Add("@IPAddress", SqlDbType.VarChar, 20);
            myCmd.SelectCommand.Parameters.Add("@ClientBrowser", SqlDbType.VarChar, 20);
            myCmd.SelectCommand.Parameters.Add("@ClientBrowserVersion", SqlDbType.VarChar, 20);
            myCmd.SelectCommand.Parameters.Add("@UserAgent", SqlDbType.VarChar, 200);
            myCmd.SelectCommand.Parameters.Add("@JSEnabled", SqlDbType.VarChar, 20);
            myCmd.SelectCommand.Parameters.Add("@RemovePreviousInstance", SqlDbType.Bit);
            myCmd.SelectCommand.Parameters["@LoginID"].Value = strLoginID.Trim();
            myCmd.SelectCommand.Parameters["@Password"].Value = strPassword.Trim();
            myCmd.SelectCommand.Parameters["@InstanceID"].Value = System.Guid.NewGuid();
            myCmd.SelectCommand.Parameters["@IPAddress"].Value = strIPAddress.Trim();
            myCmd.SelectCommand.Parameters["@ClientBrowser"].Value = strClientBrowser.Trim();
            myCmd.SelectCommand.Parameters["@ClientBrowserVersion"].Value = strClientBrowserVersion.Trim();
            myCmd.SelectCommand.Parameters["@UserAgent"].Value = strUserAgent.Trim();
            myCmd.SelectCommand.Parameters["@JSEnabled"].Value = strJSEnabled.Trim();
            myCmd.SelectCommand.Parameters["@RemovePreviousInstance"].Value = blnRemovePreviousInstance;
            conn.Open();
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }
        //Mitul - 13May2013 - Capture Client Browser and OS details and log them - END

        public DataSet ValidateLogin(string strLoginID, string strPassword, string strIPAddress, Boolean blnRemovePreviousInstance)
        {

            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_ValidateLogin", conn.Connection);
            myCmd.SelectCommand.CommandType = CommandType.StoredProcedure;
            myCmd.SelectCommand.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.SelectCommand.Parameters.Add("@Password", SqlDbType.VarChar, 100);
            myCmd.SelectCommand.Parameters.Add("@InstanceID", SqlDbType.UniqueIdentifier);
            myCmd.SelectCommand.Parameters.Add("@IPAddress", SqlDbType.VarChar, 20);
            myCmd.SelectCommand.Parameters.Add("@RemovePreviousInstance", SqlDbType.Bit);
            myCmd.SelectCommand.Parameters["@LoginID"].Value = strLoginID.Trim();
            myCmd.SelectCommand.Parameters["@Password"].Value = strPassword.Trim();
            myCmd.SelectCommand.Parameters["@InstanceID"].Value = System.Guid.NewGuid();
            myCmd.SelectCommand.Parameters["@IPAddress"].Value = strIPAddress.Trim();
            myCmd.SelectCommand.Parameters["@RemovePreviousInstance"].Value = blnRemovePreviousInstance;
            conn.Open();
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }       

        public Boolean UpdateUser(string strLoginID, string strLoginName, string strEmail, string strUserType, string strUserRole, Boolean blnActive, Boolean blnLegacy, Boolean blnBlockLogin, Boolean blnDashboard, string strGOC, Boolean blnOrderManagement, Array arrClientList, string strCreatedBy,
             string strUPN, string strADFSEmail, string strFName, string strLName, string strGUnit, string strLocation, string strTelNo, string strWorkDesc, Boolean blnCancelOrder, Boolean blnMercuryID)
        {
            string strStatus;
            ISBL.General oISBLGen = new ISBL.General();

            if (oISBLGen.ISADFSLoginID(strLoginID)) //If is ADFS Login ID, update ADFS Claims Info
            {
                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_UpdateUserADFS";
                myCmd.CommandType = CommandType.StoredProcedure;
                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                myCmd.Parameters.Add("@LoginName", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@Email", SqlDbType.VarChar, 200);
                myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 30);
                myCmd.Parameters.Add("@Role", SqlDbType.VarChar, 15);
                myCmd.Parameters.Add("@Active", SqlDbType.Bit);
                myCmd.Parameters.Add("@Legacy", SqlDbType.Bit);
                myCmd.Parameters.Add("@Dashboard", SqlDbType.Bit);
                myCmd.Parameters.Add("@GOC", SqlDbType.VarChar, 5);
                myCmd.Parameters.Add("@MercuryID", SqlDbType.Bit); //Mercury Enhancement - Jan 2014
                myCmd.Parameters.Add("@OrderManagement", SqlDbType.Bit);

                myCmd.Parameters.Add("@UPN", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@ADFSEmail", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@FName", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@LName", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@GUnit", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@Location", SqlDbType.VarChar, 250);
                myCmd.Parameters.Add("@TelNo", SqlDbType.VarChar, 50);
                myCmd.Parameters.Add("@WorkDesc", SqlDbType.VarChar, 250);
                myCmd.Parameters.Add("@CancelOrder", SqlDbType.Bit); //Implement Cancel Order Oct 2011

                myCmd.Parameters["@LoginID"].Value = strLoginID;
                myCmd.Parameters["@LoginName"].Value = strLoginName;
                myCmd.Parameters["@Email"].Value = strEmail;
                myCmd.Parameters["@Usertype"].Value = strUserType;
                myCmd.Parameters["@Role"].Value = strUserRole;
                myCmd.Parameters["@Active"].Value = blnActive;
                myCmd.Parameters["@Legacy"].Value = blnLegacy;
                myCmd.Parameters["@Dashboard"].Value = blnDashboard;
                myCmd.Parameters["@GOC"].Value = strGOC;
                myCmd.Parameters["@MercuryID"].Value = blnMercuryID; //Mercury Enhancement - Jan 2014
                myCmd.Parameters["@OrderManagement"].Value = blnOrderManagement;

                myCmd.Parameters["@UPN"].Value = strUPN;
                myCmd.Parameters["@ADFSEmail"].Value = strADFSEmail;
                myCmd.Parameters["@FName"].Value = strFName;
                myCmd.Parameters["@LName"].Value = strLName;
                myCmd.Parameters["@GUnit"].Value = strGUnit;
                myCmd.Parameters["@Location"].Value = strLocation;
                myCmd.Parameters["@TelNo"].Value = strTelNo;
                myCmd.Parameters["@WorkDesc"].Value = strWorkDesc;
                myCmd.Parameters["@CancelOrder"].Value = blnCancelOrder;//Implement Cancel Order Oct 2011

                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateUser.sp_UpdateUserADFS";
                strStatus = conn.cmdScalarStoredProc(myCmd);
                myCmd.Dispose();
                conn.Close();
            }
            else
            {
                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_UpdateUser";
                myCmd.CommandType = CommandType.StoredProcedure;
                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                myCmd.Parameters.Add("@LoginName", SqlDbType.VarChar, 100);
                myCmd.Parameters.Add("@Email", SqlDbType.VarChar, 200);
                myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 30);
                myCmd.Parameters.Add("@Role", SqlDbType.VarChar, 15);
                myCmd.Parameters.Add("@Active", SqlDbType.Bit);
                myCmd.Parameters.Add("@Legacy", SqlDbType.Bit);
                myCmd.Parameters.Add("@BlockLogin", SqlDbType.Bit);
                myCmd.Parameters.Add("@Dashboard", SqlDbType.Bit);
                myCmd.Parameters.Add("@GOC", SqlDbType.VarChar, 5);
                myCmd.Parameters.Add("@MercuryID", SqlDbType.Bit); //Mercury Enhancement - Jan 2014
                myCmd.Parameters.Add("@OrderManagement", SqlDbType.Bit);
                myCmd.Parameters.Add("@CancelOrder", SqlDbType.Bit);//Implement Cancel Order Oct 2011
                myCmd.Parameters["@LoginID"].Value = strLoginID;
                myCmd.Parameters["@LoginName"].Value = strLoginName;
                myCmd.Parameters["@Email"].Value = strEmail;
                myCmd.Parameters["@Usertype"].Value = strUserType;
                myCmd.Parameters["@Role"].Value = strUserRole;
                myCmd.Parameters["@Active"].Value = blnActive;
                myCmd.Parameters["@Legacy"].Value = blnLegacy;
                myCmd.Parameters["@BlockLogin"].Value = blnBlockLogin;
                myCmd.Parameters["@Dashboard"].Value = blnDashboard;
                myCmd.Parameters["@GOC"].Value = strGOC;
                myCmd.Parameters["@MercuryID"].Value = blnMercuryID; //Mercury Enhancement - Jan 2014
                myCmd.Parameters["@OrderManagement"].Value = blnOrderManagement;
                myCmd.Parameters["@CancelOrder"].Value = blnCancelOrder;//Implement Cancel Order Oct 2011
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateUser.sp_UpdateUser";
                strStatus = conn.cmdScalarStoredProc(myCmd);
                myCmd.Dispose();
                conn.Close();
            }

            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    if (strUserType == "BDM")
                    {
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.UpdateUser.sp_RemoveBDMClient";
                        SqlCommand mycmd = new SqlCommand();
                        mycmd.Connection = conn.Connection;
                        mycmd.CommandText = "sp_RemoveBDMClient";
                        mycmd.CommandType = CommandType.StoredProcedure;
                        mycmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                        mycmd.Parameters["@LoginID"].Value = strLoginID;
                        conn.cmdScalarStoredProc(mycmd);
                        mycmd.Dispose();

                        foreach (string str in arrClientList)
                        {
                            SqlCommand cmd = new SqlCommand();
                            cmd.Connection = conn.Connection;
                            cmd.CommandText = "sp_InsertClientCaseManager";
                            cmd.CommandType = CommandType.StoredProcedure;
                            cmd.Parameters.Add("@ClientID", SqlDbType.VarChar, 30);
                            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                            cmd.Parameters["@ClientID"].Value = str;
                            cmd.Parameters["@LoginID"].Value = strLoginID;
                            cmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                            conn.cmdScalarStoredProc(cmd);
                            cmd.Dispose();
                        }
                        conn.Close();
                        return true;
                    }
                    else if (strUserRole == "C4")
                    {
                        // Added by Stev, 21 Oct 08. Delete client codes that associated to C4 Login, Insert new client from the new selected list. 
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.UpdateUser.sp_DeleteC4Client";

                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.Connection = conn.Connection;
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.CommandText = "sp_DeleteC4Client";

                        sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                        sqlCmd.Parameters["@LoginID"].Value = strLoginID;

                        strStatus = conn.cmdScalarStoredProc(sqlCmd);
                        sqlCmd.Dispose();

                        if (strStatus == "1" || strStatus == "True")
                        {
                            conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertC4Client";
                            foreach (string sClCode in arrClientList)
                            {
                                SqlCommand sqlIstCmd = new SqlCommand();
                                sqlIstCmd.Connection = conn.Connection;
                                sqlIstCmd.CommandText = "sp_InsertC4Client";
                                sqlIstCmd.CommandType = CommandType.StoredProcedure;
                                sqlIstCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                                sqlIstCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                                sqlIstCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                                sqlIstCmd.Parameters["@LoginId"].Value = strLoginID;
                                sqlIstCmd.Parameters["@ClCode"].Value = sClCode;
                                sqlIstCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                                conn.cmdScalarStoredProc(sqlIstCmd);
                                sqlIstCmd.Dispose();
                            }
                            conn.Close();
                            return true;
                        }
                        else
                        {
                            conn.Close();
                            return false;
                        }
                    }
                    else if (strUserRole == "C5")
                    {
                        // Added by Adam, 1 July 2010, Delete client codes that associated to C5 Login, Insert new client from the new selected list. 
                        conn.Open();
                        conn.callingMethod = "ISBL.Admin.UpdateUser.sp_DeleteC5Client";

                        SqlCommand sqlCmd = new SqlCommand();
                        sqlCmd.Connection = conn.Connection;
                        sqlCmd.CommandType = CommandType.StoredProcedure;
                        sqlCmd.CommandText = "sp_DeleteC5Client";

                        sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                        sqlCmd.Parameters["@LoginID"].Value = strLoginID;

                        strStatus = conn.cmdScalarStoredProc(sqlCmd);
                        sqlCmd.Dispose();

                        if (strStatus == "1" || strStatus == "True")
                        {
                            conn.callingMethod = "ISBL.Admin.InsertNewLogin.sp_InsertC5Client";
                            foreach (string sClCode in arrClientList)
                            {
                                SqlCommand sqlIstCmd = new SqlCommand();
                                sqlIstCmd.Connection = conn.Connection;
                                sqlIstCmd.CommandText = "sp_InsertC5Client";
                                sqlIstCmd.CommandType = CommandType.StoredProcedure;
                                sqlIstCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                                sqlIstCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                                sqlIstCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                                sqlIstCmd.Parameters["@LoginId"].Value = strLoginID;
                                sqlIstCmd.Parameters["@ClCode"].Value = sClCode;
                                sqlIstCmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                                conn.cmdScalarStoredProc(sqlIstCmd);
                                sqlIstCmd.Dispose();
                            }
                            conn.Close();
                            return true;
                        }
                        else
                        {
                            conn.Close();
                            return false;
                        }
                    }
                    else
                        return true;
                }
                else
                {
                    return false;
                }
            }
            else
            {
                return false;
            }

        }

        /* Start Adam 18 June 2009 Add new Admin menu called Link Client for BDM */
        public Boolean UpdateLinkClient(string strBDMLoginID, Array arrClientList)
        {

            Boolean bResult = true; 
            General objGen = new General();
            
            ISDL.Connect connTran = new ISDL.Connect();
            connTran.setConnection("ocrsConnection");

            connTran.Open();
            connTran.BeginTransaction();

            //Remove Client
            SqlCommand myCmdRemoveClient = new SqlCommand();
            myCmdRemoveClient.Connection = connTran.Connection;
            
            //ReInsert Client
            SqlCommand myCmdReInsertClient = new SqlCommand();
            myCmdReInsertClient.Connection = connTran.Connection;

            try
            {
                //Remove Client

                myCmdRemoveClient.CommandText = "sp_RemoveBDMClient";
                myCmdRemoveClient.CommandType = CommandType.StoredProcedure;
                myCmdRemoveClient.Transaction = connTran.currentTransaction;

                myCmdRemoveClient.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                myCmdRemoveClient.Parameters["@LoginID"].Value = strBDMLoginID;

                if (connTran.cmdNoneQuery(myCmdRemoveClient) == false)
                {
                    bResult = false;
                    connTran.RollBackTransaction();
                    
                }

                if (bResult)
                {
                    myCmdReInsertClient.CommandText = "sp_InsertClientCaseManager";
                    myCmdReInsertClient.CommandType = CommandType.StoredProcedure;
                    myCmdReInsertClient.Transaction = connTran.currentTransaction;

                    myCmdReInsertClient.Parameters.Add("@ClientID", SqlDbType.VarChar, 30);
                    myCmdReInsertClient.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
                    myCmdReInsertClient.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);

                    foreach (string str in arrClientList)
                    {
                        myCmdReInsertClient.Parameters["@ClientID"].Value = str;
                        myCmdReInsertClient.Parameters["@LoginID"].Value = strBDMLoginID;
                        myCmdReInsertClient.Parameters["@CreatedBy"].Value = strBDMLoginID;

                        if (connTran.cmdNoneQuery(myCmdReInsertClient) == false)
                        {
                            bResult = false;
                            connTran.RollBackTransaction();
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                myCmdRemoveClient = null;
                myCmdReInsertClient = null;

                if (bResult)
                {
                    connTran.CommitTransaction();
                }

                connTran.Dispose();
                connTran.Close();
            }
            return bResult;
        }
        /* End Adam 18 June 2009 Add new Admin menu called Link Client for BDM */

        public DataSet GetUserDetails(string strLoginID)
        {
            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_GetUserDetails", conn.Connection);
            myCmd.SelectCommand.CommandType = CommandType.StoredProcedure;
            myCmd.SelectCommand.Parameters.Add("@LoginID", SqlDbType.VarChar, 30);
            myCmd.SelectCommand.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetUserDetails";
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }

        public SqlDataReader GetEventLog(string strEventLog, string strClientCode, string strCRN
           , string strAction, string strStartDateFrom, string strStartDateTo, string strEndDateFrom, string strEndDateTo
           , string strActionBy, string strUserType, string strErrorCode, string strErrorMessage, ref SqlConnection sqlConnect)
        {
            ISBL.General oISBLGen = new ISBL.General();
            string strSql;
            string strwhereclause;
            bool blWhere = false;
            strSql = "SELECT TOP 200 * FROM ocrsEventLog ";
            if (strEventLog.Length > 0)
            {
                strwhereclause = "ID = '" + strEventLog + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strClientCode.Length > 0)
            {
                strwhereclause = "ClientCode = '" + strClientCode + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strCRN.Length > 0)
            {
                strwhereclause = "CRN LIKE '%" + oISBLGen.SafeSqlLiteral(strCRN) + "%' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strAction.Length > 0)
            {
                strwhereclause = "Action = '" + strAction + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if ((strStartDateFrom.Length > 0) && (strStartDateTo.Length > 0))
            {
                strwhereclause = "ActionStartDate BETWEEN '" + strStartDateFrom + "' AND '" + strStartDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strStartDateFrom.Length > 0)
            {
                strwhereclause = "ActionStartDate >= '" + strStartDateFrom + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strStartDateTo.Length > 0)
            {
                strwhereclause = "ActionStartDate <= '" + strStartDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if ((strEndDateFrom.Length > 0) && (strEndDateTo.Length > 0))
            {
                strwhereclause = "ActionEndDate BETWEEN '" + strEndDateFrom + "' AND '" + strEndDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            //else if (strStartDateFrom.Length > 0) //comment wrong logic Adam 8 Jan 09
            else if (strEndDateFrom.Length > 0)
            {
                strwhereclause = "ActionEndDate >= '" + strEndDateFrom + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            // else if (strStartDateTo.Length > 0) //comment wrong logic Adam 8 Jan 09
            else if (strEndDateTo.Length > 0)
            {
                //strwhereclause = "ActionEndDate <= '" + strStartDateTo + "' "; //comment wrong logic Adam 8 Jan 09
                strwhereclause = "ActionEndDate <= '" + strEndDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if (strActionBy.Length > 0)
            {
                strwhereclause = "ActionBy = '" + strActionBy + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if ((strUserType.Length > 0) && (strUserType != "ALL"))
            {
                strwhereclause = "UserType ='" + strUserType + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strErrorCode.Length > 0)
            {
                strwhereclause = "ErrorCode = '" + oISBLGen.SafeSqlLiteral(strErrorCode) + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strErrorMessage.Length > 0)
            {
                strwhereclause = "ErrorMessage LIKE '%" + oISBLGen.SafeSqlLiteral(strErrorMessage) + "%' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            strSql = strSql + " ORDER BY ActionStartDate DESC";

            oISBLGen.DisposeConnection();
            oISBLGen = null;

            myDataReader = conn.cmdReader(strSql, ref sqlConnect);

            return myDataReader;
        }

        private string constructSQL(string strsql, string strwhereclause, bool blWhere)
        {
            string strSQL;
            if (blWhere)
            {
                strSQL = strsql + " AND " + strwhereclause;
            }
            else
            {
                strSQL = strsql + " WHERE " + strwhereclause;
            }
            return strSQL;
        }

        public DataSet GetEventLogAction()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEventLogAction";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEventLogAction";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetEventLogActionBy()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEventLogActionBy";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEventLogActionBy";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetEventLogErrorCode()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEventLogErrorCode";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEventLogErrorCode";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public string GetEventLogDetail(string strLogID, string strDataToGet)
        {
            string strData;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEventLogDetail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LogID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@Data", SqlDbType.VarChar, 10);
            myCmd.Parameters["@LogID"].Value = new System.Guid(strLogID);
            myCmd.Parameters["@Data"].Value = strDataToGet;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEventLogDetail";
            strData = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strData;

        }

        
        public DataSet GetInstanceLoginID()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetInstanceLoginID";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetInstanceLoginID";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public SqlDataReader GetLoginInstance(string strLoginID, string strDateFrom, string strDateTo, ref SqlConnection sqlConnect)
        {
            string strSQL;
            string strwhereclause;
            bool blWhere = false;
            strSQL = "SELECT * FROM ocrsLoginInstance ";

            if (strLoginID.Length > 0)
            {
                strwhereclause = "LoginID = '" + strLoginID + "' ";
                strSQL = constructSQL(strSQL, strwhereclause, blWhere);
                blWhere = true;
            }

            if ((strDateFrom.Length > 0) && (strDateTo.Length > 0))
            {
                strwhereclause = "InstanceDate BETWEEN '" + strDateFrom + "' AND '" + strDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSQL = constructSQL(strSQL, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strDateFrom.Length > 0)
            {
                strwhereclause = "InstanceDate >= '" + strDateFrom + "' ";
                strSQL = constructSQL(strSQL, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strDateTo.Length > 0)
            {
                strwhereclause = "InstanceDate <= '" + strDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSQL = constructSQL(strSQL, strwhereclause, blWhere);
                blWhere = true;
            }

            strSQL = strSQL + "ORDER BY InstanceDate DESC";
            myDataReader = conn.cmdReader(strSQL, ref sqlConnect);

            return myDataReader;

        }

        public DataSet GetClientMaster(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientMaster";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientMaster";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        // Created By: Stev, 30 Sept 2008
        #region NewClient
        // Purpose   : To Update a Client Master by calling Stored Proc with parameters pass in, return boolean to calling page - ClientMaintenance.aspx
        public Boolean UpdateClientBP(string sClCode, string sClName, string sCurrency)
        {
            SqlCommand sqlCmd = new SqlCommand();
            string sStatus;

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_UpdateClientBP";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                sqlCmd.Parameters.Add("@ClName", SqlDbType.VarChar, 100);
                sqlCmd.Parameters.Add("@Currency", SqlDbType.VarChar, 5);


                sqlCmd.Parameters["@ClCode"].Value = sClCode;
                sqlCmd.Parameters["@ClName"].Value = sClName;
                sqlCmd.Parameters["@Currency"].Value = sCurrency;


                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateClientBP";
                sStatus = conn.cmdScalarStoredProc(sqlCmd);
                sqlCmd.Dispose();
                conn.Close();
                if (sStatus == "1" || sStatus == "True")
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }        

        #endregion 

        public Boolean UpdateClientLogo(string prmStrClientCode, string prmStrLogoPath, string prmStrLogoText)
        {
            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateClientLogo";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@LogoPath", SqlDbType.VarChar, 100);
            myCmd.Parameters.Add("@LogoText", SqlDbType.VarChar, 30);

            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@LogoPath"].Value = prmStrLogoPath;
            myCmd.Parameters["@LogoText"].Value = prmStrLogoText;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.UpdateClientLogo";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            myCmd.Dispose();
            conn.Close();
            if (strStatus == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //START Default ReportType Setting Enhancement - Mitul July2013
        public Boolean UpdateClientMaster(string strClientCode, string strEmail, Boolean blnAutoAssignment, Boolean blnBudget, Boolean blnDueDate, Boolean blnREUpdate, Boolean blnBillingIndicator, Boolean blnAddEmail, Boolean blnSingleCRN, Boolean blnBulkOrder, string strLogoPath, string strLogoText, int intShowEmailSubjectName, Boolean blnDefaultReportType, string strDefaultReportType, string strDefaultSubReportType, string strRefreshOrderCriteriaID, Boolean blhEnableRefreshOrderByDefault, string strRefreshOrderAlertEmail, Boolean blnCsublevelBudget)
        {
            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateClientMaster";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@Email", SqlDbType.VarChar, 100);
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters.Add("@Budget", SqlDbType.Bit);
            myCmd.Parameters.Add("@DueDate", SqlDbType.Bit);
            myCmd.Parameters.Add("@REUpdate", SqlDbType.Bit);
            myCmd.Parameters.Add("@BillingIndicator", SqlDbType.Bit);
            myCmd.Parameters.Add("@AddEmail", SqlDbType.Bit);
            myCmd.Parameters.Add("@SingleCRN", SqlDbType.Bit);
            myCmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            myCmd.Parameters.Add("@LogoPath", SqlDbType.VarChar, 100);
            myCmd.Parameters.Add("@LogoText", SqlDbType.VarChar, 50);
            myCmd.Parameters.Add("@ShowEmailSubjectName", SqlDbType.Int);
            myCmd.Parameters.Add("@EnableDefaultReportType", SqlDbType.Bit);
            myCmd.Parameters.Add("@DefaultReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@DefaultSubReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@RefreshOrderCriteriaID", SqlDbType.VarChar, 10);//BI 20 Refresh Order
            myCmd.Parameters.Add("@EnableRefreshOrderByDefault", SqlDbType.Bit);//BI 20 Refresh Order
            myCmd.Parameters.Add("@RefreshOrderAlertEmail", SqlDbType.Text);//BI 20 Refresh Order
            myCmd.Parameters.Add("@CsublevelBudget", SqlDbType.Bit);//BI 20 Refresh Order
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@Email"].Value = strEmail;
            myCmd.Parameters["@AutoAssignment"].Value = blnAutoAssignment;
            myCmd.Parameters["@Budget"].Value = blnBudget;
            myCmd.Parameters["@DueDate"].Value = blnDueDate;
            myCmd.Parameters["@REUpdate"].Value = blnREUpdate;
            myCmd.Parameters["@BillingIndicator"].Value = blnBillingIndicator;
            myCmd.Parameters["@AddEmail"].Value = blnAddEmail;
            myCmd.Parameters["@SingleCRN"].Value = blnSingleCRN;
            myCmd.Parameters["@BulkOrder"].Value = blnBulkOrder;
            myCmd.Parameters["@LogoPath"].Value = strLogoPath;
            myCmd.Parameters["@LogoText"].Value = strLogoText;
            myCmd.Parameters["@ShowEmailSubjectName"].Value = intShowEmailSubjectName;
            myCmd.Parameters["@EnableDefaultReportType"].Value = blnDefaultReportType;
            myCmd.Parameters["@DefaultReportType"].Value = strDefaultReportType;
            myCmd.Parameters["@DefaultSubReportType"].Value = strDefaultSubReportType;
            myCmd.Parameters["@RefreshOrderCriteriaID"].Value = strRefreshOrderCriteriaID; //BI 20 Refresh Order
            myCmd.Parameters["@EnableRefreshOrderByDefault"].Value = blhEnableRefreshOrderByDefault;//BI 20 Refresh Order
            myCmd.Parameters["@RefreshOrderAlertEmail"].Value = strRefreshOrderAlertEmail; //BI 20 Refresh Order
            myCmd.Parameters["@CsublevelBudget"].Value = blnCsublevelBudget; //BI 20 Refresh Order
            conn.Open();
            conn.callingMethod = "ISBL.Admin.UpdateClientMaster";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            myCmd.Dispose();
            conn.Close();
            if (strStatus == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public bool CheckClientSetupStatusOfAStage(string prmStrClientCode, int stage)
        {
            bool bIsCompleted = true;

            string strStatus = GetClientSetupSectionStatus(prmStrClientCode);

            if (strStatus.Equals(""))
            {
                bIsCompleted = false;
            }
            else
            {
                string cStatus = strStatus.Substring(stage, 1);
                if (cStatus.Equals("0"))
                {
                    bIsCompleted = false;
                }
            }
            return bIsCompleted;
        }

        public Boolean GetClientSubjectRESetupForDefaultReportType(string prmStrClientCode)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckClientSubjectRESetupForDefaultReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientSubjectRESetupForDefaultReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }

        public Boolean GetClientSubjectRESetupForDefaultSubReportType(string prmStrClientCode)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckClientSubjectRESetupForDefaultSubReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientSubjectRESetupForDefaultSubReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }

        public Boolean CheckClientREExistForReportType(string prmStrClientCode, string prmStrReportType)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckClientREExistForReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckClientREExistForReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }

        public Boolean CheckIfClientDefaultReportExists(string prmStrClientCode)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckIfClientDefaultReportExists";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckIfClientDefaultReportExists";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }
        public DataSet GetClientDefaultReportTypes(string prmStrClientCode)
        {
            DataSet dsReportType = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDefaultReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientDefaultReportTypes";
            dsReportType = conn.FillDataSet(sda);
            conn.Close();

            myCmd.Dispose();
            sda.Dispose();
            return dsReportType;
        }
        public DataSet GetClientDefaultReportTypesnew(string prmReportType, string prmStrClientCode)
        {
            DataSet dsReportType = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDefaultReportTypenew";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportType"].Value = prmReportType;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientDefaultReportTypesnew";
            dsReportType = conn.FillDataSet(sda);
            conn.Close();

            myCmd.Dispose();
            sda.Dispose();
            return dsReportType;
        }
        public DataSet GetClientDefaultSubReportType(string prmStrClientCode,string prmReportType,string prmSubReportType,string fAction)
        {
            DataSet dsReportType = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDefaultSubReportTypenew";
            myCmd.CommandType = CommandType.StoredProcedure;
            
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportTypeCode"].Value = prmReportType;
            myCmd.Parameters.Add("@SubreportTypeCode", SqlDbType.VarChar, 15);
            myCmd.Parameters["@SubreportTypeCode"].Value = prmSubReportType;
            myCmd.Parameters.Add("@Faction", SqlDbType.VarChar, 5);
            myCmd.Parameters["@Faction"].Value = fAction;
            
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientDefaultSubReportType";
            dsReportType = conn.FillDataSet(sda);
            conn.Close();

            myCmd.Dispose();
            sda.Dispose();
            return dsReportType;
        }

        //added by sanjeeva for JPMC Changes start..........................................................//
        public DataSet GetClientDefaultSubReportTypes(string prmStrClientCode)
        {
            DataSet dsReportType = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDefaultSubReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientDefaultSubReportTypes";
            dsReportType = conn.FillDataSet(sda);
            conn.Close();

            myCmd.Dispose();
            sda.Dispose();
            return dsReportType;
        }
        //end of jpmc chnages 

        public DataSet GetAdminReportTypeMasterWOSubReportTypes(string prmStrClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetAdminReportTypeMasterWOSubReportTypes";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetAdminReportTypeMasterWOSubReportTypes";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetAdminCompletedSubReportTypes(string prmStrClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetAdminCompletedSubReportTypes";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetAdminCompletedSubReportTypes";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        //END Default ReportType Setting Enhancement - Mitul July2013
       
        public DataSet GetClientTopCountry(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientTopCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientTopCountry";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetCountryExcludeTop(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryExcludeTop";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCountryExcludeTop";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetCountryListExcluded(string strCountry)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryListExcluded";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@prmCountry", SqlDbType.VarChar, 4000);
            myCmd.Parameters["@prmCountry"].Value = strCountry;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCountryListExcluded";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public Boolean UpdateTopCountry(string strClientCode, Array arrCountryList, string strCreatedBy)
        {
            try
            {
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateTopCountry";
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_RemoveClientTopCountry";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                conn.cmdScalarStoredProc(mycmd);

                foreach (string str in arrCountryList)
                {
                    SqlCommand cmd = new SqlCommand();
                    cmd.Connection = conn.Connection;
                    cmd.CommandText = "sp_UpdateClientTopCountry";
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                    //cmd.Parameters.Add("@Country", SqlDbType.VarChar, 30); Jul 2013 Adam - ISIS Atlas Data Sync
                    cmd.Parameters.Add("@Country", SqlDbType.VarChar, 4); //Jul 2013 Adam - ISIS Atlas Data Sync
                    cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                    cmd.Parameters["@ClientCode"].Value = strClientCode;
                    cmd.Parameters["@Country"].Value = str;
                    cmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                    conn.cmdScalarStoredProc(cmd);
                }
                conn.Close();
                return true;
            }
            catch 
            {
                return false;
            }
        }

        public DataSet GetMasterRE()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetMasterRE";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetMasterRE";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetReportTypeMasterAdmin()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportTypeMasterAdmin";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetReportTypeMaster";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetReportTypeMasterAdminsubrep()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportTypeMasterAdminsubrep";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetReportTypeMasterAdminsubrep";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        // added by snajeeva for JPMC changes start-------------------------------------------------//
        public DataSet GetSubReportTypeMasterAdmin(string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetSubReportTypeMasterAdmin";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetSubReportTypeMasterAdmin";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetSubReportTypeMasterMapping(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetSubReportTypeMasterMapping";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetSubReportTypeMasterMapping";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        //public DataSet GetSubReportTypeMasterAdmin(string strReportType)
        //{
        //    DataSet ds = new DataSet();
        //    SqlCommand myCmd = new SqlCommand();
        //    myCmd.Connection = conn.Connection;
        //    myCmd.CommandText = "sp_GetSubReportTypeMasterAdmin";
        //    myCmd.CommandType = CommandType.StoredProcedure;
        //    myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
        //    myCmd.Parameters["@ReportTypeCode"].Value = strReportType;
        //    SqlDataAdapter sda = new SqlDataAdapter();
        //    sda.SelectCommand = myCmd;
        //    conn.Open();
        //    conn.callingMethod = "ISBL.Admin.GetSubReportTypeMasterAdmin";
        //    ds = conn.FillDataSet(sda);
        //    conn.Close();
        //    myCmd.Dispose();
        //    sda.Dispose();
        //    return ds;
        //}
        public DataSet GetSubReportTypeMasterCheck(string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetSubReportTypeMasterCheck";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetSubReportTypeMasterCheck";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        //end of jpmc changes     


        // added by Deepak for JPMC changes start-------------------------------------------------//
        public DataSet GetEntitySubReportTypeCountry(int flag, string selected, string reporttypecode, string Clientcode, int EntityType = 0)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "GetEntitySubReportTypeCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Flag", SqlDbType.Int, 30);
            myCmd.Parameters["@Flag"].Value = flag;

            myCmd.Parameters.Add("@Selected", SqlDbType.VarChar, 30);
            myCmd.Parameters["@Selected"].Value = selected;

            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = reporttypecode;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = Clientcode;
            // @EntityType

            myCmd.Parameters.Add("@EntityType", SqlDbType.Int);
            myCmd.Parameters["@EntityType"].Value = EntityType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEntitySubReportTypeCountry";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        // place order

        public DataSet GetEntitySubRepCountryPlaceOrder(int flag, string selected, string reporttypecode, string Clientcode, int EntityType = 0)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "SP_GetEntitySubRepCountryPlaceOrder";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Flag", SqlDbType.Int, 30);
            myCmd.Parameters["@Flag"].Value = flag;

            myCmd.Parameters.Add("@Selected", SqlDbType.VarChar, 30);
            myCmd.Parameters["@Selected"].Value = selected;

            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = reporttypecode;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = Clientcode;
            // @EntityType

            myCmd.Parameters.Add("@EntityType", SqlDbType.Int);
            myCmd.Parameters["@EntityType"].Value = EntityType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEntitySubRepCountryPlaceOrder";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetEntitySubReportTypeCountryCaseLvl(int flag, string selected, string reporttypecode, string Clientcode, int EntityType = 0)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "GetEntitySubReportTypeCountryCaseLvl";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Flag", SqlDbType.Int, 30);
            myCmd.Parameters["@Flag"].Value = flag;

            myCmd.Parameters.Add("@Selected", SqlDbType.VarChar, 30);
            myCmd.Parameters["@Selected"].Value = selected;

            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = reporttypecode;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = Clientcode;
            // @EntityType

            myCmd.Parameters.Add("@EntityType", SqlDbType.Int);
            myCmd.Parameters["@EntityType"].Value = EntityType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetEntitySubReportTypeCountry";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

       // BindBudgetIndividualCompany
        public DataSet BindBudgetIndividualCompany(int EntityType, string SubReportType, string reporttypecode, string Clientcode,string Currency)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "BindBudgetIndividualCompany";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@EntityType", SqlDbType.Int, 30);
            myCmd.Parameters["@EntityType"].Value = EntityType;

            myCmd.Parameters.Add("@subReportType", SqlDbType.VarChar, 30);
            myCmd.Parameters["@subReportType"].Value = SubReportType;

            myCmd.Parameters.Add("@reportType", SqlDbType.VarChar, 30);
            myCmd.Parameters["@reportType"].Value = reporttypecode;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = Clientcode;

            myCmd.Parameters.Add("@Currenty", SqlDbType.VarChar, 30);
            myCmd.Parameters["@Currenty"].Value = Currency;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.BindBudgetIndividualCompany";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //ended by Deepak

        public DataSet GetReportTypeMaster()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportTypeMaster";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetReportTypeMaster";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRECode(string strREDesc)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRECode";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@REDesc", SqlDbType.VarChar, 500);
            myCmd.Parameters["@REDesc"].Value = strREDesc;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetRECode";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetClientRE(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientRE";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientRE";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean UpdateClientRE(string strClientCode, string strReportType,DataTable tblRE, string strCreatedBy)
        {
            try
            {
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateClientRE";
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_RemoveClientReportRE";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                conn.cmdScalarStoredProc(mycmd);

                if (!(tblRE == null))
                {
                    foreach (DataRow tRow in tblRE.Rows)
                    {
                        SqlCommand cmd = new SqlCommand();
                        cmd.Connection = conn.Connection;
                        cmd.CommandText = "sp_UpdateClientReportRE";
                        cmd.CommandType = CommandType.StoredProcedure;
                        cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);
                        // cmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 4);  Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        cmd.Parameters.Add("@SubjectType", SqlDbType.TinyInt); // Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        cmd.Parameters.Add("@ResearchElementCode", SqlDbType.VarChar, 15);
                        cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 50);
                        cmd.Parameters["@ClientCode"].Value = strClientCode;
                        cmd.Parameters["@ReportType"].Value = strReportType;
                        //cmd.Parameters["@SubjectType"].Value = tRow["SubjectType"]; Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        cmd.Parameters["@SubjectType"].Value = Convert.ToInt16(tRow["SubjectType"].ToString()); //Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        cmd.Parameters["@ResearchElementCode"].Value = tRow["RECode"];
                        cmd.Parameters["@CreatedBy"].Value= strCreatedBy;
                        conn.cmdScalarStoredProc(cmd);
                        
                    }
                }
                conn.Close();
                return true;
            }
            catch 
            {
                return false;
            }
        }

        public DataSet CheckReportTypeExist(string strClientCode, string strReportType, string intCalculation,string strSubReportType) //BI 34 Budget recalculation
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckReportTypeExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);//BI 34 Budget recalculation
            myCmd.Parameters.Add("@Calculation", SqlDbType.Int);
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 30);

            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;            
             myCmd.Parameters["@Calculation"].Value = intCalculation;//BI 34 Budget recalculation
             myCmd.Parameters["@SubReportType"].Value = strSubReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckReportTypeExist";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean CheckClientREExist(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckClientREExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckClientREExist";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();

            if (ds.Tables[0].Rows.Count > 0)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public string DeleteClientReportTypeBudget(string prmStrClientCode, string prmStrReportType)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteClientReportTypeBudget";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.DeleteClientReportTypeBudget";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }

        public DataSet GetDefaultRE(string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetDefaultRE";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 50);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetDefaultRE";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public void InsertClientSetupSectionStatus(string prmStrClientCode)
        {

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_InsertClientSetupSectionStatus";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.InsertClientSetupSectionStatus";
            conn.cmdNoneQuery(myCmd);
            conn.Close();

            myCmd.Dispose();



        }

        /*@@@ Start By Eric 20080423 @@@*/
        public string GetClientSetupSectionStatus(string prmStrClientCode)
        {
            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientSetupSectionStatus";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientSetupSectionStatus";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }
        private void UpdateClientSetupSectionStatus(string prmStrClientCode, string prmStrStatus)
        {
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateClientSetupSectionStatus";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ClientSetupSectionStatus", SqlDbType.VarChar);
            myCmd.Parameters["@ClientSetupSectionStatus"].Value = prmStrStatus;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.UpdateClientSetupSectionStatus";
            conn.cmdNoneQuery(myCmd);
            conn.Close();

            myCmd.Dispose();

        }
        public bool IsClientSetupCompleted(string prmStrClientCode)
        {
            bool bIsCompleted = true;

            string strStatus = GetClientSetupSectionStatus(prmStrClientCode);

            if (strStatus.Equals(""))
            {
                bIsCompleted = false;
            }
            else
            {
                for (int i = 0; i < strStatus.Length; i++)
                {
                    string cStatus = strStatus.Substring(i, 1);
                    if (cStatus.Equals("0"))
                    {
                        bIsCompleted = false;
                        break;
                    }
                }
            }

            return bIsCompleted;

        }
        public string GetLatestClientSetupIncompleteSection(string prmStrClientCode)
        {

            string strSection = "";
            string strStatus = GetClientSetupSectionStatus(prmStrClientCode);

            if (strStatus.Equals(""))
            {
                strSection = "1";
            }
            else
            {
                for (int i = 0; i < strStatus.Length; i++)
                {
                    string cStatus = strStatus.Substring(i, 1);
                    if (cStatus.Equals("0"))
                    {
                        strSection = (i + 1).ToString();
                        break;
                    }
                }
            }

            return strSection;

        }
        public void UpdateClientSetupSectionStatus(string prmStrClientCode, string prmStrSection, string prmStrStatus)
        {

            string strNewStatus = "";
            string strStatus = GetClientSetupSectionStatus(prmStrClientCode);

            for (int i = 0; i < strStatus.Length; i++)
            {
                string cStatus = strStatus.Substring(i, 1);

                if (i == int.Parse(prmStrSection.ToString()) - 1)
                {
                    strNewStatus = strNewStatus + prmStrStatus;
                }
                else
                {
                    strNewStatus = strNewStatus + cStatus;
                }
            }

            UpdateClientSetupSectionStatus(prmStrClientCode, strNewStatus);

        }
        /*@@@ End By Eric 20080423 @@@*/


        /*@@@ Start By Eric 20080506 @@@*/
        public DataSet GetClientReportTypeList(string prmStrClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientReportTypeList";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportTypeList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetOfficeAssignmentList(string prmStrAssignmentType, string prmStrAutoAssignment)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOfficeAssignmentList";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@AssignmentType", SqlDbType.VarChar, 50);
            myCmd.Parameters["@AssignmentType"].Value = prmStrAssignmentType;
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.VarChar, 50);
            myCmd.Parameters["@AutoAssignment"].Value = prmStrAutoAssignment;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetOfficeAssignmentList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public string GetOfficeAssignment(string prmStrClientCode, string prmStrReportType, bool prmBAutoAssignment, bool prmBBulkOrder)
        {

            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOfficeAssignment";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@AutoAssignment"].Value = prmBAutoAssignment;

            myCmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            myCmd.Parameters["@BulkOrder"].Value = prmBBulkOrder;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetOfficeAssignment";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }
        public DataSet GetClientBranchSetupDetail(string prmStrClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBranchSetupDetail";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientBranchSetupDetail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetClientBranchSetupDetailbranch(string prmStrClientCode,string prmStrReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBranchSetupDetailbranch";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportTypeCode"].Value = prmStrReportType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientBranchSetupDetailbranch";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetBranchList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBranchList";
            myCmd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetBranchList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public string GetBranchEmail(string prmStrBranch)
        {
            string strReturn;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBranchEmail";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@Branch", SqlDbType.VarChar);
            myCmd.Parameters["@Branch"].Value = prmStrBranch;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetBranchEmail";
            strReturn = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strReturn;

        }
        public DataSet GetBranchTaxCodeList(string prmStrBranch)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBranchTaxCodeList";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@Branch", SqlDbType.VarChar, 10);
            myCmd.Parameters["@Branch"].Value = prmStrBranch;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetBranchTaxCodeList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public string InsertClientBranch(string prmStrClientCode, string prmStrReportType, string prmStrBranch, bool prmBAutoAssignment, bool prmBBulkorder, string prmStrTaxCode, string prmStrOfficeAssignment, string prmStrAssignmentType, string prmStrCreatedBy)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_InsertClientBranch";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            //myCmd.Parameters.Add("@Branch", SqlDbType.VarChar, 10); Jul 2013 Adam ISIS - Atlas Data Synch reduce to 3
            myCmd.Parameters.Add("@Branch", SqlDbType.VarChar, 3); //Jul 2013 Adam ISIS - Atlas Data Synch reduce to 3
            myCmd.Parameters["@Branch"].Value = prmStrBranch;
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@AutoAssignment"].Value = prmBAutoAssignment;
            myCmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            myCmd.Parameters["@BulkOrder"].Value = prmBBulkorder;
            myCmd.Parameters.Add("@TaxCode", SqlDbType.VarChar, 10);
            myCmd.Parameters["@TaxCode"].Value = prmStrTaxCode;
            myCmd.Parameters.Add("@OfficeAssignment", SqlDbType.VarChar, 20);
            myCmd.Parameters["@OfficeAssignment"].Value = prmStrOfficeAssignment;
            myCmd.Parameters.Add("@AssignmentType", SqlDbType.VarChar, 50);
            myCmd.Parameters["@AssignmentType"].Value = prmStrAssignmentType;
            myCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            myCmd.Parameters["@CreatedBy"].Value = prmStrCreatedBy;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.InsertClientBranch";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;


        }
        
        public string DeleteClientBranch(string prmStrClientCode, string prmStrReportType)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteClientBranch";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.DeleteClientBranch";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }
        public bool IsClientBranchReportTypeCompleted(string prmStrClientCode, string prmStrReportType)
        {

            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsCLientBranchReportTypeCompleted";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.IsClientBranchReportTypeCompleted";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            if (strStatus.Equals("True") || strStatus.Equals("1"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public DataSet GetCaseManagerMasterList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCaseManagerMasterList";
            myCmd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCaseManagerMasterList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetClientReportTypeBudgetSetupDetail(string prmStrClientCode, string prmStrReportType, int prmIntCalculation) //BI 34 Budget recalculation
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientReportTypeBudgetSetupDetail";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            myCmd.Parameters.Add("@Calculation", SqlDbType.Int); //BI 34 Budget recalculation
            myCmd.Parameters["@Calculation"].Value = prmIntCalculation; //BI 34 Budget recalculation
            //myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            //myCmd.Parameters["@SubReportType"].Value = prmStrSubReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportTypeBudgetSetupDetail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public bool IsClientReportTypeBudgetSetupCompleted(string prmStrClientCode)
        {

            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "[sp_IsClientReportTypeBudgetSetupCompleted]";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.IsClientReportTypeBudgetSetupCompleted";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            if (strStatus.Equals("True") || strStatus.Equals("1"))
            {
                return true;
            }
            else
            {
                return false;
            }

        }
        public Boolean UpdateClientReportTypeBudgetSetupnew(string prmStrClientCode, string prmStrReportType)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_UpdateClientReportTypeBudgetNew";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters["@ClientCode"].Value = prmStrClientCode;
                mycmd.Parameters["@ReportType"].Value = prmStrReportType;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateClientReportTypeBudgetNew";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }
        public bool InsertClientReportTypeBudgetSetup(string prmStrClientCode, string prmStrReportType, float prmFBasePrice, int prmIntCompanyMaxPrice, int prmIntIndividualMaxPrice, float prmFCompanyCost, float prmFIndividualCost, string prmStrCaseManager, string prmStrCreatedBy, string prmStrCurrency, string strNoIndividual, string strNoCompany, float prmFBasePriceExpress, int prmIntCompanyMaxPriceExpress, int prmIntIndividualMaxPriceExpress, float prmFCompanyCostExpress, float prmFIndividualCostExpress, int prmTATNormal, int prmTATExpress, int prmTATNormalBulk)
        {

            bool bResult = true;
            string strResultBudget = "";
            string strResultCaseManager = "";
            string strIsBudgetExists = "";
            string strIsCaseManagerExists = "";
            string strResultCurrency = ""; //Adam 6-Feb-09
            bool bAutoAddSubject = false; //Adam 24-Mar-09 Auto Add Subject

            ISDL.Connect connInsertClientReportTypeBudget = new ISDL.Connect();
            connInsertClientReportTypeBudget.setConnection("ocrsConnection");

            connInsertClientReportTypeBudget.Open();
            connInsertClientReportTypeBudget.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudget.Connection;

            SqlCommand myCmdInsertCaseManager = new SqlCommand();
            myCmdInsertCaseManager.Connection = connInsertClientReportTypeBudget.Connection;

            SqlCommand myCmdIsBudgetExist = new SqlCommand();
            myCmdIsBudgetExist.Connection = connInsertClientReportTypeBudget.Connection;

            SqlCommand myCmdIsCaseManagerExist = new SqlCommand();
            myCmdIsCaseManagerExist.Connection = connInsertClientReportTypeBudget.Connection;

            SqlCommand myCmdUpdateBudget = new SqlCommand();
            myCmdUpdateBudget.Connection = connInsertClientReportTypeBudget.Connection;

            SqlCommand myCmdUpdateCaseManager = new SqlCommand();
            myCmdUpdateCaseManager.Connection = connInsertClientReportTypeBudget.Connection;

            //Adam 6-Feb-09 Update Client Currency
            SqlCommand myCmdUpdateClientCurrency = new SqlCommand();
            myCmdUpdateClientCurrency.Connection = connInsertClientReportTypeBudget.Connection;

            //Adam 24-Mar-09 Auto Add Subject
            SqlCommand myCmdUpdateAutoAddSubject = new SqlCommand();
            myCmdUpdateAutoAddSubject.Connection = connInsertClientReportTypeBudget.Connection;

            try
            {

                //Adam 6-Feb-09 Update Client Currency in ocrsClientMaster

                //@@@ Start : Update Client Currency
                myCmdUpdateClientCurrency.CommandText = "sp_UpdateClientCurrency";
                myCmdUpdateClientCurrency.CommandType = CommandType.StoredProcedure;
                myCmdUpdateClientCurrency.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                myCmdUpdateClientCurrency.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                myCmdUpdateClientCurrency.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
                myCmdUpdateClientCurrency.Parameters["@ClientCode"].Value = prmStrClientCode;
                myCmdUpdateClientCurrency.Parameters["@Currency"].Value = prmStrCurrency;

                strResultCurrency = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdUpdateClientCurrency);

                if (strResultCurrency.Equals("") || strResultCurrency.Equals("False") || strResultCurrency.Equals("0"))
                {
                    bResult = false;
                    connInsertClientReportTypeBudget.RollBackTransaction();
                }
                //@@@ End : Update Client Currency

                if (bResult)  //Adam 6-Feb-09
                {
                    //@@@ Start : Is Budget Exist
                    myCmdIsBudgetExist.CommandText = "sp_IsClientReportTypeBudgetExists";
                    myCmdIsBudgetExist.CommandType = CommandType.StoredProcedure;
                    myCmdIsBudgetExist.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                    myCmdIsBudgetExist.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                    myCmdIsBudgetExist.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                   // myCmdIsBudgetExist.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);

                    //myCmdIsBudgetExist.Parameters.Add("@pEntityType", SqlDbType.VarChar, 30);
                    myCmdIsBudgetExist.Parameters.Add("@pCurrency", SqlDbType.VarChar, 30);
                   // myCmdIsBudgetExist.Parameters.Add("@pCountry", SqlDbType.VarChar, 4000);
                    //myCmdIsBudgetExist.Parameters.Add("@pAction", SqlDbType.VarChar, 30);

                    myCmdIsBudgetExist.Parameters["@ClientCode"].Value = prmStrClientCode;
                    myCmdIsBudgetExist.Parameters["@ReportType"].Value = prmStrReportType;
                    //myCmdIsBudgetExist.Parameters["@SubReportType"].Value = prmSubreporttype;
                    //myCmdIsBudgetExist.Parameters["@pEntityType"].Value = prmEntityType;
                    myCmdIsBudgetExist.Parameters["@pCurrency"].Value = prmStrCurrency;
                   // myCmdIsBudgetExist.Parameters["@pCountry"].Value = prmCountry;
                    //myCmdIsBudgetExist.Parameters["@pAction"].Value = prmAction;

                    strIsBudgetExists = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdIsBudgetExist);

                    if (strIsBudgetExists.Equals("True") || strIsBudgetExists.Equals("1"))
                    { //Update
                        //@@@ Start : Update Budget
                        myCmdUpdateBudget.CommandText = "sp_UpdateClientReportTypeBudget";
                        myCmdUpdateBudget.CommandType = CommandType.StoredProcedure;
                        myCmdUpdateBudget.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                        myCmdUpdateBudget.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        myCmdUpdateBudget.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                        //myCmdUpdateBudget.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
                        myCmdUpdateBudget.Parameters.Add("@BasePrice", SqlDbType.Float);
                        myCmdUpdateBudget.Parameters.Add("@MaxCompanyForBasePrice", SqlDbType.Int);
                        myCmdUpdateBudget.Parameters.Add("@MaxIndividualForBasePrice", SqlDbType.Int);
                        myCmdUpdateBudget.Parameters.Add("@AdditionalCostPerCompany", SqlDbType.Float);
                        myCmdUpdateBudget.Parameters.Add("@AdditionalCostPerIndividual", SqlDbType.Float);
                        //Start Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdUpdateBudget.Parameters.Add("@BasePriceExpress", SqlDbType.Float);
                        myCmdUpdateBudget.Parameters.Add("@MaxCompanyForBasePriceExpress", SqlDbType.Int);
                        myCmdUpdateBudget.Parameters.Add("@MaxIndividualForBasePriceExpress", SqlDbType.Int);
                        myCmdUpdateBudget.Parameters.Add("@AdditionalCostPerCompanyExpress", SqlDbType.Float);
                        myCmdUpdateBudget.Parameters.Add("@AdditionalCostPerIndividualExpress", SqlDbType.Float);
                        myCmdUpdateBudget.Parameters.Add("@TATNormal", SqlDbType.Int);
                        myCmdUpdateBudget.Parameters.Add("@TATExpress", SqlDbType.Int);
                        //End Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdUpdateBudget.Parameters.Add("@TATNormalBulk", SqlDbType.Int); //ISIS v2 Phase 1 Release 3
                        // jpmc changes added by sanjeeva start
                        //myCmdUpdateBudget.Parameters.Add("@pEntityType", SqlDbType.VarChar, 30);
                        myCmdUpdateBudget.Parameters.Add("@pCurrency", SqlDbType.VarChar, 30);
                        //myCmdUpdateBudget.Parameters.Add("@pCountry", SqlDbType.VarChar, 4000);
                        //myCmdUpdateBudget.Parameters.Add("@pAction", SqlDbType.VarChar, 30);
                        // end of jpmc 
                        myCmdUpdateBudget.Parameters["@ClientCode"].Value = prmStrClientCode;
                        myCmdUpdateBudget.Parameters["@ReportType"].Value = prmStrReportType;
                        //myCmdUpdateBudget.Parameters["@SubReportType"].Value = prmSubreporttype;
                        myCmdUpdateBudget.Parameters["@BasePrice"].Value = prmFBasePrice;
                        myCmdUpdateBudget.Parameters["@MaxCompanyForBasePrice"].Value = prmIntCompanyMaxPrice;
                        myCmdUpdateBudget.Parameters["@MaxIndividualForBasePrice"].Value = prmIntIndividualMaxPrice;
                        myCmdUpdateBudget.Parameters["@AdditionalCostPerCompany"].Value = prmFCompanyCost;
                        myCmdUpdateBudget.Parameters["@AdditionalCostPerIndividual"].Value = prmFIndividualCost;
                        //Start Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdUpdateBudget.Parameters["@BasePriceExpress"].Value = prmFBasePriceExpress;
                        myCmdUpdateBudget.Parameters["@MaxCompanyForBasePriceExpress"].Value = prmIntCompanyMaxPriceExpress;
                        myCmdUpdateBudget.Parameters["@MaxIndividualForBasePriceExpress"].Value = prmIntIndividualMaxPriceExpress;
                        myCmdUpdateBudget.Parameters["@AdditionalCostPerCompanyExpress"].Value = prmFCompanyCostExpress;
                        myCmdUpdateBudget.Parameters["@AdditionalCostPerIndividualExpress"].Value = prmFIndividualCostExpress;
                        myCmdUpdateBudget.Parameters["@TATNormal"].Value = prmTATNormal;
                        myCmdUpdateBudget.Parameters["@TATExpress"].Value = prmTATExpress;
                        //End Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdUpdateBudget.Parameters["@TATNormalBulk"].Value = prmTATNormalBulk; //ISIS v2 Phase 1 Release 3
                        // jpmc changes added by sanjeeva start
                        //myCmdUpdateBudget.Parameters["@pEntityType"].Value = prmEntityType;
                        myCmdUpdateBudget.Parameters["@pCurrency"].Value = prmStrCurrency;
                       // myCmdUpdateBudget.Parameters["@pCountry"].Value = prmCountry;
                        //myCmdUpdateBudget.Parameters["@pAction"].Value = prmAction;
                        // end of jpmc 

                        strResultBudget = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdUpdateBudget);
                        if (strResultBudget.Equals("") || strResultBudget.Equals("False") || strResultBudget.Equals("0"))
                        {
                            bResult = false;
                            connInsertClientReportTypeBudget.RollBackTransaction();
                        }
                        //@@@ End : Update Budget
                    }
                    else //Insert
                    {
                        //@@@ Start : Insert Budget
                        myCmdInsertBudget.CommandText = "sp_InsertClientReportTypeBudget";
                        myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                        myCmdInsertBudget.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                        myCmdInsertBudget.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        myCmdInsertBudget.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                        myCmdInsertBudget.Parameters.Add("@BasePrice", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@MaxCompanyForBasePrice", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@MaxIndividualForBasePrice", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerCompany", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerIndividual", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                        //Start Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdInsertBudget.Parameters.Add("@BasePriceExpress", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@MaxCompanyForBasePriceExpress", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@MaxIndividualForBasePriceExpress", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerCompanyExpress", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerIndividualExpress", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@TATNormal", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@TATExpress", SqlDbType.Int);
                        //End Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdInsertBudget.Parameters.Add("@TATNormalBulk", SqlDbType.Int); //ISIS v2 Phase 1 Release 3
                        myCmdInsertBudget.Parameters.Add("@pCurrency", SqlDbType.VarChar, 30);
                        //myCmdInsertBudget.Parameters.Add("@pCountry", SqlDbType.VarChar, 30);
                        //myCmdInsertBudget.Parameters.Add("@psubreporttype", SqlDbType.VarChar, 15);
                        
                        myCmdInsertBudget.Parameters["@ClientCode"].Value = prmStrClientCode;
                        myCmdInsertBudget.Parameters["@ReportType"].Value = prmStrReportType;
                        myCmdInsertBudget.Parameters["@BasePrice"].Value = prmFBasePrice;
                        myCmdInsertBudget.Parameters["@MaxCompanyForBasePrice"].Value = prmIntCompanyMaxPrice;
                        myCmdInsertBudget.Parameters["@MaxIndividualForBasePrice"].Value = prmIntIndividualMaxPrice;
                        myCmdInsertBudget.Parameters["@AdditionalCostPerCompany"].Value = prmFCompanyCost;
                        myCmdInsertBudget.Parameters["@AdditionalCostPerIndividual"].Value = prmFIndividualCost;
                        myCmdInsertBudget.Parameters["@CreatedBy"].Value = prmStrCreatedBy;
                        //Start Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdInsertBudget.Parameters["@BasePriceExpress"].Value = prmFBasePriceExpress;
                        myCmdInsertBudget.Parameters["@MaxCompanyForBasePriceExpress"].Value = prmIntCompanyMaxPriceExpress;
                        myCmdInsertBudget.Parameters["@MaxIndividualForBasePriceExpress"].Value = prmIntIndividualMaxPriceExpress;
                        myCmdInsertBudget.Parameters["@AdditionalCostPerCompanyExpress"].Value = prmFCompanyCostExpress;
                        myCmdInsertBudget.Parameters["@AdditionalCostPerIndividualExpress"].Value = prmFIndividualCostExpress;
                        myCmdInsertBudget.Parameters["@TATNormal"].Value = prmTATNormal;
                        myCmdInsertBudget.Parameters["@TATExpress"].Value = prmTATExpress;
                        //End Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
                        myCmdInsertBudget.Parameters["@TATNormalBulk"].Value = prmTATNormalBulk; //ISIS v2 Phase 1 Release 3
                        myCmdInsertBudget.Parameters["@pCurrency"].Value = prmStrCurrency;
                       //myCmdInsertBudget.Parameters["@pCountry"].Value = prmCountry;
                        //myCmdInsertBudget.Parameters["@psubreporttype"].Value = prmSubreporttype;
                        
                        strResultBudget = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdInsertBudget);
                        if (strResultBudget.Equals("") || strResultBudget.Equals("False") || strResultBudget.Equals("0"))
                        {
                            bResult = false;
                            connInsertClientReportTypeBudget.RollBackTransaction();
                        }
                        //@@@ End : Insert Budget
                    }
                    //@@@ END : Is Budget Exist
                }

                // Code Commented By Deepak to not to save case manager from Step 5
                //if (bResult)  //Adam 6-Feb-09
                //{
                //    //@@@ Start : Is Case manage Exists
                //    myCmdIsCaseManagerExist.CommandText = "sp_IsClientCaseManagerExists";
                //    myCmdIsCaseManagerExist.CommandType = CommandType.StoredProcedure;
                //    myCmdIsCaseManagerExist.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                //    myCmdIsCaseManagerExist.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                //    myCmdIsCaseManagerExist.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                //    myCmdIsCaseManagerExist.Parameters["@ClientCode"].Value = prmStrClientCode;
                //    myCmdIsCaseManagerExist.Parameters["@ReportType"].Value = prmStrReportType;

                //    strIsCaseManagerExists = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdIsCaseManagerExist);

                //    if (strIsCaseManagerExists.Equals("True") || strIsCaseManagerExists.Equals("1"))
                //    {

                //        //@@@ Start : Update Case Manager
                //        myCmdUpdateCaseManager.CommandText = "sp_UpdateClientCaseManagerSetup";
                //        myCmdUpdateCaseManager.CommandType = CommandType.StoredProcedure;
                //        myCmdUpdateCaseManager.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                //        myCmdUpdateCaseManager.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                //        myCmdUpdateCaseManager.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                //        myCmdUpdateCaseManager.Parameters.Add("@CaseManager", SqlDbType.VarChar, 50);
                //        myCmdUpdateCaseManager.Parameters["@ClientCode"].Value = prmStrClientCode;
                //        myCmdUpdateCaseManager.Parameters["@ReportType"].Value = prmStrReportType;
                //        myCmdUpdateCaseManager.Parameters["@CaseManager"].Value = prmStrCaseManager;
                //        strResultCaseManager = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdUpdateCaseManager);
                //        if (strResultCaseManager.Equals("") || strResultCaseManager.Equals("False") || strResultCaseManager.Equals("0"))
                //        {
                //            bResult = false;
                //            connInsertClientReportTypeBudget.RollBackTransaction();
                //        }
                //        //@@@ End : Update Case Manager

                //    }
                //    else
                //    {
                //        //@@@ Start : Insert Case Manager
                //        myCmdInsertCaseManager.CommandText = "sp_InsertClientCaseManagerSetup";
                //        myCmdInsertCaseManager.CommandType = CommandType.StoredProcedure;
                //        myCmdInsertCaseManager.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                //        myCmdInsertCaseManager.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                //        myCmdInsertCaseManager.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                //        myCmdInsertCaseManager.Parameters.Add("@CaseManager", SqlDbType.VarChar, 50);
                //        myCmdInsertCaseManager.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                //        myCmdInsertCaseManager.Parameters["@ClientCode"].Value = prmStrClientCode;
                //        myCmdInsertCaseManager.Parameters["@ReportType"].Value = prmStrReportType;
                //        myCmdInsertCaseManager.Parameters["@CaseManager"].Value = prmStrCaseManager;
                //        myCmdInsertCaseManager.Parameters["@CreatedBy"].Value = prmStrCreatedBy;
                //        strResultCaseManager = connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdInsertCaseManager);
                //        if (strResultCaseManager.Equals("") || strResultCaseManager.Equals("False") || strResultCaseManager.Equals("0"))
                //        {
                //            bResult = false;
                //            connInsertClientReportTypeBudget.RollBackTransaction();
                //        }
                //        //@@@ End : Insert Case Manager

                //    }
                //}
                // Code Commented By Deepak to not to save case manager from Step 5 Ended

                //Start Adam 24-Mar-09 Auto Add Subject
                if (!((strNoIndividual.ToUpper() == "NULL") && (strNoCompany.ToUpper() == "NULL")))
                {
                    if (bResult)
                    {
                        //@@@ Start : Is Case manage Exists
                        myCmdUpdateAutoAddSubject.CommandText = "sp_UpdateAutoSubject";
                        myCmdUpdateAutoAddSubject.CommandType = CommandType.StoredProcedure;
                        myCmdUpdateAutoAddSubject.Transaction = connInsertClientReportTypeBudget.currentTransaction;
                        myCmdUpdateAutoAddSubject.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        myCmdUpdateAutoAddSubject.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                        myCmdUpdateAutoAddSubject.Parameters.Add("@NoIndividual", SqlDbType.Int);
                        myCmdUpdateAutoAddSubject.Parameters.Add("@NoCompany", SqlDbType.Int);
                        myCmdUpdateAutoAddSubject.Parameters["@ClientCode"].Value = prmStrClientCode;
                        myCmdUpdateAutoAddSubject.Parameters["@ReportType"].Value = prmStrReportType;
                        myCmdUpdateAutoAddSubject.Parameters["@NoIndividual"].Value = Convert.ToInt16(strNoIndividual);
                        myCmdUpdateAutoAddSubject.Parameters["@NoCompany"].Value = Convert.ToInt16(strNoCompany);
                        bAutoAddSubject = Convert.ToBoolean(connInsertClientReportTypeBudget.cmdScalarStoredProc(myCmdUpdateAutoAddSubject));
                        if (!bAutoAddSubject)
                        {
                            bResult = false;
                            connInsertClientReportTypeBudget.RollBackTransaction();
                        }
                    }
                }
                //End  Adam 24-Mar-09 Auto Add Subject
            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudget.RollBackTransaction();
            }
            finally
            {
                myCmdIsBudgetExist.Dispose();
                myCmdIsCaseManagerExist.Dispose();
                myCmdInsertCaseManager.Dispose();
                myCmdInsertBudget.Dispose();
                myCmdUpdateBudget.Dispose();
                myCmdUpdateCaseManager.Dispose();
                myCmdUpdateClientCurrency.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudget.CommitTransaction();
                }

                connInsertClientReportTypeBudget.Dispose();
                connInsertClientReportTypeBudget.Close();

            }


            return bResult;

        }
        public string GetCaseManagerMasterEmail(string prmStrCaseManager)
        {
            string strReturn;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCaseManagerMasterEmail";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@CaseManager", SqlDbType.VarChar);
            myCmd.Parameters["@CaseManager"].Value = prmStrCaseManager;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCaseManagerMasterEmail";
            strReturn = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strReturn;
        }
        /*@@@ End By Eric 20080506 @@@*/


        // Overloaded function
        public Boolean SaveClientBP(string sClCode, string sClName, string sCurrency, string sCreatedBy)
        {
            SqlCommand sqlCmd = new SqlCommand();
            string sStatus;

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_SaveClientBP";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                sqlCmd.Parameters.Add("@ClName", SqlDbType.VarChar, 100);
                sqlCmd.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
                //sqlCmd.Parameters.Add("@BPRegComplete", SqlDbType.Bit);
                sqlCmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
                sqlCmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);

                sqlCmd.Parameters["@ClCode"].Value = sClCode;
                sqlCmd.Parameters["@ClName"].Value = sClName;
                sqlCmd.Parameters["@Currency"].Value = sCurrency;
                //sqlCmd.Parameters["@BPRegComplete"].Value = false;
                sqlCmd.Parameters["@CreatedDate"].Value = System.DateTime.Now;
                sqlCmd.Parameters["@CreatedBy"].Value = sCreatedBy;

                conn.Open();
                conn.callingMethod = "ISBL.Admin.SaveClientBP";
                sStatus = conn.cmdScalarStoredProc(sqlCmd);
                sqlCmd.Dispose();
                conn.Close();
                if (sStatus == "1" || sStatus == "True")
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Created by Stev, 30 Oct 08
        // Purpose : Delete Client, but only disable the entire Client Login, return true once done.
        // Related : App_Code/ocrs_ws.vb. 
        public Boolean DeactivateClientLogin(string sClCode)
        {
            SqlCommand sqlCmd = new SqlCommand();
            string sStatus;

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "sp_DeactivateClientLogin";

                sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 15);
                sqlCmd.Parameters["@ClCode"].Value = sClCode;

                conn.Open();
                conn.callingMethod = "ISBL.Admin.DeactivateClientLogin";
                sStatus = conn.cmdScalarStoredProc(sqlCmd);
                sqlCmd.Dispose();
                conn.Close();

                if (sStatus == "1" || sStatus == "True")
                    return true;
                else
                    return false;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        /*@@@ End by Adam 20081201 OCRS - Phase 2 Web Service Security and Subject Updates from Savvion @@@ */

        /*@@@ Start by Adam 20090310 ISIS3 - Report Variant */
        public DataSet CheckReportVariantExist(string strClientCode, string strReportType, string strVariant)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckReportVariantExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@Variant"].Value = strVariant;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckReportVariantExist";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean IsReportVariantEmpty(string strClientCode, string strReportType)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_IsReportVariantEmpty";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.IsReportVariantEmpty";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }            
        }

        public Boolean DeleteClientReportVariant(string strClientCode, string strReportType, string strVariant)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_DeleteClientReportVariant";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                mycmd.Parameters["@Variant"].Value = strVariant;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.DeleteClientReportVariant";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }

        public Boolean UpdateClientReportVariantDefault(string strClientCode, string strReportType, string strVariant, Boolean bDefault)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_UpdateClientReportVariantDefault";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
                mycmd.Parameters.Add("@Default", SqlDbType.Bit);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                mycmd.Parameters["@Variant"].Value = strVariant;
                mycmd.Parameters["@Default"].Value = bDefault;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateClientReportVariantDefault";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }


        public Boolean UpdateClientReportVariant(string strClientCode, string strReportType, string strOriVariant, string strVariant, string strCountry, DataTable tblRE, Boolean bDefault, string strCreatedBy)
        {
            try
            {
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateClientReportVariant.sp_RemoveClientReportVariant";
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_RemoveClientReportVariant";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters.Add("@OriVariant", SqlDbType.VarChar, 250);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                mycmd.Parameters["@OriVariant"].Value = strOriVariant;
                conn.cmdScalarStoredProc(mycmd);

                if (!(tblRE == null))
                {
                    foreach (DataRow tRow in tblRE.Rows)
                    {
                        SqlCommand Cmd = new SqlCommand();
                        conn.callingMethod = "ISBL.Admin.UpdateClientReportVariant.sp_InsertClientReportVariant";
                        Cmd.Connection = conn.Connection;
                        Cmd.CommandText = "sp_InsertClientReportVariant";
                        Cmd.CommandType = CommandType.StoredProcedure;

                        Cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        Cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                        Cmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
                        Cmd.Parameters.Add("@Country", SqlDbType.VarChar, 4000);
                        //Cmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 15); //Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        Cmd.Parameters.Add("@SubjectType", SqlDbType.TinyInt); //Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        Cmd.Parameters.Add("@RE", SqlDbType.VarChar, 4);
                        Cmd.Parameters.Add("@Default", SqlDbType.Bit);
                        Cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                        Cmd.Parameters["@ClientCode"].Value = strClientCode;
                        Cmd.Parameters["@ReportType"].Value = strReportType;
                        Cmd.Parameters["@Variant"].Value = strVariant;
                        Cmd.Parameters["@Country"].Value = strCountry;
                        //Cmd.Parameters["@SubjectType"].Value = tRow["SubjectType"];//Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        Cmd.Parameters["@SubjectType"].Value = Convert.ToInt16(tRow["SubjectType"].ToString());//Jul 2013 Adam ISIS - Atlas Data Synch change to varchar to TinyInt
                        Cmd.Parameters["@RE"].Value = tRow["RECode"];
                        Cmd.Parameters["@Default"].Value = bDefault;
                        Cmd.Parameters["@CreatedBy"].Value = strCreatedBy;
                        conn.cmdScalarStoredProc(Cmd);
                    }
                }
                conn.Close();
                return true;
            }
            catch
            {
                return false;
            }            
        }

        public DataSet GetClientReportVariantRE(string strClientCode, string strReportType, string strVariant)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientReportVariantRE";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@Variant"].Value = strVariant;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportVariantRE";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        //added for JPMC changes start

        public DataSet GetClientReportSubReportBudget(string strClientCode, string strReportTypeCode, string strSubReportTypeCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmdrep = new SqlCommand();
            myCmdrep.Connection = conn.Connection;
            myCmdrep.CommandText = "sp_GetClientReportSubReportBudget";
            myCmdrep.CommandType = CommandType.StoredProcedure;
            myCmdrep.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmdrep.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 15);
            myCmdrep.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 250);
            myCmdrep.Parameters["@ClientCode"].Value = strClientCode;
            myCmdrep.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmdrep.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            SqlDataAdapter sdarep = new SqlDataAdapter();
            sdarep.SelectCommand = myCmdrep;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportSubReportBudget";
            ds = conn.FillDataSet(sdarep);
            conn.Close();
            myCmdrep.Dispose();
            sdarep.Dispose();
            return ds;
        }
        public DataSet GetClientReportSubReportMap(string strClientCode, string strReportTypeCode, string strSubReportTypeCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmdrep = new SqlCommand();
            myCmdrep.Connection = conn.Connection;
            myCmdrep.CommandText = "sp_GetClientReportSubReportMap";
            myCmdrep.CommandType = CommandType.StoredProcedure;
            myCmdrep.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmdrep.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 15);
            myCmdrep.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 250);
            myCmdrep.Parameters["@ClientCode"].Value = strClientCode;
            myCmdrep.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmdrep.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            SqlDataAdapter sdarep = new SqlDataAdapter();
            sdarep.SelectCommand = myCmdrep;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportSubReportMap";
            ds = conn.FillDataSet(sdarep);
            conn.Close();
            myCmdrep.Dispose();
            sdarep.Dispose();
            return ds;
        }
        public DataSet GetClientReportSubReportMapauto(string strClientCode, string strReportTypeCode, string strSubReportTypeCode, string strEntityType)
        {
            if (strEntityType == "")
            {
                strEntityType = "0";
            }
            if (strEntityType.Length == 0)
            {
                strEntityType = "0";
            }
            DataSet ds = new DataSet();
            SqlCommand myCmdrep = new SqlCommand();
            myCmdrep.Connection = conn.Connection;
            myCmdrep.CommandText = "sp_GetClientReportSubReportMapauto";
            myCmdrep.CommandType = CommandType.StoredProcedure;
            myCmdrep.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmdrep.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 15);
            myCmdrep.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 250);
            myCmdrep.Parameters["@ClientCode"].Value = strClientCode;
            myCmdrep.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmdrep.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            myCmdrep.Parameters.Add("@EntityType", SqlDbType.TinyInt);
            myCmdrep.Parameters["@EntityType"].Value = Convert.ToByte(strEntityType.ToString());
            SqlDataAdapter sdarep = new SqlDataAdapter();
            sdarep.SelectCommand = myCmdrep;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportSubReportMapauto";
            ds = conn.FillDataSet(sdarep);
            conn.Close();
            myCmdrep.Dispose();
            sdarep.Dispose();
            return ds;
        }
        public string ClientReportSubReportMap(string strClientCode, string strReportTypeCode, string strSubReportTypeCode,string strCountry,string strEntityType,Boolean bDefault,string strAction,string strSelCountry)
        {
            string blnrepStatus = "";
            bool bResult = true;
            if (strEntityType == "")
            {
                strEntityType = "0";
            }
            if (strEntityType.Length == 0)
            {
                strEntityType = "0";
            }
            SqlCommand myCmd = new SqlCommand();
            try
            {
               
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_InsertClientReportSubReportMap";
                myCmd.CommandType = CommandType.StoredProcedure;
                myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@ClientCode"].Value = strClientCode;
                myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
                myCmd.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
                myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 4000);
                myCmd.Parameters["@Country"].Value = strCountry;
                myCmd.Parameters.Add("@EntityType", SqlDbType.TinyInt);
                myCmd.Parameters["@EntityType"].Value = Convert.ToByte(strEntityType.ToString());
                myCmd.Parameters.Add("@Action", SqlDbType.VarChar, 30);
                myCmd.Parameters["@Action"].Value = strAction;
                myCmd.Parameters.Add("@Default", SqlDbType.Bit);
                myCmd.Parameters["@Default"].Value = bDefault;

                myCmd.Parameters.Add("@SelCountry", SqlDbType.VarChar, 4000);
                myCmd.Parameters["@SelCountry"].Value = strSelCountry;

                myCmd.Parameters.Add("@INStatus", SqlDbType.VarChar, 50);
                myCmd.Parameters["@INStatus"].Direction = ParameterDirection.Output;
                conn.Open();
                SqlDataReader myReader;

                myReader = myCmd.ExecuteReader();
                return blnrepStatus = myCmd.Parameters["@INStatus"].Value.ToString();
                //conn.Open();
                //conn.callingMethod = "ISBL.General.ClientReportSubReportMap";
                //blnrepStatus = conn.cmdScalarStoredProc(myCmd);
                //if (blnrepStatus.Equals("") || blnrepStatus.Equals("False") || blnrepStatus.Equals("0"))
                //{
                //    bResult = false;
                    //conn.RollBackTransaction();
                //}
                
            }
            catch(Exception ex)
            {
                //bResult = false;
                blnrepStatus = "INSERT REPORT - ERROR";
                ///conn.RollBackTransaction();
            }
            finally
            {
                conn.Close();
                myCmd.Dispose();
            }
            return blnrepStatus;
        }
        public string ClientReportSubReportMapnew(string strClientCode, string strReportTypeCode, string strSubReportTypeCode, string strCountry, Boolean bDefault, string strAction)
        {
            string blnrepStatus = "";
            bool bResult = true;
            
            SqlCommand myCmd = new SqlCommand();
            try
            {

                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_InsertClientReportSubReportMap1";
                myCmd.CommandType = CommandType.StoredProcedure;
                myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@ClientCode"].Value = strClientCode;
                myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
                myCmd.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 30);
                myCmd.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
                myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 4000);
                myCmd.Parameters["@Country"].Value = strCountry;
                myCmd.Parameters.Add("@Action", SqlDbType.VarChar, 30);
                myCmd.Parameters["@Action"].Value = strAction;
                myCmd.Parameters.Add("@Default", SqlDbType.Bit);
                myCmd.Parameters["@Default"].Value = bDefault;

                myCmd.Parameters.Add("@INStatus", SqlDbType.VarChar, 50);
                myCmd.Parameters["@INStatus"].Direction = ParameterDirection.Output;
                conn.Open();
                SqlDataReader myReader;

                myReader = myCmd.ExecuteReader();
                return blnrepStatus = myCmd.Parameters["@INStatus"].Value.ToString();
                //conn.Open();
                //conn.callingMethod = "ISBL.General.ClientReportSubReportMap";
                //blnrepStatus = conn.cmdScalarStoredProc(myCmd);
                //if (blnrepStatus.Equals("") || blnrepStatus.Equals("False") || blnrepStatus.Equals("0"))
                //{
                //    bResult = false;
                //conn.RollBackTransaction();
                //}

            }
            catch (Exception ex)
            {
                //bResult = false;
                blnrepStatus = "INSERT REPORT - ERROR";
                ///conn.RollBackTransaction();
            }
            finally
            {
                conn.Close();
                myCmd.Dispose();
            }
            return blnrepStatus;
        }

        public string DelClientReportSubReportMap1(string strReportID)
        {
            string blnrepdelStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteClientReportSubReportMap1";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportID", SqlDbType.VarChar, 150);
            myCmd.Parameters["@ReportID"].Value = strReportID;
            //myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 4000);
            //myCmd.Parameters["@Country"].Value = strCountry;
            //myCmd.Parameters.Add("@EntityType", SqlDbType.TinyInt);
            //myCmd.Parameters["@EntityType"].Value = Convert.ToByte(strEntityType.ToString());
            conn.Open();
            conn.callingMethod = "ISBL.General.DelClientReportSubReportMap1";
            blnrepdelStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return blnrepdelStatus;
        }
        public string DelClientReportSubReportMap(string strClientCode, string strReportTypeCode, string strSubReportTypeCode,string strReportID)
        {
            string  blnrepdelStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteClientReportSubReportMap";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmd.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            myCmd.Parameters.Add("@ReportID", SqlDbType.VarChar, 150);
            myCmd.Parameters["@ReportID"].Value = strReportID;
            //myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 4000);
            //myCmd.Parameters["@Country"].Value = strCountry;
            //myCmd.Parameters.Add("@EntityType", SqlDbType.TinyInt);
            //myCmd.Parameters["@EntityType"].Value = Convert.ToByte(strEntityType.ToString());
            conn.Open();
            conn.callingMethod = "ISBL.General.DelClientReportSubReportMap";
            blnrepdelStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return blnrepdelStatus;
        }
        public Boolean DelClientReportSubReportBudget(string strClientCode, string strReportTypeCode, string strSubReportTypeCode)
        {
            Boolean blnrepdelStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteClientReportSubReportBudget";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmd.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.DelClientReportSubReportBudget";
            blnrepdelStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnrepdelStatus;
        }
        // end of Jpmc changes
        public DataSet GetCountryListAll()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryListAll";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCountryListAll";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetCountryListMapping(string strClientCode, string strReportTypeCode, string strSubReportTypeCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryListMapping";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strReportTypeCode;
            myCmd.Parameters.Add("@SubReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@SubReportTypeCode"].Value = strSubReportTypeCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetCountryListMapping";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet ListClientReportVariant(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_ListClientReportVariant";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.ListClientReportVariant";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /*@@@ End by Adam 20090310 ISIS3 - Report Variant */

        /*@@@ Start by Adam 20090324 ISIS3 - Auto Add Subject */
        public DataSet GetAutoAddSubject(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetAutoAddSubject";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetAutoAddSubject";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean UpdateAutoSubject(string strClientCode, string strReportType, int intNoIndividual, int intNoCompany)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_UpdateAutoSubject";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters.Add("@NoIndividual", SqlDbType.Int);
                mycmd.Parameters.Add("@NoCompany", SqlDbType.Int);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                mycmd.Parameters["@NoIndividual"].Value = intNoIndividual;
                mycmd.Parameters["@NoCompany"].Value = intNoCompany;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.UpdateAutoSubject";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }
        /*@@@ End by Adam 20090324 ISIS3 - Auto Add Subject */


        /*@@@ Start By Adam 20090515 Update Report Due Date */
        public Boolean UpdateReportDueDate(int DayDue, string ReportType)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateReportDueDate";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@DayDue", SqlDbType.Int);
            myCmd.Parameters["@DayDue"].Value = DayDue;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = ReportType;
            conn.Open();
            conn.callingMethod = "ISBL.General.sp_UpdateReportDueDate";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }
        /*@@@ End By Adam 20090515 Update Report Due Date */

        /*@@@ Start By Adam 20090525 Clear Login Instance Session by LoginID */
        public Boolean ClearLoginInstanceSession(string LoginID)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_ClearLoginInstanceSession";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = LoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.sp_UpdateReportDueDate";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }
        /*@@@ End By Adam 20090515 Update Report Due Date */


        /*@@@ Start By Adam Nov 2009 Export Clients Variant Infomation OCRS Phase 4 v 2.8b */
        public SqlDataReader GetClientVariant(string strLoginID, string strRole, ref SqlConnection sqlConnect)
        {
            string strSQL;

            if (strRole.ToUpper() == "A1") //Admin user
            {
                strSQL = "SELECT a.clientCode, a.ClientName, a.CreatedDate, a.CreatedBy, b.ReportType, " ;
                strSQL = strSQL + "b.Variant, b.SubjectType, b.RE, ResearchElementDescription ";
                strSQL = strSQL + "FROM ocrsClientMaster a, ocrsClientReportVariant b, ocrsREMaster c ";
                strSQL = strSQL + "WHERE a.ClientCode in ( ";
                strSQL = strSQL + "SELECT ClientCode FROM ocrsLoginDBMClient ";
                strSQL = strSQL + "WHERE a.ClientCode = b.ClientCode ";
                strSQL = strSQL + "AND c.ResearchElementCode = b.RE)  ";
                strSQL = strSQL + "ORDER BY a.ClientCode, b.ReportType, b.Variant, b.SubjectType, b.RE ";
            }
            else
            {
                strSQL = "SELECT a.clientCode, a.ClientName, a.CreatedDate, a.CreatedBy, b.ReportType, ";
                strSQL = strSQL + "b.Variant, b.SubjectType, b.RE, ResearchElementDescription ";
                strSQL = strSQL + "FROM ocrsClientMaster a, ocrsClientReportVariant b, ocrsREMaster c ";
                strSQL = strSQL + "WHERE a.ClientCode in ( ";
                strSQL = strSQL + "SELECT ClientCode FROM ocrsLoginDBMClient ";
                strSQL = strSQL + "WHERE LoginID = '" + strLoginID + "' ";
                strSQL = strSQL + "AND a.ClientCode = b.ClientCode ";
                strSQL = strSQL + "AND c.ResearchElementCode = b.RE) ";
                strSQL = strSQL + "ORDER BY a.ClientCode, b.ReportType, b.Variant, b.SubjectType, b.RE ";
            }

            myDataReader = conn.cmdReader(strSQL, ref sqlConnect);

            return myDataReader;

        }
        /*@@@ End By Adam Nov 2009 Export Clients Variant Infomation OCRS Phase 4 v 2.8b */


        /*@@@ Start ISIS Failure Proofing*/

        public SqlDataReader GetOfflineOrderTrail(string strEventLog, string strOrderID, string strCRN
           , string strClientUpdateDateFrom, string strClientUpdateDateTo, string strStatus, string strNewEventLog, string strAction, ref SqlConnection sqlConnect)
        {
            ISBL.General oISBLGen = new ISBL.General();
            string strSql;
            string strwhereclause;
            bool blWhere = false;
            strSql = "SELECT * FROM ocrsOfflineTrail ";
            if (strEventLog.Length > 0)
            {
                strwhereclause = "EventLogID = '" + strEventLog + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if (strNewEventLog.Length > 0)
            {
                strwhereclause = "NewEventLogID = '" + strNewEventLog + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if (strOrderID.Length > 0)
            {
                strwhereclause = "OrderID = '" + strOrderID + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            if (strCRN.Length > 0)
            {
                strwhereclause = "CRN LIKE '%" + oISBLGen.SafeSqlLiteral(strCRN) + "%' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if (strStatus.Length > 0)
            {
                strwhereclause = "Status = '" + strStatus + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if (strAction.Length > 0)
            {
                strwhereclause = "ActionPerform = '" + strAction + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            if ((strClientUpdateDateFrom.Length > 0) && (strClientUpdateDateTo.Length > 0))
            {
                strwhereclause = "ClientUpdateDate BETWEEN '" + strClientUpdateDateFrom + "' AND '" + strClientUpdateDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strClientUpdateDateFrom.Length > 0)
            {
                strwhereclause = "ClientUpdateDate >= '" + strClientUpdateDateFrom + "' ";
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }
            else if (strClientUpdateDateTo.Length > 0)
            {
                strwhereclause = "ClientUpdateDate <= '" + strClientUpdateDateTo + " 23:59:59:999' "; //added 23:59:59:999 by Adam 8 Jan 09
                strSql = constructSQL(strSql, strwhereclause, blWhere);
                blWhere = true;
            }

            strSql = strSql + " ORDER BY ClientUpdateDate DESC";

            oISBLGen.DisposeConnection();
            oISBLGen = null;

            myDataReader = conn.cmdReader(strSql, ref sqlConnect);

            return myDataReader;
        }
        /*@@@ End  ISIS Failure Proofing*/

        /*@@@ Start ISIS v2 Phase 1 Release 3 */
        public Boolean IsNewDayDueLessThanStandard(string strReportType, int intNewDayDue)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_IsNewDayDueLessThanStandard";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters.Add("@NewDayDue", SqlDbType.Int);
                mycmd.Parameters["@ReportType"].Value = strReportType;
                mycmd.Parameters["@NewDayDue"].Value = intNewDayDue;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.IsNewDayDueLessThanStandard";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }

        public Boolean IsReportTypeSetup(string strClientCode, string strReportType)
        {
            try
            {
                Boolean blnStatus;
                SqlCommand mycmd = new SqlCommand();
                mycmd.Connection = conn.Connection;
                mycmd.CommandText = "sp_IsReportTypeSetup";
                mycmd.CommandType = CommandType.StoredProcedure;
                mycmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                mycmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                mycmd.Parameters["@ClientCode"].Value = strClientCode;
                mycmd.Parameters["@ReportType"].Value = strReportType;
                conn.Open();
                conn.callingMethod = "ISBL.Admin.IsReportTypeSetup";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(mycmd));
                conn.Close();
                mycmd.Dispose();
                return blnStatus;
            }
            catch
            {
                return false;
            }
        }

        /*@@@ End ISIS v2 Phase 1 Release 3 */

        /*@@@ Start ISIS and Atlas Integration 2011 */
        public string GetOfflineTrailDetail(string strLogID, string strDataToGet)
        {
            string strData;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOfflineTrailDetail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LogID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@Data", SqlDbType.VarChar, 10);
            myCmd.Parameters["@LogID"].Value = new System.Guid(strLogID);
            myCmd.Parameters["@Data"].Value = strDataToGet;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetOfflineTrailDetail";
            strData = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strData;

        }
        /*@@@ End ISIS and Atlas Integration 2011 */

        /*@@@ Start CR 24 Incomplete Report Setup cannot place order Mar 2014 */
        public bool IsClientReportSetupCompleted(string strClientCode, string strReportType)
        {
            bool bIsCompleted = true;

            string strStatus = GetClientReportSetupSectionStatus(strClientCode, strReportType);

            if (!strStatus.Equals("0"))
            {
                bIsCompleted = false;
            }

            return bIsCompleted;
        }

        public string GetClientReportSetupSectionStatus(string strClientCode, string strReportType)
        {
            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientReportSetupSectionStatus";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;

            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar);
            myCmd.Parameters["@ReportType"].Value = strReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportSetupSectionStatus";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }

        public void UpdateClientReportSetupSectionStatus(string strClientCode, string strReportType, string strStatus,string strSubReportType)
        {
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateClientReportSetupSectionStatus";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters.Add("@SetupSectionStatus", SqlDbType.VarChar);
            myCmd.Parameters["@SetupSectionStatus"].Value = strStatus;
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar);
            myCmd.Parameters["@SubReportType"].Value = strSubReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.UpdateClientReportSetupSectionStatus";
            conn.cmdNoneQuery(myCmd);
            conn.Close();

            myCmd.Dispose();
        }

        public string DeleteAllClientReportSetupData(string prmStrClientCode, string prmStrReportType)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DeleteAllClientReportSetupData";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.DeleteAllClientReportSetupData";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }

        public string GetLatestClientSetupSection(string prmStrClientCode)
        {

            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetLatestClientSetupSection";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetLatestClientSetupSection";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }





        /*@@@ End  CR 24 Incomplete Report Setup cannot place order Mar 2014 */



        /*Code for separate step 5 in JPMC client*/
        /// <summary>
        /// Code to insert Individual Budget
        /// </summary>
        /// <param name="prmStrClientCode"></param>
        /// <param name="prmStrReportType"></param>
        /// <param name="prmFBasePrice"></param>
        /// <param name="prmIntIndividualMaxPrice"></param>
        /// <param name="prmFIndividualCost"></param>
        /// <param name="prmStrCreatedBy"></param>
        /// <param name="prmStrCurrency"></param>
        /// <param name="strNoIndividual"></param>
        /// <param name="prmFBasePriceExpress"></param>
        /// <param name="prmIntIndividualMaxPriceExpress"></param>
        /// <param name="prmFIndividualCostExpress"></param>
        /// <param name="prmTATNormal"></param>
        /// <param name="prmTATExpress"></param>
        /// <param name="prmTATNormalBulk"></param>
        /// <param name="subreporttype"></param>
        /// <param name="prmCountry"></param>
        /// <param name="prmEntityType"></param>
        /// <returns></returns>
        public bool InsertClientReportTypeBudgetIndividualMap(string prmStrClientCode, string prmStrReportType, 
            int prmIntIndividualMaxPrice, float prmFIndividualCost, string prmStrCreatedBy, 
             int prmIntIndividualMaxPriceExpress, float prmFIndividualCostExpress, string subreporttype, string prmCountry, string prmEntityType)
        {

            bool bResult = true;
            string strResultBudget = "";


            ISDL.Connect connInsertClientReportTypeBudgetIndividual = new ISDL.Connect();
            connInsertClientReportTypeBudgetIndividual.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetIndividual.Open();
            connInsertClientReportTypeBudgetIndividual.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetIndividual.Connection;

  

            try
            {


                if (bResult)  
                {
                        myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetIndividualMap";
                        myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                        myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetIndividual.currentTransaction;
                        myCmdInsertBudget.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        myCmdInsertBudget.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                     
                        myCmdInsertBudget.Parameters.Add("@MaxIndividualForBasePrice", SqlDbType.Int);
                        
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerIndividual", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                        
                        myCmdInsertBudget.Parameters.Add("@MaxIndividualForBasePriceExpress", SqlDbType.Int);
                        
                        myCmdInsertBudget.Parameters.Add("@AdditionalCostPerIndividualExpress", SqlDbType.Float);
                      
                        myCmdInsertBudget.Parameters.Add("@pEntityType", SqlDbType.VarChar, 30);
                       
                        myCmdInsertBudget.Parameters.Add("@pCountry", SqlDbType.VarChar, 150);
                      
                        myCmdInsertBudget.Parameters.Add("@psubreporttype", SqlDbType.VarChar, 15);
                        
                        myCmdInsertBudget.Parameters["@ClientCode"].Value = prmStrClientCode;
                        myCmdInsertBudget.Parameters["@ReportType"].Value = prmStrReportType;
                       
                        myCmdInsertBudget.Parameters["@MaxIndividualForBasePrice"].Value = prmIntIndividualMaxPrice;
                        
                        myCmdInsertBudget.Parameters["@AdditionalCostPerIndividual"].Value = prmFIndividualCost;
                        myCmdInsertBudget.Parameters["@CreatedBy"].Value = prmStrCreatedBy;
                        
                      
                        myCmdInsertBudget.Parameters["@MaxIndividualForBasePriceExpress"].Value = prmIntIndividualMaxPriceExpress;
                        
                        myCmdInsertBudget.Parameters["@AdditionalCostPerIndividualExpress"].Value = prmFIndividualCostExpress;
                       
                        myCmdInsertBudget.Parameters["@pEntityType"].Value = prmEntityType;
                       
                        myCmdInsertBudget.Parameters["@pCountry"].Value = prmCountry;
                        
                        myCmdInsertBudget.Parameters["@psubreporttype"].Value = subreporttype;
                      
                        strResultBudget = connInsertClientReportTypeBudgetIndividual.cmdScalarStoredProc(myCmdInsertBudget);
                        if (strResultBudget.Equals("") || strResultBudget.Equals("False") || strResultBudget.Equals("0"))
                        {
                            bResult = false;
                            connInsertClientReportTypeBudgetIndividual.RollBackTransaction();
                        }
                        
                    }
                    
               

               

              
            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetIndividual.RollBackTransaction();
            }
            finally
            {
               
                myCmdInsertBudget.Dispose();
               
                

                if (bResult)
                {
                    connInsertClientReportTypeBudgetIndividual.CommitTransaction();
                }

                connInsertClientReportTypeBudgetIndividual.Dispose();
                connInsertClientReportTypeBudgetIndividual.Close();

            }


            return bResult;

        }



        public bool InsertClientReportTypeBudgetCompanyMap(string prmStrClientCode, string prmStrReportType,
           int prmIntComapnyMaxPrice, float prmFCompanyCost, string prmStrCreatedBy,
            int prmIntComapnyMaxPriceExpress, float prmFComapnyCostExpress, string subreporttype, string prmCountry, string prmEntityType)
        {

            bool bResult = true;
            string strResultBudget = "";


            ISDL.Connect connInsertClientReportTypeBudgetCompanyMap = new ISDL.Connect();
            connInsertClientReportTypeBudgetCompanyMap.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetCompanyMap.Open();
            connInsertClientReportTypeBudgetCompanyMap.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetCompanyMap.Connection;



            try
            {


                if (bResult)
                {
                    myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetCompanyMap";
                    myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                    myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetCompanyMap.currentTransaction;
                    myCmdInsertBudget.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                    myCmdInsertBudget.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                    myCmdInsertBudget.Parameters.Add("@MaxCompanyForBasePrice", SqlDbType.Int);
                    myCmdInsertBudget.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                    myCmdInsertBudget.Parameters.Add("@MaxCompanyForBasePriceExpress", SqlDbType.Int);
                    myCmdInsertBudget.Parameters.Add("@AdditionalCostPerCompanyExpress", SqlDbType.Float);
                    myCmdInsertBudget.Parameters.Add("@pEntityType", SqlDbType.Int);
                    myCmdInsertBudget.Parameters.Add("@pCountry", SqlDbType.VarChar,150);
                    myCmdInsertBudget.Parameters.Add("@AdditionalCostPerCompany", SqlDbType.Float);
                    myCmdInsertBudget.Parameters.Add("@psubreporttype", SqlDbType.VarChar, 15);
                    // values
                    myCmdInsertBudget.Parameters["@ClientCode"].Value = prmStrClientCode;
                    myCmdInsertBudget.Parameters["@ReportType"].Value = prmStrReportType;
                    myCmdInsertBudget.Parameters["@MaxCompanyForBasePrice"].Value = prmIntComapnyMaxPrice;
                    myCmdInsertBudget.Parameters["@CreatedBy"].Value = prmStrCreatedBy;
                    myCmdInsertBudget.Parameters["@MaxCompanyForBasePriceExpress"].Value = prmIntComapnyMaxPriceExpress;
                    myCmdInsertBudget.Parameters["@AdditionalCostPerCompanyExpress"].Value = prmFComapnyCostExpress;
                    myCmdInsertBudget.Parameters["@pEntityType"].Value = prmEntityType;
                    myCmdInsertBudget.Parameters["@pCountry"].Value = prmCountry;
                    myCmdInsertBudget.Parameters["@AdditionalCostPerCompany"].Value = prmFCompanyCost;                    
                    myCmdInsertBudget.Parameters["@psubreporttype"].Value = subreporttype;

                    strResultBudget = connInsertClientReportTypeBudgetCompanyMap.cmdScalarStoredProc(myCmdInsertBudget);
                    if (strResultBudget.Equals("") || strResultBudget.Equals("False") || strResultBudget.Equals("0"))
                    {
                        bResult = false;
                        connInsertClientReportTypeBudgetCompanyMap.RollBackTransaction();
                    }

                }






            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetCompanyMap.RollBackTransaction();
            }
            finally
            {

                myCmdInsertBudget.Dispose();



                if (bResult)
                {
                    connInsertClientReportTypeBudgetCompanyMap.CommitTransaction();
                }

                connInsertClientReportTypeBudgetCompanyMap.Dispose();
                connInsertClientReportTypeBudgetCompanyMap.Close();

            }


            return bResult;

        }



        /// <summary>
        /// Code to insert Company
        /// </summary>
        /// <param name="prmStrClientCode"></param>
        /// <param name="prmStrReportType"></param>
        /// <param name="prmFBasePrice"></param>
        /// <param name="prmIntCompanyMaxPrice"></param>
        /// <param name="prmFCompanyCost"></param>
        /// <param name="prmStrCreatedBy"></param>
        /// <param name="prmStrCurrency"></param>
        /// <param name="strNoCompany"></param>
        /// <param name="prmFBasePriceExpress"></param>
        /// <param name="prmIntCompanyMaxPriceExpress"></param>
        /// <param name="prmFCompanyCostExpress"></param>
        /// <param name="prmTATNormal"></param>
        /// <param name="prmTATExpress"></param>
        /// <param name="prmTATNormalBulk"></param>
        /// <param name="subreporttype"></param>
        /// <param name="prmCountry"></param>
        /// <param name="prmEntityType"></param>
        /// <returns></returns>
        public bool InsertClientReportTypeBudgetCompany(string prmStrClientCode, string prmStrReportType, float prmFBasePrice,
            string prmStrCreatedBy, string prmStrCurrency,  string strNoCompany, 
            float prmFBasePriceExpress,  int prmTATNormal, int prmTATExpress,
            int prmTATNormalBulk, string subreporttype, string strNoIndividual)
        {

            bool bResult = true;
            string strResultBudget = "";
            
            string strResultCurrency = ""; 
            

            ISDL.Connect connInsertClientReportTypeBudgetCompany = new ISDL.Connect();
            connInsertClientReportTypeBudgetCompany.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetCompany.Open();
            connInsertClientReportTypeBudgetCompany.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetCompany.Connection;
            SqlCommand myCmdUpdateClientCurrency = new SqlCommand();
            myCmdUpdateClientCurrency.Connection = connInsertClientReportTypeBudgetCompany.Connection;

       

            try
            {

                
                myCmdUpdateClientCurrency.CommandText = "sp_UpdateClientCurrency";
                myCmdUpdateClientCurrency.CommandType = CommandType.StoredProcedure;
                myCmdUpdateClientCurrency.Transaction = connInsertClientReportTypeBudgetCompany.currentTransaction;
                myCmdUpdateClientCurrency.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                myCmdUpdateClientCurrency.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
                myCmdUpdateClientCurrency.Parameters["@ClientCode"].Value = prmStrClientCode;
                myCmdUpdateClientCurrency.Parameters["@Currency"].Value = prmStrCurrency;

                strResultCurrency = connInsertClientReportTypeBudgetCompany.cmdScalarStoredProc(myCmdUpdateClientCurrency);

                if (strResultCurrency.Equals("") || strResultCurrency.Equals("False") || strResultCurrency.Equals("0"))
                {
                    bResult = false;
                    connInsertClientReportTypeBudgetCompany.RollBackTransaction();
                }
                //@@@ End : Update Client Currency

                if (bResult)  //Adam 6-Feb-09
                {                   
                        myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetCompany";
                        myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                        myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetCompany.currentTransaction;
                        myCmdInsertBudget.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                        myCmdInsertBudget.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                        myCmdInsertBudget.Parameters.Add("@BasePrice", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
                        myCmdInsertBudget.Parameters.Add("@BasePriceExpress", SqlDbType.Float);
                        myCmdInsertBudget.Parameters.Add("@TATNormal", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@TATExpress", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@TATNormalBulk", SqlDbType.Int); 
                        myCmdInsertBudget.Parameters.Add("@pCurrency", SqlDbType.VarChar, 30);
                        myCmdInsertBudget.Parameters.Add("@AutoAddSubjectIndividual", SqlDbType.VarChar, 30);
                        myCmdInsertBudget.Parameters.Add("@AutoAddSubjectCompany", SqlDbType.Int);
                        myCmdInsertBudget.Parameters.Add("@psubreporttype", SqlDbType.VarChar, 15);
                    // values
                        myCmdInsertBudget.Parameters["@ClientCode"].Value = prmStrClientCode;
                        myCmdInsertBudget.Parameters["@ReportType"].Value = prmStrReportType;
                        myCmdInsertBudget.Parameters["@BasePrice"].Value = prmFBasePrice;
                        myCmdInsertBudget.Parameters["@CreatedBy"].Value = prmStrCreatedBy;
                        myCmdInsertBudget.Parameters["@BasePriceExpress"].Value = prmFBasePriceExpress;
                        myCmdInsertBudget.Parameters["@TATNormal"].Value = prmTATNormal;
                        myCmdInsertBudget.Parameters["@TATExpress"].Value = prmTATExpress;
                        myCmdInsertBudget.Parameters["@TATNormalBulk"].Value = prmTATNormalBulk; 
                        myCmdInsertBudget.Parameters["@pCurrency"].Value = prmStrCurrency;
                        myCmdInsertBudget.Parameters["@AutoAddSubjectIndividual"].Value = strNoIndividual;
                        myCmdInsertBudget.Parameters["@AutoAddSubjectCompany"].Value = strNoCompany;
                        myCmdInsertBudget.Parameters["@psubreporttype"].Value = subreporttype;

                        strResultBudget = connInsertClientReportTypeBudgetCompany.cmdScalarStoredProc(myCmdInsertBudget);
                        if (strResultBudget.Equals("") || strResultBudget.Equals("False") || strResultBudget.Equals("0"))
                        {
                            bResult = false;
                            connInsertClientReportTypeBudgetCompany.RollBackTransaction();
                        }
                        
                    }
            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetCompany.RollBackTransaction();
            }
            finally
            {
                myCmdInsertBudget.Dispose();
                myCmdUpdateClientCurrency.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudgetCompany.CommitTransaction();
                }

                connInsertClientReportTypeBudgetCompany.Dispose();
                connInsertClientReportTypeBudgetCompany.Close();

            }


            return bResult;

        }

        public DataSet BindBaseValuesStep5(string strClientCode, string strReportType,string strsubreporttype)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "BindBaseValuesStep5";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@psubreporttype", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@psubreporttype"].Value = strsubreporttype;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.BindBaseValuesStep5";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        /*@@@ Bind step 5 report , Created By Deepak
         *  to bind only those report which status is completed in step 5
         * 
         * @@@*/
        public DataSet GetClientReportTypeListStep5(string prmStrClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientReportTypeListStep5";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.Admin.GetClientReportTypeListStep5";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }


        public Boolean CheckSubReportExistReportType(string prmStrClientCode, string prmStrReportType, string SubreportType)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckSubReportExistReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            myCmd.Parameters["@SubReportType"].Value = SubreportType;
            

            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckSubReportExistReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }

        public Boolean CheckSubjectTypeExistReportType(string prmStrClientCode, string prmStrReportType, string SubjectType)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckSubjectTypeExistReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 10);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            myCmd.Parameters["@SubjectType"].Value = SubjectType;


            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckSubjectTypeExistReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }

        public Boolean CheckCountryExistReportType(string prmStrClientCode, string prmStrReportType, string SubreportType, string SubjectType, string Country)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckCountryExistReportType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 10);
            myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 400);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            myCmd.Parameters["@SubReportType"].Value = SubreportType;
            myCmd.Parameters["@SubjectType"].Value = SubjectType;
            myCmd.Parameters["@Country"].Value = Country;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckCountryExistReportType";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }
        public Boolean CheckCountryExistNonJPMC(string prmStrClientCode, string prmStrReportType, string SubreportType, string Country)
        {
            Boolean blnStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckCountryExistForNonJPMC";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            
            myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 400);
            myCmd.Parameters["@ClientCode"].Value = prmStrClientCode;
            myCmd.Parameters["@ReportType"].Value = prmStrReportType;
            myCmd.Parameters["@SubReportType"].Value = SubreportType;
            
            myCmd.Parameters["@Country"].Value = Country;

            conn.Open();
            conn.callingMethod = "ISBL.Admin.CheckCountryExistNonJPMC";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();

            myCmd.Dispose();

            return blnStatus;
        }
        // new fuctions for step 5

        public void InsertClientReportTypeBudgetCompanyPaging(DataTable dtparam)
        {

            bool bResult = true;
            bool strResultBudget = true;


            ISDL.Connect connInsertClientReportTypeBudgetCompanyMap = new ISDL.Connect();
            connInsertClientReportTypeBudgetCompanyMap.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetCompanyMap.Open();
            connInsertClientReportTypeBudgetCompanyMap.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetCompanyMap.Connection;

            try
            {
                if (bResult)
                {
                    myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetCompanyAll";
                    myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                    myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetCompanyMap.currentTransaction;
                    myCmdInsertBudget.Parameters.Add("@Tablles", SqlDbType.Structured);
                    myCmdInsertBudget.Parameters["@Tablles"].Value = dtparam;

                    strResultBudget = connInsertClientReportTypeBudgetCompanyMap.cmdNoneQuery(myCmdInsertBudget);

                }

            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetCompanyMap.RollBackTransaction();
            }
            finally
            {

                myCmdInsertBudget.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudgetCompanyMap.CommitTransaction();
                }

                connInsertClientReportTypeBudgetCompanyMap.Dispose();
                connInsertClientReportTypeBudgetCompanyMap.Close();

            }


            

        }
        public void InsertClientReportTypeBudgetIndividualPaging(DataTable dtparam)
        {

            bool bResult = true;
            bool strResultBudget = true;

            ISDL.Connect connInsertClientReportTypeBudgetIndividual = new ISDL.Connect();
            connInsertClientReportTypeBudgetIndividual.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetIndividual.Open();
            connInsertClientReportTypeBudgetIndividual.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetIndividual.Connection;

            try
            {
                if (bResult)
                {
                    myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetIndividualAll";
                    myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                    myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetIndividual.currentTransaction;
                    myCmdInsertBudget.Parameters.Add("@Tablles", SqlDbType.Structured);
                    myCmdInsertBudget.Parameters["@Tablles"].Value = dtparam;
                  strResultBudget= connInsertClientReportTypeBudgetIndividual.cmdNoneQuery(myCmdInsertBudget);

                }
            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetIndividual.RollBackTransaction();
            }
            finally
            {

                myCmdInsertBudget.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudgetIndividual.CommitTransaction();
                }

                connInsertClientReportTypeBudgetIndividual.Dispose();
                connInsertClientReportTypeBudgetIndividual.Close();

            }


           

        }

        public void InsertClientReportTypeBudgetCompanyAll(DataTable dtparam)
        {

            bool bResult = true;
            bool strResultBudget = true;


            ISDL.Connect connInsertClientReportTypeBudgetCompanyMap = new ISDL.Connect();
            connInsertClientReportTypeBudgetCompanyMap.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetCompanyMap.Open();
            connInsertClientReportTypeBudgetCompanyMap.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetCompanyMap.Connection;

            try
            {
                if (bResult)
                {
                    myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetCompanyAllRow";
                    myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                    myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetCompanyMap.currentTransaction;
                    myCmdInsertBudget.Parameters.Add("@Tablles", SqlDbType.Structured);
                    myCmdInsertBudget.Parameters["@Tablles"].Value = dtparam;

                    strResultBudget = connInsertClientReportTypeBudgetCompanyMap.cmdNoneQuery(myCmdInsertBudget);

                }

            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetCompanyMap.RollBackTransaction();
            }
            finally
            {

                myCmdInsertBudget.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudgetCompanyMap.CommitTransaction();
                }

                connInsertClientReportTypeBudgetCompanyMap.Dispose();
                connInsertClientReportTypeBudgetCompanyMap.Close();

            }


            

        }

        public void InsertClientReportTypeBudgetIndividualAll(DataTable dtparam)
        {

            bool bResult = true;
            bool strResultBudget = true;

            ISDL.Connect connInsertClientReportTypeBudgetIndividual = new ISDL.Connect();
            connInsertClientReportTypeBudgetIndividual.setConnection("ocrsConnection");

            connInsertClientReportTypeBudgetIndividual.Open();
            connInsertClientReportTypeBudgetIndividual.BeginTransaction();

            SqlCommand myCmdInsertBudget = new SqlCommand();
            myCmdInsertBudget.Connection = connInsertClientReportTypeBudgetIndividual.Connection;

            try
            {
                if (bResult)
                {
                    myCmdInsertBudget.CommandText = "InsUpdClientReportBudgetIndividualAllRow";
                    myCmdInsertBudget.CommandType = CommandType.StoredProcedure;
                    myCmdInsertBudget.Transaction = connInsertClientReportTypeBudgetIndividual.currentTransaction;
                    myCmdInsertBudget.Parameters.Add("@Tablles", SqlDbType.Structured);
                    myCmdInsertBudget.Parameters["@Tablles"].Value = dtparam;
                    strResultBudget = connInsertClientReportTypeBudgetIndividual.cmdNoneQuery(myCmdInsertBudget);

                }
            }
            catch
            {
                bResult = false;
                connInsertClientReportTypeBudgetIndividual.RollBackTransaction();
            }
            finally
            {

                myCmdInsertBudget.Dispose();

                if (bResult)
                {
                    connInsertClientReportTypeBudgetIndividual.CommitTransaction();
                }

                connInsertClientReportTypeBudgetIndividual.Dispose();
                connInsertClientReportTypeBudgetIndividual.Close();

            }




        }

        /*Code Ends Here- Deepak*/
        /*Code Ends Here- Deepak*/

    }
}
