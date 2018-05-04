using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using ISDL;
using System.Data.SqlClient;
using System.Data;

namespace ISBL
{
    public class DashBoard
    {
        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        #endregion

        #region Public Variables
            
            public static readonly string  TYPE_COUNTRY = "Country";
            public static readonly string TYPE_REGION = "Region";
            public static readonly string TYPE_INDUSTRY = "Industry";
            public static readonly string TYPE_EMAILALERTS = "EmailAlerts";

        #endregion

        #region Constructors
        public DashBoard()
        {
            conn.setConnection("ocrsConnection");
        }
        #endregion

        //Dispose Conncetion
        public void DisposeConnection()
        {
            conn.Dispose();
        }

        #region Case Summary Widget Methods

        //Get Pie Chart1 - Risk Data
        public DataSet GetCaseRisk(string ClientCode, string LoginID, string AppName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetUserCaseRisk";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char, 50);
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters.Add("@AppName", SqlDbType.Char, 15);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@AppName"].Value = AppName;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetCaseRisk.sp_DbrdGetUserCaseRisk";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Pie Chart1 - Risk Data
        public DataSet GetCaseRisk(string ClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetClientCaseRisk";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetCaseRisk.sp_DbrdGetClientCaseRisk";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Pie Chart1 - On Time Data
        public DataSet GetOnTimeCase(string ClientCode, string LoginID, string AppName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetUserOnTimeCase";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char, 50);
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters.Add("@AppName", SqlDbType.Char, 15);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@AppName"].Value = AppName;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetOnTimeCase.sp_DbrdGetUserOnTimeCase";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Pie Chart1 - On Time Data
        public DataSet GetOnTimeCase(string ClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetClientOnTimeCase";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetOnTimeCase.sp_DbrdGetClientOnTimeCase";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }
        
        //Gets Case Summary
        public DataSet GetCaseSummary(string ClientCode, string LoginID, string AppName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetUserCaseSummary";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char,50);
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char,30);
            cmd.Parameters.Add("@AppName", SqlDbType.Char,15);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@AppName"].Value = AppName;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetCaseSummary.sp_DbrdGetUserCaseSummary";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Gets Case Summary
        public DataSet GetCaseSummary(string ClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetClientCaseSummary";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char,30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetCaseSummary.sp_DbrdGetClientCaseSummary";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get User Subjects
        public DataSet GetUserSubjects(string ClientCode, string LoginID, string AppName, string caseStatus)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            if (caseStatus == "WIP")
            {
                cmd.CommandText = "sp_DbrdGetUserWIPSubjects";
            }
            else
            {
                cmd.CommandText = "sp_DbrdGetUserClosedSubjects";
            }
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char, 50);
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters.Add("@AppName", SqlDbType.Char, 15);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@AppName"].Value = AppName;

            if (caseStatus != "WIP")
            {
                cmd.Parameters.Add("@Status", SqlDbType.Char, 10);
                cmd.Parameters["@Status"].Value = caseStatus;
            }

            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetUserSubjects";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Client Subjects
        public DataSet GetClientSubjects(string ClientCode, string caseStatus)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            if (caseStatus == "WIP")
            {
                cmd.CommandText = "sp_DbrdGetClientWIPSubjects";
            }
            else
            {
                cmd.CommandText = "sp_DbrdGetClientClosedSubjects";
            }
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;

            if (caseStatus != "WIP")
            {
                cmd.Parameters.Add("@Status", SqlDbType.Char, 10);
                cmd.Parameters["@Status"].Value = caseStatus;
            }

            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetClientSubjects";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Client Subjects
        public string GetClosedCaseNewsLink(string strCRN)
        {
            string strReturn;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_DbrdGetClosedCaseNewsLink";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar);
            myCmd.Parameters["@CRN"].Value = strCRN;

            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetClosedCaseNewsLink";
            strReturn = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strReturn;
        }



        #endregion

        #region GOC Widget Methods

        //Gets GOC Record Count
        public DataSet GetGOCRecordCount(Int16 currentMonthYear, Int16 currentMonth, Int16 lastMonthYear, Int16 lastMonth)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetGOCRecordCount";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@CurrentMonthYear", SqlDbType.SmallInt);
            cmd.Parameters.Add("@CurrentMonth", SqlDbType.SmallInt);
            cmd.Parameters.Add("@LastMonthYear", SqlDbType.SmallInt);
            cmd.Parameters.Add("@LastMonth", SqlDbType.SmallInt);
            cmd.Parameters["@CurrentMonthYear"].Value = currentMonthYear;
            cmd.Parameters["@CurrentMonth"].Value = currentMonth;
            cmd.Parameters["@LastMonthYear"].Value = lastMonthYear;
            cmd.Parameters["@LastMonth"].Value = lastMonth;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetGOCRecordCount";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Gets GOC New Datasets
        public DataSet GetNewGOCDatasets()
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetNewGOCDatasets";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetNewGOCDatasets";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        #endregion

        #region PPD Widget Methods

        //Gets PPD Risk
        public DataSet GetPPDRisk(string ClientCode, string LoginID, string AppName,Int16 StartMonth, Int16 EndMonth, Int16 Year)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetUserPartnerRisk";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char, 50);
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters.Add("@AppName", SqlDbType.Char, 15);
            cmd.Parameters.Add("@StartMonth", SqlDbType.TinyInt);
            cmd.Parameters.Add("@EndMonth", SqlDbType.TinyInt);
            cmd.Parameters.Add("@Year", SqlDbType.SmallInt);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters["@StartMonth"].Value = StartMonth;
            cmd.Parameters["@EndMonth"].Value = EndMonth;
            cmd.Parameters["@Year"].Value = Year;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetPPDRisk.sp_DbrdGetUserPartnerRisk";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Gets PPD Risk
        public DataSet GetPPDRisk(string ClientCode, Int16 StartMonth, Int16 EndMonth, Int16 Year)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetClientPartnerRisk";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters.Add("@StartMonth", SqlDbType.TinyInt);
            cmd.Parameters.Add("@EndMonth", SqlDbType.TinyInt);
            cmd.Parameters.Add("@Year", SqlDbType.SmallInt);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters["@StartMonth"].Value = StartMonth;
            cmd.Parameters["@EndMonth"].Value = EndMonth;
            cmd.Parameters["@Year"].Value = Year;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetPPDRisk.sp_DbrdGetClientPartnerRisk";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Converts month number to words
        public static String GetMonthName(Int16  month)
        {
            DateTime date = new DateTime(1900, month, 1);
            return date.ToString("MMM");  
        }

        #endregion

        #region Alert widget Methods

        //Get Subject Risk Summary
        public DataSet GetSubjectRiskSummary(string CRN, Int16 subjectID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_DbrdGetSubjectRiskSummary";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@CRN", SqlDbType.Char, 80);
            cmd.Parameters["@CRN"].Value = CRN;
            cmd.Parameters.Add("@SubjectID", SqlDbType.SmallInt);
            cmd.Parameters["@SubjectID"].Value = subjectID;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetSubjectRiskSummary";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get User Alert Settings
        public DataSet GetUserAlertSettings(string ClientCode, string LoginID, string AppName, string Type)
        { 
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;

            if (Type == "Country")
            {
                cmd.CommandText = "sp_DbrdGetUserAlertCountries";
            }
            else if (Type == "Region")
            {
                cmd.CommandText = "sp_DbrdGetUserAlertRegions";
            }
            else if (Type == "Industry")
            {
                cmd.CommandText = "sp_DbrdGetUserAlertIndustries";
            }

            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@LoginID", SqlDbType.Char, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@AppName", SqlDbType.Char, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@ClientCode", SqlDbType.Char, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetUserAlertSettings";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Alert Country Subjects
        public DataSet GetAlertSubjects(Int16 CountryID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdGetAlertCountrySubjects";
            cmd.Parameters.Add("@CountryID", SqlDbType.SmallInt);
            cmd.Parameters["@CountryID"].Value = CountryID ;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetAlertSubjects.sp_DbrdGetAlertCountrySubjects";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Alert Country Subjects
        public DataSet GetAlertSubjects(Int16 RegionID, string ClientCode, string LoginID, string AppName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdGetAlertRegionSubjects";
            cmd.Parameters.Add("@RegionID", SqlDbType.SmallInt);
            cmd.Parameters["@RegionID"].Value = RegionID ;
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50 );
            cmd.Parameters["@LoginID"].Value = LoginID ;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar , 15);
            cmd.Parameters["@AppName"].Value = AppName ;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetAlertSubjects.sp_DbrdGetAlertRegionSubjects";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Get Alert Country Subjects
        public DataSet GetAlertSubjects(Int32 IndustryID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdGetAlertIndustrySubjects";
            cmd.Parameters.Add("@IndustryID", SqlDbType.Int);
            cmd.Parameters["@IndustryID"].Value = IndustryID;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetAlertSubjects.sp_DbrdGetAlertIndustrySubjects";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }

        #endregion

        #region Alert Widget Configuration Methods

        //Gets Settings for Config Panel
        public DataSet GetConfigSettings(string ClientCode, string LoginID, string AppName, string Type)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;

            if (Type == "Country")
            {
                cmd.CommandText = "sp_DbrdGetUserConfigCountries";
            }
            else if (Type == "Region")
            {
                cmd.CommandText = "sp_DbrdGetUserConfigRegions";
            }
            else if (Type == "Industry")
            {
                cmd.CommandText = "sp_DbrdGetUserConfigIndustries";
            }
            else if (Type == "EmailAlerts")
            {
                cmd.CommandText = "sp_DbrdGetUserAlertEmails";
            }

            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetConfigSettings";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }


        //Gets Frequency List
        public DataSet GetFrequencyList()
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdGetFrequencyList";
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetFrequencyList";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }


        //Gets Email Alerts Settings
        public ArrayList GetEmailAlertSettings(string ClientCode, string LoginID, string AppName)
        {
            ArrayList data = new ArrayList(); 
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdGetUserConfigEmailAlerts";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;

            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.GetEmailAlertSettings";
            try
            {
                SqlDataReader dr = conn.cmdReader(cmd);
                if (dr != null)
                {
                    int i = 0;
                    while (dr.Read())
                    {
                        data.Add(dr["Activate"]);
                        data.Add(dr["FrequencyID"]);
                        i++;
                    }
                    if (i == 0)
                    {
                        data.Add(false);
                        data.Add(0);
                    }
                }
                else
                {
                    data.Add(false);
                    data.Add(0);
                }
            }
            catch
            {
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return data;
        }


        public Boolean SaveAlertConfig(String LoginID, String AppName,
            String ClientCode, String ClientName,
            String countries, String industries, System.Collections.ArrayList regionCountries,
            Boolean activate, short frequencyID, DataSet dsAlertEmails)
        {
            conn.Open();
            conn.callingMethod = "ISBL.DashBoard.SaveAlertConfig";
            conn.BeginTransaction();

            if (!CreateUserIfNotExists(LoginID ,AppName ,ClientCode, ClientName))
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            if (!DeleteAlertConfig (ClientCode, LoginID, AppName))
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            String[] countryID = countries.Split(new Char[] { ',' });
            Boolean success = true;
            if (countries != "")
            {
                for (int i = 0; i < countryID.Length; i++)
                {
                    Int16 iCountryID = Int16.Parse(countryID[i].ToString());
                    if (!InsertAlertCountry(ClientCode, LoginID, AppName, iCountryID))
                    {
                        success = false;
                        break;
                    }
                }
            }
            if (!success)
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            String[] industryID = industries.Split(new Char[] { ',' });
            success = true;
            if (industries != "")
            {
                for (int i = 0; i < industryID.Length; i++)
                {
                    Int32 iIndustryID = Int32.Parse(industryID[i].ToString());
                    if (!InsertAlertIndustry(ClientCode, LoginID, AppName, iIndustryID))
                    {
                        success = false;
                        break;
                    }
                }
            }
            if (!success)
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            success = true;
            for (int i = 0; i < regionCountries.Count; i++)
            {
                System.Collections.ArrayList row = (System.Collections.ArrayList)regionCountries[i];
                if (!InsertAlertRegionCountry(ClientCode, LoginID, AppName, (Int16)row[1], (Int16)row[0]))
                {
                    success = false;
                    break;
                }
            }
            if (!success)
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            if (!SaveAlertEmailConfig(ClientCode, LoginID, AppName, activate, frequencyID))
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }  

            success = true;
            for (int i = 0; i < dsAlertEmails.Tables[0].Rows.Count; i++)
            {
                DataRow row = dsAlertEmails.Tables[0].Rows[i];

                try
                {
                    if (!InsertAlertEmail(ClientCode, LoginID, AppName, (Guid)row["AddressID"], row["EmailAddress"].ToString()))
                    {
                        success = false;
                        break;
                    }
                }
                catch (DeletedRowInaccessibleException)
                {
                    if (!DeleteAlertEmail(ClientCode, LoginID, AppName, (Guid)row["AddressID", DataRowVersion.Original]))
                    {
                        success = false;    
                        break;
                    }
                }
            }
            if (!success)
            {
                conn.RollBackTransaction();
                conn.Close();
                return false;
            }

            conn.CommitTransaction();
            conn.Close();
            return true;
        }

        //Dealetes Current User Settings
        private Boolean DeleteAlertConfig(String ClientCode, String LoginID, String AppName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdDeleteAlertConfig";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Creates user for saving settings if user doesnt exist in dashboard's user master table
        private Boolean CreateUserIfNotExists(String LoginID, String AppName, String ClientCode, String ClientName)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdCreateUserIfNotExists";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode ;
            cmd.Parameters.Add("@ClientName", SqlDbType.NVarChar, 100);
            cmd.Parameters["@ClientName"].Value = ClientName ;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Inserts User Selected Country from Config Panel
        private Boolean InsertAlertCountry(String ClientCode, String LoginID, String AppName, Int16 CountryID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdInsertAlertCountry";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@CountryID", SqlDbType.SmallInt);
            cmd.Parameters["@CountryID"].Value = CountryID;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Inserts User Selected Region Country from Config Panel
        private Boolean InsertAlertRegionCountry(String ClientCode, String LoginID, String AppName, Int16 CountryID, Int16 RegionID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdInsertAlertRegionCountry";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@CountryID", SqlDbType.SmallInt );
            cmd.Parameters["@CountryID"].Value = CountryID;
            cmd.Parameters.Add("@RegionID", SqlDbType.SmallInt );
            cmd.Parameters["@RegionID"].Value = RegionID;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Inserts User Selected Industry from Config Panel
        private Boolean InsertAlertIndustry(String ClientCode, String LoginID, String AppName, Int32 IndustryID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdInsertAlertIndustry";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@IndustryID", SqlDbType.Int);
            cmd.Parameters["@IndustryID"].Value = IndustryID;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Saves User's EMail alerts setting from Config Panel
        private Boolean SaveAlertEmailConfig(String ClientCode, String LoginID, String AppName, Boolean Activate, short FrequencyID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdSaveAlertEmailConfig";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@Activate", SqlDbType.Bit );
            cmd.Parameters["@Activate"].Value = Activate;
            cmd.Parameters.Add("@FrequencyID", SqlDbType.SmallInt);
            cmd.Parameters["@FrequencyID"].Value = FrequencyID;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Inserts User Added Alert Email from Config Panel
        private Boolean InsertAlertEmail(String ClientCode, String LoginID, String AppName, Guid AddressID, String EmailID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdInsertAlertEmail";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@AddressID", SqlDbType.UniqueIdentifier );
            cmd.Parameters["@AddressID"].Value = AddressID ;
            cmd.Parameters.Add("@EmailAddress", SqlDbType.VarChar, 512);
            cmd.Parameters["@EmailAddress"].Value = EmailID;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        //Deletes User Removed Alert Email from Config Panel
        private Boolean DeleteAlertEmail(String ClientCode, String LoginID, String AppName, Guid AddressID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.CommandText = "sp_DbrdDeleteAlertEmail";
            cmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 50);
            cmd.Parameters["@LoginID"].Value = LoginID;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            cmd.Parameters.Add("@AppName", SqlDbType.VarChar, 15);
            cmd.Parameters["@AppName"].Value = AppName;
            cmd.Parameters.Add("@AddressID", SqlDbType.UniqueIdentifier );
            cmd.Parameters["@AddressID"].Value = AddressID ;
            cmd.Transaction = conn.currentTransaction;
            String strStatus = conn.cmdScalarStoredProc(cmd);
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

        #endregion
    }
}