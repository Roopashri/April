using System;
using System.Collections.Generic;
using System.Collections;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.IO;
using System.Drawing;
using System.Drawing.Imaging;
using System.Data.OleDb;
using System.Threading;
using System.Xml;
using System.Configuration;
using System.Web.UI;
using System.Security.AccessControl;
using System.Security.Principal;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
/* For Diagnostics */
using System.Diagnostics;
namespace ISBL
{
    public class BulkOrder
    {
        const string ExcelPlaceOrderMsg = "Excel Bulk Order - Creation";
        const string WebPlaceOrderMsg = "Web Bulk Order - Creation";
        const string ExcelEmailMsg = "Excel Bulk Order - Email";
        const string WebEmailMsg = "Web Bulk Order - Email";
        Boolean isExcel;
        private Boolean blnExcelFailed = false;
        ISBL.General oISBLGen = new ISBL.General();
        ISBL.User oISBLUser = new ISBL.User();
        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        #endregion

        #region Constructors
        public BulkOrder()
        {
            conn.setConnection("ocrsConnection");
        }
        #endregion

        #region Public Member

        //Dispose Conncetion
        public void DisposeConnection()
        {
            conn.Dispose();
        }


        //Validates xml data against xsd file
        static public void ValidateXML(string XML, string XSDFile)
        {
            try
            {
                StringReader XMLStringReader = new StringReader(XML);
                XmlReaderSettings Settings = new XmlReaderSettings();
                Settings.ValidationType = ValidationType.Schema;
                Settings.Schemas.Add("", XSDFile);
                XmlReader objXmlReader = XmlReader.Create(XMLStringReader, Settings);
                while (objXmlReader.Read())
                {
                    //Do Nothing
                }
            }
            catch (Exception e)
            {
                throw e;
            }
        }

        //Generates Internal Error XML Reponse
        public String getInternalError()
        {
            XmlDocument xmlDoc = new XmlDocument();
            XmlNode result, status;
            result = xmlDoc.CreateElement("result");
            status = xmlDoc.CreateElement("status");
            status.InnerText = "Internal Error";
            result.AppendChild(status);
            xmlDoc.AppendChild(result);
            return xmlDoc.InnerXml;
        }

        //Place Order Via Excel Sheet
        public String PlaceOrderViaExcel(ISBL.User oUser, String XMLData)
        {
            bool success = true;
            ISBL.General oGeneral = new ISBL.General();
            try
            {
                ValidateXML(XMLData, System.Configuration.ConfigurationManager.AppSettings["ExcelOrderXSD"].ToString());
            }
            catch (Exception ex)
            {
                if (XMLData.Length > 2048000) XMLData = XMLData.Substring(0, 204800);
                oGeneral.SaveEventLog(Guid.NewGuid().ToString(), oUser.ClientCode, "", ExcelPlaceOrderMsg, XMLData, "", DateTime.Now, DateTime.Now, oUser.LoginID, oUser.UserType, "EBO_0", "XML Validation Failed: " + ex.Message);
                return getInternalError();
            }

            DataSet dsOrders = new DataSet();
            dsOrders.ReadXml(new StringReader(XMLData));
            DataView dvOrders = dsOrders.Tables[1].DefaultView;
            dvOrders.Sort = "CaseNo";
            DataTable dtOrders = new DataTable();
            dtOrders = dvOrders.ToTable();
            //dtOrders = dsOrders.Tables[1];  

            XmlDocument xmlResponse = new XmlDocument();
            XmlNode xmlResult = xmlResponse.CreateElement("result");
            xmlResponse.AppendChild(xmlResult);

            XmlNode transactionStatus = xmlResponse.CreateElement("status");
            xmlResult.AppendChild(transactionStatus);

            string strReportType = dsOrders.Tables[0].Rows[0]["ReportType"].ToString();
            //Start Aug 09 Adam: to cater for old and new excel bulk order for report type name change
            switch (strReportType)
            {   
                case "IS Lite" :
                    strReportType = "IEDD1";
                    break;
                case "IS Lite (IEDD1)" :
                    strReportType = "IEDD1";
                    break;
                case "IS Report" :
                    strReportType = "IEDD2";
                    break;
                case "IS Report (IEDD2)" :
                    strReportType = "IEDD2";
                    break;
                case "IS Premium" :
                    strReportType = "IPDD";
                    break;
                case "IS Premium (IPDD)" :
                    strReportType = "IPDD";
                    break;
                case "IS ORR" :
                    strReportType = "ORR";
                    break;
            }
            //End Aug 09 Adam: to cater for old and new excel bulk order for report type name change
            XmlNode xmlReportType = xmlResponse.CreateElement("ReportType");
            xmlResult.AppendChild(xmlReportType);

            ISBL.Admin oAdmin = new ISBL.Admin();
            DataSet reportTypeMaster = oAdmin.GetReportTypeMasterAdmin(); //Jul 2009 Adam - Change Report Type Description Enhancement
            if (reportTypeMaster != null)
            {
                DataRow[] reportTypes = reportTypeMaster.Tables[0].Select("ReportType='" + strReportType + "'");
                if (reportTypes.Length == 0)
                {
                    success = false;
                    xmlReportType.InnerText = "Invalid";
                }
                else
                {
                    xmlReportType.InnerText = "Valid";
                }
            }
            else
            {
                return getInternalError();
            }

            DataSet dsSubjectTypeList = this.GetSubjectType();
            if (dsSubjectTypeList == null)
            {
                return getInternalError();
            }

            DataSet dsCountryList = oGeneral.GetCountryList();
            if (dsCountryList == null)
            {
                return getInternalError();
            }

            XmlNode xmlRows = xmlResponse.CreateElement("rows");
            xmlResult.AppendChild(xmlRows);

            //Get Unique Cases
            Hashtable htCaseNos = new Hashtable();
            foreach (DataRow dr in dtOrders.Rows)
            {
                if (!htCaseNos.Contains(dr["CaseNo"]))
                {
                    htCaseNos.Add(dr["CaseNo"], null);
                }
            }

            ISDL.Connect connPlaceOrder = new ISDL.Connect();
            connPlaceOrder.setConnection("ocrsConnection");

            connPlaceOrder.Open();
            connPlaceOrder.BeginTransaction();

            //Persist Bulk Order SessionID in db
            Guid BulkSessionID = Guid.NewGuid();
            if (!SaveClientBulkOrderMaster(ref connPlaceOrder, ref oUser, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                return getInternalError();
            }

            //Create Excel Attachment File
            // Try catch part added by sanjeeva reddy on 18/06/2015
            String strAttachmentFilePath = System.Configuration.ConfigurationManager.AppSettings["excelAttachmentDir"].ToString();
            String strAttachmentFile = Guid.NewGuid().ToString() + ".xls";
                           
                if (!CreateExcelFile(strAttachmentFilePath + strAttachmentFile))
                {
                    blnExcelFailed = false;                    
                }
            
            //Loop through unique cases
            IDictionaryEnumerator Enum = htCaseNos.GetEnumerator();
            int caseCount = 0;
            while (Enum.MoveNext())
            {
                caseCount++;
                string CaseNo = (string)Enum.Key;
                ArrayList subjectRows = new ArrayList();
                DataRow[] subject = dtOrders.Select("CaseNo='" + CaseNo + "'");
                for (int i = 0; i < subject.Length; i++)
                {
                    XmlNode xmlRow = xmlResponse.CreateElement("row");
                    XmlNode xmlRowNum = xmlResponse.CreateElement("rownum");
                    xmlRowNum.InnerText = subject[i]["RowNum"].ToString();
                    xmlRow.AppendChild(xmlRowNum);
                    XmlNode cols = xmlResponse.CreateElement("cols");
                    xmlRow.AppendChild(cols);

                    XmlNode col;

                    String country = "";
                    //Jul 09 Adam encode country with single quote
                    DataRow[] countries = dsCountryList.Tables[0].Select("Description='" + subject[i]["Country"].ToString().Replace("'","''") + "'");
                    if (countries.Length == 0)
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "4";
                        cols.AppendChild(col);
                    }
                    else
                    {
                        country = countries[0]["Country"].ToString();
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    string[] arrtype = subject[i]["Type"].ToString().Split(new char[] {';'});
                    DataRow[] drSubjectList = dsSubjectTypeList.Tables[0].Select("SubjectTypeDescription = '" + arrtype[0].ToString() + "'");
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist

                    String subjectTypeID = "";
                    if (drSubjectList.Length > 0)
                    {
                        subjectTypeID = drSubjectList[0]["SubjectType"].ToString();
                    }
                    else
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "2";
                        cols.AppendChild(col);
                    }

                    if (subject[i]["Name"].ToString() == "")
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "3";
                        cols.AppendChild(col);
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    if (System.Configuration.ConfigurationManager.AppSettings["MakeRefNoMandatoryByClient"].ToString().Contains(oUser.ClientCode))
                    {
                        if (arrtype.Length > 1)
                        {
                            if (subject[i]["CRN"].ToString().Trim().Length == 0)
                            {
                                success = false;
                                col = xmlResponse.CreateElement("col");
                                col.InnerText = "1";
                                cols.AppendChild(col);
                            }
                        }
                    }
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory for Tyco Entity

                    if (cols.ChildNodes.Count > 0) xmlRows.AppendChild(xmlRow);

                    if (success)
                    {
                        Hashtable subjectRow = new Hashtable();
                        subjectRow.Add("SubjectType", subjectTypeID);
                        subjectRow.Add("SubjectName", subject[i]["Name"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountry", country);
                        subjectRow.Add("OtherDetails", subject[i]["Details"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ClientRefNum", subject[i]["CRN"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountryDesc", subject[i]["Country"].ToString());

                        subjectRows.Add(subjectRow);
                    }
                }

                if (success)
                {
                    if (!SaveCase(ref connPlaceOrder, oUser, subjectRows, strReportType, strAttachmentFilePath , strAttachmentFile, true,BulkSessionID))
                    {
                        try
                        {
                            connPlaceOrder.RollBackTransaction();
                        }
                        catch { }
                        connPlaceOrder.Dispose();
                        connPlaceOrder = null;
                        return getInternalError();
                    }
                }
            }

            if (success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.CommitTransaction();
                }
                catch { }
                transactionStatus.InnerText = "Success";
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

                Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails(strAttachmentFilePath + strAttachmentFile, oUser.ClientCode, oUser.ClientName, oUser.LoginIDEmail, strReportType, BulkSessionID, true); });
                threadSendMails.IsBackground = true;
                threadSendMails.Start();
            }
            else if (!success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                transactionStatus.InnerText = "Failed";
            }

            return xmlResponse.InnerXml;
        }

        //Place Order Via Web Interface
        public bool PlaceOrderViaWeb(ISBL.User oISBLUser, ArrayList lstSubjects, String strReportType, String strAttachmentFilePath, Boolean OnBehalf, string OnBehalfUserID)
        {
            ISDL.Connect connPlaceOrder = new ISDL.Connect();
            String OnBehalfUserIDEmail = "";
            String OriginalLoginIDEmail = "";
            Boolean IsOnBehalf = false;

            connPlaceOrder.setConnection("ocrsConnection");
            
            connPlaceOrder.Open();
            connPlaceOrder.BeginTransaction();

            //Persist Bulk Order SessionID in db
            Guid BulkSessionID = Guid.NewGuid();
            if (!SaveClientBulkOrderMaster(ref connPlaceOrder, ref oISBLUser, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                return false;
            }

            //Create Excel Attachment File
            // Try catch part added by sanjeeva reddy on 18/06/2015
            String strAttachmentFile = Guid.NewGuid().ToString() + ".xls";
                           
                if (!CreateExcelFile(strAttachmentFilePath + strAttachmentFile))
                {
                    //blnExcelFailed = true; 
                     blnExcelFailed = true;                  
                }           

            //Place Order

            if (OnBehalf)
            {
                ISBL.Admin oAdmin = new ISBL.Admin();
                DataSet dsUser = new DataSet();
                dsUser = oAdmin.GetUserDetails(OnBehalfUserID);
                OnBehalfUserIDEmail = dsUser.Tables[0].Rows[0]["LoginIDEmail"].ToString();

                oISBLUser.IsImpersonate = true;
                oISBLUser.ImpersonateLoginID = oISBLUser.LoginID;
                oISBLUser.LoginID = OnBehalfUserID;
                OriginalLoginIDEmail = oISBLUser.LoginIDEmail;
                oISBLUser.LoginIDEmail = OnBehalfUserIDEmail;
                IsOnBehalf = true;
            }

            if (SaveCase(ref connPlaceOrder, oISBLUser, lstSubjects, strReportType, strAttachmentFilePath, strAttachmentFile, false, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.CommitTransaction();
                    //added by snajeeva on 19/10/2015 for bulk order email issue start
                    DataSet dsCasedet = new DataSet();
                    dsCasedet = GetCaseorderDetails(BulkSessionID);
                    //end of email issue sanjeeva
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

                if (IsOnBehalf)
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails(strAttachmentFilePath + strAttachmentFile, oISBLUser.ClientCode, oISBLUser.ClientName, OnBehalfUserIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();
                }
                else
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails(strAttachmentFilePath + strAttachmentFile, oISBLUser.ClientCode, oISBLUser.ClientName, oISBLUser.LoginIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();
                }

                 if (OnBehalf) //Revert to Orginal
                {
                    oISBLUser.IsImpersonate = false;
                    oISBLUser.LoginID = oISBLUser.ImpersonateLoginID;
                    oISBLUser.ImpersonateLoginID = "";
                    oISBLUser.LoginIDEmail = OriginalLoginIDEmail;
                }
                
                return true;
            }
            else
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

                if (OnBehalf) //Revert to Orginal
                {
                    oISBLUser.IsImpersonate = false;
                    oISBLUser.LoginID = oISBLUser.ImpersonateLoginID;
                    oISBLUser.ImpersonateLoginID = "";
                    oISBLUser.LoginIDEmail = OriginalLoginIDEmail;
                }
                return false;
            }
        }
        //added by sanjeeva for bulk order email issue start
        public DataSet GetCaseorderDetails(Guid strBulkMasterid)
        {
            string strSQL;
            int intCompare;
            int intIsBDM;
            Boolean blAnd = false;
            ISBL.General oISBLGen = new ISBL.General();
            strSQL = "select  b.id OrderID,b.clientreferencenumber ClientReferenceNumber,b.reporttype ReportType,case when c.subjecttype=1 then 'Company' else 'Individual' end SubjectType,c.subjectname SubjectName,c.country SubjectCountry,case when c.[Primary]=1 then 'Yes' else 'No' end PrimarySubject,c.otherdetails OtherDetails from ocrsclientbulkorder a,ocrsclientorder b,ocrsclientordersubject c where a.orderid=b.id and a.orderid=c.orderid and a.bulkmasterid='" + strBulkMasterid + "'";
            
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
        //end of email issue sanjeeva
        //Place Order Via Web with Case No
        public bool PlaceOrderViaWebCaseNo(ISBL.User oISBLUser, String XMLData, String strReportType, String strAttachmentFilePath, Boolean OnBehalf, string OnBehalfUserID)
        {
            bool success = true;
            ISBL.General oGeneral = new ISBL.General();
            String OnBehalfUserIDEmail = "";
            String OriginalLoginIDEmail = "";
            Boolean IsOnBehalf = false;
            
            try
            {
                ValidateXML(XMLData, System.Configuration.ConfigurationManager.AppSettings["ExcelOrderXSD"].ToString());
            }
            catch (Exception ex)
            {
                if (XMLData.Length > 2048000) XMLData = XMLData.Substring(0, 204800);
                oGeneral.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", ExcelPlaceOrderMsg, XMLData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, "EBO_0", "XML Validation Failed: " + ex.Message);
                return false;
            }

            DataSet dsOrders = new DataSet();
            dsOrders.ReadXml(new StringReader(XMLData));
            DataView dvOrders = dsOrders.Tables[1].DefaultView;
            dvOrders.Sort = "CaseNo";
            DataTable dtOrders = new DataTable();
            dtOrders = dvOrders.ToTable();

            XmlDocument xmlResponse = new XmlDocument();
            XmlNode xmlResult = xmlResponse.CreateElement("result");
            xmlResponse.AppendChild(xmlResult);

            XmlNode transactionStatus = xmlResponse.CreateElement("status");
            xmlResult.AppendChild(transactionStatus);

            XmlNode xmlReportType = xmlResponse.CreateElement("ReportType");
            xmlResult.AppendChild(xmlReportType);

            DataSet dsSubjectTypeList = this.GetSubjectType();
            if (dsSubjectTypeList == null)
            {
                return false;
            }

            DataSet dsCountryList = oGeneral.GetCountryList();
            if (dsCountryList == null)
            {
                return false;
            }

            XmlNode xmlRows = xmlResponse.CreateElement("rows");
            xmlResult.AppendChild(xmlRows);

            //Get Unique Cases
            Hashtable htCaseNos = new Hashtable();
            foreach (DataRow dr in dtOrders.Rows)
            {
                if (!htCaseNos.Contains(dr["CaseNo"]))
                {
                    htCaseNos.Add(dr["CaseNo"], null);
                }
            }

            ISDL.Connect connPlaceOrder = new ISDL.Connect();
            connPlaceOrder.setConnection("ocrsConnection");

            connPlaceOrder.Open();
            connPlaceOrder.BeginTransaction();

            //Persist Bulk Order SessionID in db
            Guid BulkSessionID = Guid.NewGuid();
            if (!SaveClientBulkOrderMaster(ref connPlaceOrder, ref oISBLUser, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                return false;
            }

            String strAttachmentFile = BulkSessionID.ToString() + ".xls";
            if (OnBehalf)
            {
                ISBL.Admin oAdmin = new ISBL.Admin();
                DataSet dsUser = new DataSet();
                dsUser = oAdmin.GetUserDetails(OnBehalfUserID);
                OnBehalfUserIDEmail = dsUser.Tables[0].Rows[0]["LoginIDEmail"].ToString();

                oISBLUser.IsImpersonate = true;
                oISBLUser.ImpersonateLoginID = oISBLUser.LoginID;
                oISBLUser.LoginID = OnBehalfUserID;
                OriginalLoginIDEmail = oISBLUser.LoginIDEmail;
                oISBLUser.LoginIDEmail = OnBehalfUserIDEmail;
                IsOnBehalf = true;
            }

            //Loop through unique cases
            IDictionaryEnumerator Enum = htCaseNos.GetEnumerator();
            int caseCount = 0;
            while (Enum.MoveNext())
            {
                caseCount++;
                string CaseNo = (string)Enum.Key;
                ArrayList subjectRows = new ArrayList();
                DataRow[] subject = dtOrders.Select("CaseNo='" + CaseNo + "'");
                for (int i = 0; i < subject.Length; i++)
                {
                    XmlNode xmlRow = xmlResponse.CreateElement("row");
                    XmlNode xmlRowNum = xmlResponse.CreateElement("rownum");
                    xmlRowNum.InnerText = subject[i]["RowNum"].ToString();
                    xmlRow.AppendChild(xmlRowNum);
                    XmlNode cols = xmlResponse.CreateElement("cols");
                    xmlRow.AppendChild(cols);

                    XmlNode col;

                    String country = "";
                    //Jul 09 Adam encode country with single quote
                    DataRow[] countries = dsCountryList.Tables[0].Select("Description='" + subject[i]["Country"].ToString().Replace("'", "''") + "'");
                    if (countries.Length == 0)
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "4";
                        cols.AppendChild(col);
                    }
                    else
                    {
                        country = countries[0]["Country"].ToString();
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    string[] arrtype = subject[i]["Type"].ToString().Split(new char[] { ';' });
                    DataRow[] drSubjectList = dsSubjectTypeList.Tables[0].Select("SubjectTypeDescription = '" + arrtype[0].ToString() + "'");
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist

                    String subjectTypeID = "";
                    if (drSubjectList.Length > 0)
                    {
                        subjectTypeID = drSubjectList[0]["SubjectType"].ToString();
                    }
                    else
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "2";
                        cols.AppendChild(col);
                    }

                    if (subject[i]["Name"].ToString() == "")
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "3";
                        cols.AppendChild(col);
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    if (System.Configuration.ConfigurationManager.AppSettings["MakeRefNoMandatoryByClient"].ToString().Contains(oISBLUser.ClientCode))
                    {
                        if (arrtype.Length > 1)
                        {
                            if (subject[i]["CRN"].ToString().Trim().Length == 0)
                            {
                                success = false;
                                col = xmlResponse.CreateElement("col");
                                col.InnerText = "1";
                                cols.AppendChild(col);
                            }
                        }
                    }
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory for Tyco Entity

                    if (cols.ChildNodes.Count > 0) xmlRows.AppendChild(xmlRow);

                    if (success)
                    {
                        Hashtable subjectRow = new Hashtable();
                        subjectRow.Add("SubjectType", subjectTypeID);
                        subjectRow.Add("SubjectName", subject[i]["Name"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountry", country);
                        subjectRow.Add("OtherDetails", subject[i]["Details"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ClientRefNum", subject[i]["CRN"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountryDesc", subject[i]["Country"].ToString());
                        subjectRow.Add("SubReportType", subject[i]["SubReportType"].ToString());
                        subjectRows.Add(subjectRow);
                    }
                }

                if (success)
                {
                    if (!SaveCase(ref connPlaceOrder, oISBLUser, subjectRows, strReportType, strAttachmentFilePath, strAttachmentFile, true, BulkSessionID))
                    {
                        try
                        {
                            connPlaceOrder.RollBackTransaction();
                        }
                        catch { }
                        connPlaceOrder.Dispose();
                        connPlaceOrder = null;

                        return false;
                    }
                }
            }

            if (success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.CommitTransaction();
                    if (!CreateExcelFilenew(strAttachmentFilePath + strAttachmentFile,BulkSessionID))
                    {
                        //blnExcelFailed = true;                    
                        blnExcelFailed = false;                    
                    }
                }
                catch { }
                transactionStatus.InnerText = "Success";
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

                if (IsOnBehalf)
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails("", oISBLUser.ClientCode, oISBLUser.ClientName, OnBehalfUserIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();

                    //Revert to Orginal
                    oISBLUser.IsImpersonate = false;
                    oISBLUser.LoginID = oISBLUser.ImpersonateLoginID;
                    oISBLUser.ImpersonateLoginID = "";
                    oISBLUser.LoginIDEmail = OriginalLoginIDEmail;
                }
                else
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails("", oISBLUser.ClientCode, oISBLUser.ClientName, oISBLUser.LoginIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();
                }

            }
            else if (!success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                transactionStatus.InnerText = "Failed";
            }

            return success;
        }
        //************************************************************************JPMC Changes***************************************************************
        public bool PlaceOrderViaWebCaseNo_jpmc(ISBL.User oISBLUser, String XMLData, String strReportType, String strAttachmentFilePath, Boolean OnBehalf, string OnBehalfUserID)
        {
            bool success = true;
            ISBL.General oGeneral = new ISBL.General();
            String OnBehalfUserIDEmail = "";
            String OriginalLoginIDEmail = "";
            Boolean IsOnBehalf = false;
            String SubReportTypeID = "";
            String strSubjectAliases = "";
            String strAdditionalInformation = "";
            String strADDRESS = "";
            String strID = "";
            String strDOB = "";
           
            try
            {
                ValidateXML(XMLData, System.Configuration.ConfigurationManager.AppSettings["JpmcExcelOrderXSD"].ToString());
            }
            catch (Exception ex)
            {
                if (XMLData.Length > 2048000) XMLData = XMLData.Substring(0, 204800);
                oGeneral.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", ExcelPlaceOrderMsg, XMLData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, "EBO_0", "XML Validation Failed: " + ex.Message);
                return false;
            }

            DataSet dsOrders = new DataSet();
            dsOrders.ReadXml(new StringReader(XMLData));
            DataView dvOrders = dsOrders.Tables[1].DefaultView;
            dvOrders.Sort = "CaseNo";
            DataTable dtOrders = new DataTable();
            dtOrders = dvOrders.ToTable();

            XmlDocument xmlResponse = new XmlDocument();
            XmlNode xmlResult = xmlResponse.CreateElement("result");
            xmlResponse.AppendChild(xmlResult);

            XmlNode transactionStatus = xmlResponse.CreateElement("status");
            xmlResult.AppendChild(transactionStatus);

            XmlNode xmlReportType = xmlResponse.CreateElement("ReportType");
            xmlResult.AppendChild(xmlReportType);

            DataSet dsSubjectTypeList = this.GetSubjectType();
            if (dsSubjectTypeList == null)
            {
                return false;
            }

            DataSet dsCountryList = oGeneral.GetCountryList();
            if (dsCountryList == null)
            {
                return false;
            }

            XmlNode xmlRows = xmlResponse.CreateElement("rows");
            xmlResult.AppendChild(xmlRows);

            //Get Unique Cases
            Hashtable htCaseNos = new Hashtable();
            foreach (DataRow dr in dtOrders.Rows)
            {
                if (!htCaseNos.Contains(dr["CaseNo"]))
                {
                    htCaseNos.Add(dr["CaseNo"], null);
                }
            }

            ISDL.Connect connPlaceOrder = new ISDL.Connect();
            connPlaceOrder.setConnection("ocrsConnection");

            connPlaceOrder.Open();
            connPlaceOrder.BeginTransaction();

            //Persist Bulk Order SessionID in db
            Guid BulkSessionID = Guid.NewGuid();
            if (!SaveClientBulkOrderMaster(ref connPlaceOrder, ref oISBLUser, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                return false;
            }

            String strAttachmentFile = BulkSessionID.ToString() + ".xls";
            if (OnBehalf)
            {
                ISBL.Admin oAdmin = new ISBL.Admin();
                DataSet dsUser = new DataSet();
                dsUser = oAdmin.GetUserDetails(OnBehalfUserID);
                OnBehalfUserIDEmail = dsUser.Tables[0].Rows[0]["LoginIDEmail"].ToString();

                oISBLUser.IsImpersonate = true;
                oISBLUser.ImpersonateLoginID = oISBLUser.LoginID;
                oISBLUser.LoginID = OnBehalfUserID;
                OriginalLoginIDEmail = oISBLUser.LoginIDEmail;
                oISBLUser.LoginIDEmail = OnBehalfUserIDEmail;
                IsOnBehalf = true;
            }

            //Loop through unique cases
            IDictionaryEnumerator Enum = htCaseNos.GetEnumerator();
            int caseCount = 0;
            while (Enum.MoveNext())
            {
                caseCount++;
                string CaseNo = (string)Enum.Key;
                ArrayList subjectRows = new ArrayList();
                DataRow[] subject = dtOrders.Select("CaseNo='" + CaseNo + "'");
                for (int i = 0; i < subject.Length; i++)
                {
                    XmlNode xmlRow = xmlResponse.CreateElement("row");
                    XmlNode xmlRowNum = xmlResponse.CreateElement("rownum");
                    xmlRowNum.InnerText = subject[i]["RowNum"].ToString();
                    xmlRow.AppendChild(xmlRowNum);
                    XmlNode cols = xmlResponse.CreateElement("cols");
                    xmlRow.AppendChild(cols);

                    XmlNode col;

                    String country = "";
                    //Jul 09 Adam encode country with single quote
                    DataRow[] countries = dsCountryList.Tables[0].Select("Description='" + subject[i]["Country"].ToString().Replace("'", "''") + "'");
                    if (countries.Length == 0)
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "4";
                        cols.AppendChild(col);
                    }
                    else
                    {
                        country = countries[0]["Country"].ToString();
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    string[] arrtype = subject[i]["Type"].ToString().Split(new char[] { ';' });
                    DataRow[] drSubjectList = dsSubjectTypeList.Tables[0].Select("SubjectTypeDescription = '" + arrtype[0].ToString() + "'");
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist

                    String subjectTypeID = "";
                    if (drSubjectList.Length > 0)
                    {
                        subjectTypeID = drSubjectList[0]["SubjectType"].ToString();
                    }
                    else
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "2";
                        cols.AppendChild(col);
                    }

                    if (subject[i]["Name"].ToString() == "")
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "3";
                        cols.AppendChild(col);
                    }
                    SubReportTypeID = subject[i]["SubreportType"].ToString();
                    strSubjectAliases = subject[i]["SubjectAliases"].ToString();
                    strDOB = subject[i]["DOB"].ToString();
                    strID = subject[i]["ID"].ToString();
                    strADDRESS = subject[i]["ADDRESS"].ToString();
                    strAdditionalInformation = subject[i]["AdditionalInformation"].ToString();
                    

                    //if (drSubjectList.Length > 0)
                    //{
                    //    SubReportTypeID = drSubjectList[0]["SubreportType"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "7";
                    //    cols.AppendChild(col);
                    //}

                    ///////////////
                    //if (drSubjectList.Length > 0)
                    //{
                    //    strSubjectAliases = drSubjectList[0]["SubjectAliases"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "6";
                    //    cols.AppendChild(col);
                    //}
                    //if (drSubjectList.Length > 0)
                    //{
                    //    strDOB = drSubjectList[0]["DOB"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "9";
                    //    cols.AppendChild(col);
                    //}

                    //if (drSubjectList.Length > 0)
                    //{
                    //    strID = drSubjectList[0]["ID"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "10";
                    //    cols.AppendChild(col);
                    //}
                    //if (drSubjectList.Length > 0)
                    //{
                    //    strADDRESS = drSubjectList[0]["ADDRESS"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "11";
                    //    cols.AppendChild(col);
                    //}
                    //if (drSubjectList.Length > 0)
                    //{
                    //    strAdditionalInformation = drSubjectList[0]["AdditionalInformation"].ToString();
                    //}
                    //else
                    //{
                    //    success = false;
                    //    col = xmlResponse.CreateElement("col");
                    //    col.InnerText = "12";
                    //    cols.AppendChild(col);
                    //}
                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    if (System.Configuration.ConfigurationManager.AppSettings["MakeRefNoMandatoryByClient"].ToString().Contains(oISBLUser.ClientCode))
                    {
                        if (arrtype.Length > 1)
                        {
                            if (subject[i]["CRN"].ToString().Trim().Length == 0)
                            {
                                success = false;
                                col = xmlResponse.CreateElement("col");
                                col.InnerText = "1";
                                cols.AppendChild(col);
                            }
                        }
                    }
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory for Tyco Entity

                    if (cols.ChildNodes.Count > 0) xmlRows.AppendChild(xmlRow);

                    if (success)
                    {
                        Hashtable subjectRow = new Hashtable();
                        subjectRow.Add("SubjectType", subjectTypeID);
                        subjectRow.Add("SubjectName", subject[i]["Name"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountry", country);
                        subjectRow.Add("SubreportType", SubReportTypeID.ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectAliases", strSubjectAliases);
                        subjectRow.Add("DOB", strDOB.ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ID", strID.ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ADDRESS", strADDRESS.ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        //subjectRow.Add("AdditionalInformation", strAdditionalInformation);
                        subjectRow.Add("AdditionalInformation", strAdditionalInformation.ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ClientRefNum", subject[i]["CRN"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountryDesc", subject[i]["Country"].ToString());
                        
                        subjectRows.Add(subjectRow);
                    }
                }

                if (success)
                {
                    if (!SaveCase(ref connPlaceOrder, oISBLUser, subjectRows, strReportType, strAttachmentFilePath, strAttachmentFile, true, BulkSessionID))
                    {
                        try
                        {
                            connPlaceOrder.RollBackTransaction();
                        }
                        catch { }
                        connPlaceOrder.Dispose();
                        connPlaceOrder = null;

                        return false;
                    }
                }
            }

            if (success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.CommitTransaction();
                    if (!CreateJpmcExcelFilenew(strAttachmentFilePath + strAttachmentFile, BulkSessionID))
                    {
                        //blnExcelFailed = true;                    
                        blnExcelFailed = false;
                    }
                }
                catch { }
                transactionStatus.InnerText = "Success";
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

                if (IsOnBehalf)
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails("", oISBLUser.ClientCode, oISBLUser.ClientName, OnBehalfUserIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();

                    //Revert to Orginal
                    oISBLUser.IsImpersonate = false;
                    oISBLUser.LoginID = oISBLUser.ImpersonateLoginID;
                    oISBLUser.ImpersonateLoginID = "";
                    oISBLUser.LoginIDEmail = OriginalLoginIDEmail;
                }
                else
                {
                    Thread threadSendMails = new Thread(delegate() { SendBulkOrderMails("", oISBLUser.ClientCode, oISBLUser.ClientName, oISBLUser.LoginIDEmail, strReportType, BulkSessionID, false); });
                    threadSendMails.IsBackground = true;
                    threadSendMails.Start();
                }

            }
            else if (!success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                transactionStatus.InnerText = "Failed";
            }

            return success;
        }
        ///*************************************************************************JPMC Chnages End*****************************************************
        //BI 20 ISIS Refresh Order - Place Refresh Order
        public bool PlaceRefreshOrderViaWebCaseNo(ISBL.User oISBLUser, String XMLData, String strAttachmentFilePath, Boolean OnBehalf, string OnBehalfUserID)
        {
            bool success = true;
            ISBL.General oGeneral = new ISBL.General();
            String OnBehalfUserIDEmail = "";
            String OriginalLoginIDEmail = "";
            Boolean IsOnBehalf = false;
            string strReportType = "";
            Boolean blnEnableRefreshOrder = false;
            string strClientCode = "";
            string strOriginalCRN = "";
            Guid BatchID = Guid.NewGuid();

            DataSet dsOrders = new DataSet();
            dsOrders.ReadXml(new StringReader(XMLData));
            DataView dvOrders = dsOrders.Tables[1].DefaultView;
            dvOrders.Sort = "CaseNo";
            DataTable dtOrders = new DataTable();
            dtOrders = dvOrders.ToTable();

            XmlDocument xmlResponse = new XmlDocument();
            XmlNode xmlResult = xmlResponse.CreateElement("result");
            xmlResponse.AppendChild(xmlResult);

            XmlNode transactionStatus = xmlResponse.CreateElement("status");
            xmlResult.AppendChild(transactionStatus);

            XmlNode xmlReportType = xmlResponse.CreateElement("ReportType");
            xmlResult.AppendChild(xmlReportType);

            DataSet dsSubjectTypeList = this.GetSubjectType();
            if (dsSubjectTypeList == null)
            {
                return false;
            }

            DataSet dsCountryList = oGeneral.GetCountryList();
            if (dsCountryList == null)
            {
                return false;
            }

            XmlNode xmlRows = xmlResponse.CreateElement("rows");
            xmlResult.AppendChild(xmlRows);

            //Get Unique Cases
            Hashtable htCaseNos = new Hashtable();
            foreach (DataRow dr in dtOrders.Rows)
            {
                if (!htCaseNos.Contains(dr["CaseNo"]))
                {
                    htCaseNos.Add(dr["CaseNo"], null);
                }
            }

            ISDL.Connect connPlaceOrder = new ISDL.Connect();
            connPlaceOrder.setConnection("ocrsConnection");

            connPlaceOrder.Open();
            connPlaceOrder.BeginTransaction();

            //Persist Bulk Order SessionID in db
            Guid BulkSessionID = Guid.NewGuid();
            if (!SaveClientBulkOrderMaster(ref connPlaceOrder, ref oISBLUser, BulkSessionID))
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                return false;
            }

            //Create Excel Attachment File
             // Try catch part added by sanjeeva reddy on 18/06/2015
            //String strAttachmentFile = Guid.NewGuid().ToString() + ".xls";
                          
            //    if (!CreateExcelFile(strAttachmentFilePath + strAttachmentFile))
            //    {
            //        blnExcelFailed = true;                    
            //    }
            
            String strAttachmentFile = BulkSessionID.ToString() + ".xls";
            
            if (OnBehalf)
            {
                ISBL.Admin oAdmin = new ISBL.Admin();
                DataSet dsUser = new DataSet();
                dsUser = oAdmin.GetUserDetails(OnBehalfUserID);
                OnBehalfUserIDEmail = dsUser.Tables[0].Rows[0]["LoginIDEmail"].ToString();

                oISBLUser.IsImpersonate = true;
                oISBLUser.ImpersonateLoginID = oISBLUser.LoginID;
                oISBLUser.LoginID = OnBehalfUserID;
                OriginalLoginIDEmail = oISBLUser.LoginIDEmail;
                oISBLUser.LoginIDEmail = OnBehalfUserIDEmail;
                IsOnBehalf = true;
            }

            //Loop through unique cases
            IDictionaryEnumerator Enum = htCaseNos.GetEnumerator();
            int caseCount = 0;
            while (Enum.MoveNext())
            {
                caseCount++;
                string CaseNo = (string)Enum.Key;
                ArrayList subjectRows = new ArrayList();
                DataRow[] subject = dtOrders.Select("CaseNo='" + CaseNo + "'");
                strReportType = "";
                blnEnableRefreshOrder = false;
                strClientCode = "";
                strOriginalCRN = "";                
                for (int i = 0; i < subject.Length; i++)
                {
                    if (i == 0)
                    {
                        strReportType = subject[i]["ReportType"].ToString();
                        blnEnableRefreshOrder = Boolean.Parse(subject[i]["EnableRefreshOrder"].ToString().ToLower());
                        strOriginalCRN = subject[i]["OriginalCRN"].ToString();
                        strClientCode = subject[i]["ClientCode"].ToString();
                    }
                                        
                    XmlNode xmlRow = xmlResponse.CreateElement("row");
                    XmlNode xmlRowNum = xmlResponse.CreateElement("rownum");
                    xmlRowNum.InnerText = subject[i]["RowNum"].ToString();
                    xmlRow.AppendChild(xmlRowNum);
                    XmlNode cols = xmlResponse.CreateElement("cols");
                    xmlRow.AppendChild(cols);

                    XmlNode col;

                    String country = "";
                    //Jul 09 Adam encode country with single quote
                    DataRow[] countries = dsCountryList.Tables[0].Select("Description='" + subject[i]["Country"].ToString().Replace("'", "''") + "'");
                    if (countries.Length == 0)
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "4";
                        cols.AppendChild(col);
                    }
                    else
                    {
                        country = countries[0]["Country"].ToString();
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    string[] arrtype = subject[i]["Type"].ToString().Split(new char[] { ';' });
                    DataRow[] drSubjectList = dsSubjectTypeList.Tables[0].Select("SubjectTypeDescription = '" + arrtype[0].ToString() + "'");
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist

                    String subjectTypeID = "";
                    if (drSubjectList.Length > 0)
                    {
                        subjectTypeID = drSubjectList[0]["SubjectType"].ToString();
                    }
                    else
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "2";
                        cols.AppendChild(col);
                    }

                    if (subject[i]["Name"].ToString() == "")
                    {
                        success = false;
                        col = xmlResponse.CreateElement("col");
                        col.InnerText = "3";
                        cols.AppendChild(col);
                    }

                    //Start 16-Feb-09 Adam: Make Client Reference No Mandatory field if Exist
                    if (System.Configuration.ConfigurationManager.AppSettings["MakeRefNoMandatoryByClient"].ToString().Contains(strClientCode))
                    {
                        if (arrtype.Length > 1)
                        {
                            if (subject[i]["CRN"].ToString().Trim().Length == 0)
                            {
                                success = false;
                                col = xmlResponse.CreateElement("col");
                                col.InnerText = "1";
                                cols.AppendChild(col);
                            }
                        }
                    }
                    //End 16-Feb-09 Adam: Make Client Reference No Mandatory for Tyco Entity

                    if (cols.ChildNodes.Count > 0) xmlRows.AppendChild(xmlRow);

                    if (success)
                    {
                        Hashtable subjectRow = new Hashtable();
                        subjectRow.Add("SubjectType", subjectTypeID);
                        subjectRow.Add("SubjectName", subject[i]["Name"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountry", country);
                        subjectRow.Add("OtherDetails", subject[i]["Details"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("ClientRefNum", subject[i]["CRN"].ToString().Replace("&", " and ").Replace("~", " ").Replace("`", " ").Replace("!", " ").Replace("$", " ").Replace("%", " ").Replace("^", " ").Replace("*", " ").Replace("=", " ").Replace("+", " ").Replace("<", " ").Replace(">", " ").Replace(",", " ").Replace("?", " ").Replace(":", " ").Replace(";", " ").Replace("|", " ").Replace("\"", " "));
                        subjectRow.Add("SubjectCountryDesc", subject[i]["Country"].ToString());
                        subjectRow.Add("OriginalSubjectID", subject[i]["OriginalSubjectID"].ToString());
                        subjectRows.Add(subjectRow);
                    }


                }

                if (success)
                {
                    if (!SaveCase(ref connPlaceOrder, oISBLUser, subjectRows, strReportType, strAttachmentFilePath, strAttachmentFile, true, BulkSessionID, blnEnableRefreshOrder, strClientCode, strOriginalCRN, BatchID))
                    {
                        try
                        {
                            connPlaceOrder.RollBackTransaction();
                        }
                        catch { }
                        connPlaceOrder.Dispose();
                        connPlaceOrder = null;

                        return false;
                    }
                }


                //end of loop report type
            }

            if (success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.CommitTransaction();
                    if (!CreateExcelFilenew(strAttachmentFilePath + strAttachmentFile, BulkSessionID))
                    {
                        //blnExcelFailed = true;                    
                        blnExcelFailed = false;
                    }

                }
                catch { }
                transactionStatus.InnerText = "Success";
                connPlaceOrder.Dispose();
                connPlaceOrder = null;

            }
            else if (!success && caseCount > 0)
            {
                try
                {
                    connPlaceOrder.RollBackTransaction();
                }
                catch { }
                connPlaceOrder.Dispose();
                connPlaceOrder = null;
                transactionStatus.InnerText = "Failed";
            }

            return success;
        }


        //Reads Excel File
        public static DataSet readExcel(String strFilePath)
        {
            ISBL.General oGeneral = new ISBL.General();
            //Error Messages
            String strMsgInvalidFile = "The excel file you uploaded is invalid. " +
                                     "Please ensure that the sheet format is intact and worksheet is named 'Sheet1'." +
                                     "You can re-download the template from the link below if required";
            String strFileName = Path.GetFileName(strFilePath);
            String strFileExtension;

            String strExcelConn, strQuery, strSheetName;
            //Connection String to Excel Workbook
            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter ExcelAdp = new OleDbDataAdapter();

            //DataTables
            DataTable dtExcelSchema = new DataTable();
            DataTable dtExcelTbl = new DataTable();
            DataSet dsExcel = new DataSet();
            strSheetName = "";

            //Get file Extension
            strFileExtension = Path.GetExtension(strFileName);
            //Convert File Extension to LowerCase
            strFileExtension = strFileExtension.ToLower();

            //Check whether the File is Excel or Not
            if (strFileExtension != ".xls")
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }

            //valid excel file
            try
            {
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                ";Extended Properties=Excel 8.0";
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }
            if (dtExcelSchema.Rows.Count < 1)
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            else
            {
                //Get sheet name
                strSheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            }
            
            //Read the Complete Block A3:E65536 in Excel File 
            try
            {
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                 ";Extended Properties=Excel 8.0";
                String sheetname = strSheetName + "A3:E65536";
                //strQuery = oGeneral.GetExcelQuery(sheetname); //prevent sql injection
                strQuery = "Select * From [Sheet1$A3:E65536]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);
                dtExcelSchema = new DataTable();
                ExcelAdp.Fill(dtExcelSchema);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
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
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }


            //Read the Single Cell in Excel File 

            try
            {
                DataTable dtReportType = new DataTable();
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                "Data Source=" + strFilePath + ";" +
                                "Extended Properties='Excel 8.0;HDR=No'";
                String sheetname = strSheetName + "C1:C1";
                strQuery = "Select * From [Sheet1$C1:C1]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtReportType);
                dsExcel.Tables.Add(dtReportType);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            //Create a new datatable
            DataColumn dcDataColumn;
            dcDataColumn = new DataColumn("Checked", System.Type.GetType("System.Boolean"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Client Reference Number", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Type", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Name", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Country", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Other Details", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);

            int intRowNumber = 0;
            int i = 0;
            try
            {
                while (i < dtExcelSchema.Rows.Count)
                {
                    if (dtExcelSchema.Rows[i].IsNull("Client Reference Number") && dtExcelSchema.Rows[i].IsNull("Subject Type")
                        && dtExcelSchema.Rows[i].IsNull("Subject Name") && dtExcelSchema.Rows[i].IsNull("Subject Country")
                         && dtExcelSchema.Rows[i].IsNull("Other Details"))
                    {
                        dtExcelSchema.Rows[i].Delete();
                    }
                    else
                    {
                        dtExcelTbl.Rows.Add();
                        dtExcelTbl.Rows[intRowNumber]["Checked"] = "True";
                        dtExcelTbl.Rows[intRowNumber]["Client Reference Number"] = dtExcelSchema.Rows[i]["Client Reference Number"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Type"] = dtExcelSchema.Rows[i]["Subject Type"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Name"] = dtExcelSchema.Rows[i]["Subject Name"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Country"] = dtExcelSchema.Rows[i]["Subject Country"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Other Details"] = dtExcelSchema.Rows[i]["Other Details"].ToString().Trim();
                        intRowNumber++;
                    }
                    i++;
                }
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }
            //Add table to the dataset
            dsExcel.Tables.Add(dtExcelTbl);
            return dsExcel;
        }

        //Reads Excel File
        public static DataSet readExcelNew(String strFilePath)
        {
            ISBL.General oGeneral = new ISBL.General();
            //Error Messages
            String strMsgInvalidFile = "The excel file you uploaded is invalid. " +
                                     "Please ensure that the sheet format is intact. " +
                                     "You can re-download the template from the link below if required";
            String strFileName = Path.GetFileName(strFilePath);
            String strFileExtension;

            String strExcelConn, strQuery, strSheetName;
            //Connection String to Excel Workbook
            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter ExcelAdp = new OleDbDataAdapter();

            //DataTables
            DataTable dtExcelSchema = new DataTable();
            DataTable dtExcelTbl = new DataTable();
            DataSet dsExcel = new DataSet();
            strSheetName = "";

            //Get file Extension
            strFileExtension = Path.GetExtension(strFileName);
            //Convert File Extension to LowerCase
            strFileExtension = strFileExtension.ToLower();

            //Check whether the File is Excel or Not
            if (strFileExtension == ".xlsx")
            {
                strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                   ";Extended Properties=Excel 12.0";
            }
            else if (strFileExtension == ".xls")
            {
                      strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                    ";Extended Properties=Excel 8.0";
            }
            else
            {
                 throw new System.InvalidOperationException(strMsgInvalidFile);
            }

            //valid excel file
            try
            {

                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            if (dtExcelSchema.Rows.Count < 1)
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            else
            {
                //Get sheet name
                strSheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            }

            //Read the Complete Block A3:E65536 in Excel File 
            try
            {
                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties=Excel 12.0;";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
                }
                String sheetname = strSheetName + "A3:F65536";
                //strQuery = oGeneral.GetExcelQuery(sheetname);//prevent sql injection
                strQuery = "Select * From [Sheet1$A3:F65536]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);


                dtExcelSchema = new DataTable();
                ExcelAdp.Fill(dtExcelSchema);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
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
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }


            //Read the Single Cell in Excel File 

            try
            {
                DataTable dtReportType = new DataTable();

                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties='Excel 12.0;HDR=No'";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                  ";Extended Properties='Excel 8.0;HDR=No'";
                }

                String sheetname = strSheetName + "C1:C1";
                strQuery = "Select * From [Sheet1$C1:C1]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtReportType);
                dsExcel.Tables.Add(dtReportType);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            //Create a new datatable
            DataColumn dcDataColumn;
            dcDataColumn = new DataColumn("Checked", System.Type.GetType("System.Boolean"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Client Reference Number", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Type", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Name", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Country", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Other Details", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Group Your Order", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);


            int intRowNumber = 0;
            int i = 0;
            try
            {
                while (i < dtExcelSchema.Rows.Count)
                {
                    if (dtExcelSchema.Rows[i].IsNull("Client Reference Number") && dtExcelSchema.Rows[i].IsNull("Subject Type")
                        && dtExcelSchema.Rows[i].IsNull("Subject Name") && dtExcelSchema.Rows[i].IsNull("Subject Country")
                         && dtExcelSchema.Rows[i].IsNull("Other Details") && dtExcelSchema.Rows[i].IsNull("Group Your Order"))
                    {
                        dtExcelSchema.Rows[i].Delete();
                    }
                    else
                    {
                        dtExcelTbl.Rows.Add();
                        dtExcelTbl.Rows[intRowNumber]["Checked"] = "True";
                        dtExcelTbl.Rows[intRowNumber]["Client Reference Number"] = dtExcelSchema.Rows[i]["Client Reference Number"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Type"] = dtExcelSchema.Rows[i]["Subject Type"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Name"] = dtExcelSchema.Rows[i]["Subject Name"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Country"] = dtExcelSchema.Rows[i]["Subject Country"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Other Details"] = dtExcelSchema.Rows[i]["Other Details"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Group Your Order"] = dtExcelSchema.Rows[i]["Group Your Order"].ToString().Trim();
                        intRowNumber++;
                    }
                    i++;
                }
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }
            //Add table to the dataset
            dsExcel.Tables.Add(dtExcelTbl);



            //Read the Single Cell in Excel File for Version

            try
            {
                DataTable dtVersion = new DataTable();

                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties='Excel 12.0;HDR=No'";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                  ";Extended Properties='Excel 8.0;HDR=No'";
                }

                strQuery = "Select * FROM [Sheet2$E2:E2]";
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtVersion);
                dsExcel.Tables.Add(dtVersion);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }
            // added new code for binding Sub report type from excel selected
            try
            {
                DataTable dtSubreport = new DataTable();

                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties='Excel 12.0;HDR=No'";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                  ";Extended Properties='Excel 8.0;HDR=No'";
                }

                strQuery = "Select * FROM [Sheet1$E1:E1]";
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtSubreport);
                dsExcel.Tables.Add(dtSubreport);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            return dsExcel;
        }
        //New Addition Start
        //Reads JPMC Excel File
        public static DataSet readJpmcExcelNew(String strFilePath)
        {
            ISBL.General oGeneral = new ISBL.General();
            //Error Messages
            String strMsgInvalidFile = "The excel file you uploaded is invalid. " +
                                     "Please ensure that the sheet format is intact. " +
                                     "You can re-download the template from the link below if required";
            String strFileName = Path.GetFileName(strFilePath);
            String strFileExtension;

            String strExcelConn, strQuery, strSheetName;
            //Connection String to Excel Workbook
            OleDbConnection connExcel = new OleDbConnection();
            OleDbCommand cmdExcel = new OleDbCommand();
            OleDbDataAdapter ExcelAdp = new OleDbDataAdapter();

            //DataTables
            DataTable dtExcelSchema = new DataTable();
            DataTable dtExcelTbl = new DataTable();
            DataSet dsExcel = new DataSet();
            strSheetName = "";

            //Get file Extension
            strFileExtension = Path.GetExtension(strFileName);
            //Convert File Extension to LowerCase
            strFileExtension = strFileExtension.ToLower();

            //Check whether the File is Excel or Not
            if (strFileExtension == ".xlsx")
            {
                strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                   ";Extended Properties=Excel 12.0";
            }
            else if (strFileExtension == ".xls")
            {
                strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                              ";Extended Properties=Excel 8.0";
            }
            else
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }

            //valid excel file
            try
            {

                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                dtExcelSchema = connExcel.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            if (dtExcelSchema.Rows.Count < 1)
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            else
            {
                //Get sheet name
                strSheetName = dtExcelSchema.Rows[0]["TABLE_NAME"].ToString();
            }

            //Read the Complete Block A3:E65536 in Excel File 
            try
            {
                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties=Excel 12.0";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=1\"";
                }
               

                // changes the range from K to L
                String sheetname = strSheetName + "A3:L65536";
                //strQuery = oGeneral.GetExcelQuery(sheetname);//prevent sql injection
                strQuery = "Select * From [Sheet1$A3:L65536]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);
                dtExcelSchema = new DataTable();
                

                ExcelAdp.Fill(dtExcelSchema);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
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
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }


            //Read the Single Cell in Excel File 

            try
            {
                DataTable dtReportType = new DataTable();

                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties='Excel 12.0;HDR=NO'";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                  ";Extended Properties='Excel 8.0;HDR=NO'";
                }

                String sheetname = strSheetName + "C1:C1";
                strQuery = "Select * From [Sheet1$C1:C1]"; //prevent sql injection
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtReportType);
                dsExcel.Tables.Add(dtReportType);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }

            //Create a new datatable
            DataColumn dcDataColumn;
            dcDataColumn = new DataColumn("Checked", System.Type.GetType("System.Boolean"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Client Reference Number", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Type", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Name", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subject Country", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);

            dcDataColumn = new DataColumn("Subject Aliases", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Subreport Type", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("DOB", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("ID", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("ADDRESS", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            dcDataColumn = new DataColumn("Additional Information", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            
            dcDataColumn = new DataColumn("Group Your Order", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);
            //added new column as Initial Renewal - code added by Deepak

            dcDataColumn = new DataColumn("InitialRenewal", System.Type.GetType("System.String"));
            dtExcelTbl.Columns.Add(dcDataColumn);

            int intRowNumber = 0;
            int i = 0;
            try
            {
                while (i < dtExcelSchema.Rows.Count)
                {
                    if (dtExcelSchema.Rows[i].IsNull("Client Reference Number") && dtExcelSchema.Rows[i].IsNull("Subject Type")
                        && dtExcelSchema.Rows[i].IsNull("Subject Name") && dtExcelSchema.Rows[i].IsNull("Subject Country")
                         && dtExcelSchema.Rows[i].IsNull("Subject Aliases") && dtExcelSchema.Rows[i].IsNull("Subreport Type") && dtExcelSchema.Rows[i].IsNull("DOB") && dtExcelSchema.Rows[i].IsNull("ID") && dtExcelSchema.Rows[i].IsNull("ADDRESS") && dtExcelSchema.Rows[i].IsNull("Additional Information") && dtExcelSchema.Rows[i].IsNull("Group Your Order"))
                    {
                        dtExcelSchema.Rows[i].Delete();
                    }
                    else
                    {
                        dtExcelTbl.Rows.Add();
                        dtExcelTbl.Rows[intRowNumber]["Checked"] = "True";
                        dtExcelTbl.Rows[intRowNumber]["Client Reference Number"] = dtExcelSchema.Rows[i]["Client Reference Number"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Type"] = dtExcelSchema.Rows[i]["Subject Type"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Name"] = dtExcelSchema.Rows[i]["Subject Name"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Country"] = dtExcelSchema.Rows[i]["Subject Country"].ToString().Trim();

                        //dtExcelTbl.Rows[intRowNumber]["Other Details"] = dtExcelSchema.Rows[i]["Other Details"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subject Aliases"] = dtExcelSchema.Rows[i]["Subject Aliases"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Subreport Type"] = dtExcelSchema.Rows[i]["Subreport Type"].ToString().Trim();                        
                        dtExcelTbl.Rows[intRowNumber]["DOB"] = dtExcelSchema.Rows[i]["DOB"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["ID"] = dtExcelSchema.Rows[i]["ID"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["ADDRESS"] = dtExcelSchema.Rows[i]["ADDRESS"].ToString().Trim();
                        dtExcelTbl.Rows[intRowNumber]["Additional Information"] = dtExcelSchema.Rows[i]["Additional Information"].ToString().Trim();

                        dtExcelTbl.Rows[intRowNumber]["Group Your Order"] = dtExcelSchema.Rows[i]["Group Your Order"].ToString().Trim();
                        // added new column as Initial Renewal - code added by Deepak
                        // Combining other details in Additional field
                        dtExcelTbl.Rows[intRowNumber]["InitialRenewal"] = dtExcelSchema.Rows[i]["Initial/Renewal"].ToString().Trim();

                        dtExcelTbl.Rows[intRowNumber]["Additional Information"]="Subject Alias -" +
                            dtExcelSchema.Rows[i]["Subject Aliases"].ToString().Trim() + Environment.NewLine +"DOB -" + dtExcelSchema.Rows[i]["DOB"].ToString().Trim() + Environment.NewLine +
                            "ID -" + dtExcelSchema.Rows[i]["ID"].ToString().Trim() + Environment.NewLine + "Address -" + dtExcelSchema.Rows[i]["ADDRESS"].ToString().Trim() + Environment.NewLine +
                            "Additional Information -" + dtExcelSchema.Rows[i]["Additional Information"].ToString().Trim() + Environment.NewLine + "Initial/Renewal -" + dtExcelSchema.Rows[i]["Initial/Renewal"].ToString().Trim();

                        intRowNumber++;
                    }
                    i++;
                }
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }
            //Add table to the dataset
            dsExcel.Tables.Add(dtExcelTbl);



            //Read the Single Cell in Excel File for Version

            try
            {
                DataTable dtVersion = new DataTable();

                if (strFileExtension == ".xlsx")
                {
                    strExcelConn = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + strFilePath +
                                       ";Extended Properties='Excel 12.0;HDR=No'";
                }
                else if (strFileExtension == ".xls")
                {
                    strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + strFilePath +
                                  ";Extended Properties='Excel 8.0;HDR=No'";
                }

                strQuery = "Select * FROM [Sheet2$E2:E2]";
                connExcel = new OleDbConnection(strExcelConn);
                connExcel.Open();
                cmdExcel = new OleDbCommand(strQuery, connExcel);
                ExcelAdp = new OleDbDataAdapter(cmdExcel);

                ExcelAdp.Fill(dtVersion);
                dsExcel.Tables.Add(dtVersion);
            }
            catch
            {
                throw new System.InvalidOperationException(strMsgInvalidFile);
            }
            finally
            {
                connExcel.Close();
                connExcel.Dispose();
                cmdExcel.Dispose();
                ExcelAdp.Dispose();
            }



            return dsExcel;
        }
        //Create New Excel File
        private Boolean CreateExcelFile(String strFileName)
        {
            try
            {
                //Get the Attachment Template File Path from Web.Config
                String strOutputTemplateFileName = System.Configuration.ConfigurationManager.AppSettings["OutputTemplateFilePath"];
                //Copy File to Temp folder
                File.Copy(strOutputTemplateFileName, strFileName);
                return true;
            }
            catch
            {
                return false;
            }
        }
        private Boolean CreateExcelFilenew(String strFileName, Guid bulkid)
        {
            String strMsgInvalidFile = "Excel File Creation Failed. Try Again..."; 
            try
            {
                FileInfo newfile = new FileInfo(strFileName.ToString());
                using (ExcelPackage xlpackage = new ExcelPackage(newfile))
                {
                    ExcelWorksheet ws = xlpackage.Workbook.Worksheets.Add("BulkOrderData");
                    ws.View.ShowGridLines = true;
                    DataSet dsbulk3 = new DataSet();
                    dsbulk3 = GetBulkExcelOrderDetails(bulkid);
                    DataTable dt1 = dsbulk3.Tables[0];
                    ws.Cells["A1"].LoadFromDataTable(dt1, true);
                    ws.Column(1).AutoFit();
                    ws.Column(2).AutoFit();
                    ws.Column(3).AutoFit();
                    ws.Column(4).AutoFit();
                    ws.Column(5).AutoFit();
                    ws.Column(6).AutoFit();
                    ws.Column(7).AutoFit();
                    ws.Column(8).AutoFit();
                    ws.Column(9).AutoFit();
                    ws.Row(1).Style.Font.Bold = true;
                    xlpackage.SaveAs(newfile);
                    xlpackage.Dispose();

                }

             }
             catch (Exception ex)
             {
                 Guid EventLogGUID4 = new Guid();
                 EventLogGUID4 = Guid.NewGuid();
                 oISBLGen.SaveEventLog(EventLogGUID4.ToString(), "T056", "", "Excel File Creation", "", "", DateTime.Now, DateTime.Now, "", "", "", "Exception in Excel Objectcreation " + ex.ToString());
                 return true;
                 throw new System.InvalidOperationException(strMsgInvalidFile);

             }
             finally
             {
                 //connExcel.Close();
                 //connExcel.Dispose();
                 //cmdExcel.Dispose();
                 //ExcelAdp.Dispose();
             }
            return true;
        }
        private Boolean CreateJpmcExcelFilenew(String strFileName, Guid bulkid)
        {
            String strMsgInvalidFile = "Excel File Creation Failed. Try Again...";
            try
            {
                FileInfo newfile = new FileInfo(strFileName.ToString());
                using (ExcelPackage xlpackage = new ExcelPackage(newfile))
                {
                    ExcelWorksheet ws = xlpackage.Workbook.Worksheets.Add("BulkOrderData");
                    ws.View.ShowGridLines = true;
                    DataSet dsbulk3 = new DataSet();
                    dsbulk3 = GetJpmcBulkExcelOrderDetails(bulkid);
                    DataTable dt1 = dsbulk3.Tables[0];
                    ws.Cells["A1"].LoadFromDataTable(dt1, true);
                    ws.Column(1).AutoFit();
                    ws.Column(2).AutoFit();
                    ws.Column(3).AutoFit();
                    ws.Column(4).AutoFit();
                    ws.Column(5).AutoFit();
                    ws.Column(6).AutoFit();
                    ws.Column(7).AutoFit();
                    ws.Column(8).AutoFit();
                    ws.Column(9).AutoFit();
                    //ws.Column(10).AutoFit();
                    ws.Row(1).Style.Font.Bold = true;
                    xlpackage.SaveAs(newfile);
                    xlpackage.Dispose();

                }

            }
            catch (Exception ex)
            {
                Guid EventLogGUID4 = new Guid();
                EventLogGUID4 = Guid.NewGuid();
                oISBLGen.SaveEventLog(EventLogGUID4.ToString(), "T056", "", "Excel File Creation", "", "", DateTime.Now, DateTime.Now, "", "", "", "Exception in Jpmc Excel Objectcreation " + ex.ToString());
                return true;
                throw new System.InvalidOperationException(strMsgInvalidFile);

            }
            finally
            {
                //connExcel.Close();
                //connExcel.Dispose();
                //cmdExcel.Dispose();
                //ExcelAdp.Dispose();
            }
            return true;
        }
        private bool GrantAccess(string fullPath)
        {
            String dirName = Path.GetDirectoryName(fullPath);
            DirectoryInfo dInfo = new DirectoryInfo(dirName);
            DirectorySecurity dSecurity = dInfo.GetAccessControl();
            dSecurity.AddAccessRule(new FileSystemAccessRule(new SecurityIdentifier(WellKnownSidType.WorldSid, null), FileSystemRights.FullControl, InheritanceFlags.ObjectInherit | InheritanceFlags.ContainerInherit, PropagationFlags.NoPropagateInherit, AccessControlType.Allow));
            dInfo.SetAccessControl(dSecurity);
            return true;
        }
        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                
            }
            finally
            {
                GC.Collect();
            }
        }

        //Add New Row To Excel File
        private Boolean AddExcelRow(String strFileName, String strOrderID, String strClientRefNum, String strReportType, String strSubjectType, String strSubjectName, String strCountry, String strOtherDetails, int intPrimary)
        {
            if (!File.Exists(strFileName)) return false;

            String strExcelConn = "Provider=Microsoft.Jet.OLEDB.4.0;" +
                                   "Data Source=" + strFileName + ";" +
                                   "Extended Properties='Excel 8.0;HDR=Yes'";

            String strPrimary;
            OleDbConnection connExcel = new OleDbConnection(strExcelConn);
            OleDbCommand cmdExcel = new OleDbCommand();

            try
            {
                connExcel.Open();
                cmdExcel.Connection = connExcel;

                if (intPrimary == 1)
                    strPrimary = "Yes";
                else
                    strPrimary = "No";

                //Add New Row to Excel File 
                cmdExcel.CommandText = "INSERT INTO [OrderDetails$] (OrderID, ClientReferenceNumber, ReportType, " +
                                         "SubjectType, SubjectName, SubjectCountry, PrimarySubject, OtherDetails) " +
                                         "values ('" + strOrderID + "', '" + strClientRefNum.Replace("'","''") + "', '" + strReportType + "', '" + strSubjectType +
                                         "', '" + strSubjectName.Replace("'", "''") + "', '" + strCountry.Replace("'", "''") + "', '" + strPrimary + "', '" + strOtherDetails.Replace("'", "''") + "')";
                cmdExcel.ExecuteNonQuery();
                return true;
            }
            catch
            {
                return false;                
            }
            finally
            {
                connExcel.Close();
                cmdExcel.Dispose();
                connExcel.Dispose();
            }
        }
        //New Addition End

        //Saves Case with multiple subjects.  
        private Boolean SaveCase(ref ISDL.Connect connPlaceOrder, ISBL.User oISBLUser, ArrayList lstSubjects, String strReportType, String strAttachmentFilePath, String strAttachmentFile, Boolean isExcel, Guid BulkSessionID)
        {
            ISBL.General oISBLGen = new ISBL.General();
            Hashtable hmSubject;
            int intCompanyCount, intIndvCount;
            Guid OrderID, OrderSubjectID;
            DateTime dtOrderDateTime, dtReportDueDate;
            Boolean blnEnableAutoAssignment, blnEnableSingleCRNPerExcel;
            String strOfficeAssignment, strDueDate, strflsDueDate, strResearchElementCode;
            String strSubjectAliases,strDOB,strID,strADDRESS,strAdditionalInformation;
            int intDueDay;
            int intRECount;
            float flBudget = 0;
            float flSBudget = 0;
            String strSubjectName, strSubjectType, strCountry, strCountryDesc, strClientRefNum, strOtherDetails, strSubjectTypeDesc,strSubReportType;
            Boolean blnStatus = true;
            Boolean isCalculationRevert = false; //BI 34
            Boolean isflsCalculationRevert = false; 
            
            try
            {
                blnEnableSingleCRNPerExcel = ISEnableSingleCRNPerExcel(ref oISBLUser);
            }
            catch (Exception ex)
            {
                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_1" : "BO_10", "Error Getting Single CRN Per Excel: " + ex.Message);
                oISBLGen.DisposeConnection();
                oISBLGen = null;

                //New Addition Start
                //14/10/2015 on by sanjeeva start
                if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                    File.Delete(strAttachmentFilePath + strAttachmentFile);
                // end of sanjeeva
                //New Addition End

                return false;
            }

            strClientRefNum = "";

            try
            {
                blnEnableAutoAssignment = ISEnableAutoAssignment(oISBLUser.ClientCode);
            }
            catch (Exception ex)
            {
                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_2" : "BO_11", "Error Getting Enable Auto Assignment: " + ex.Message);
                oISBLGen.DisposeConnection();
                oISBLGen = null;

                //New Addition Start
                if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                    File.Delete(strAttachmentFilePath + strAttachmentFile);
                //New Addition End

                return false;
            }

            if (blnEnableSingleCRNPerExcel || isExcel)
            {
                OrderID = Guid.NewGuid();
                dtOrderDateTime = DateTime.Now;

                hmSubject = (Hashtable)lstSubjects[0];
                strClientRefNum = hmSubject["ClientRefNum"].ToString();

                intCompanyCount = 0;
                intIndvCount = 0;

                 DataSet dsSubReport1 = null;
                    Boolean bolReportFlag1 = false;
                    dsSubReport1 = oISBLGen.GetReportFlag(oISBLUser.ClientCode);
                    if (dsSubReport1.Tables[0].Rows.Count > 0)
                    {
                        if (dsSubReport1.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString() == "")
                        {
                            bolReportFlag1 = false;
                        }
                        bolReportFlag1 = Convert.ToBoolean(dsSubReport1.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString());
                    }
                    else
                    {
                        bolReportFlag1 = false;
                    }
                   
                /////////////////////////////////////Client Order Insertion/////////////////////////////////////////////////////////////////////////////////////////////
                //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam 
                if (oISBLUser.IsImpersonate)
                {
                    if (bolReportFlag1 == false)
                    {
                        blnStatus = SaveClientOrderNonJPMC(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", oISBLUser.ImpersonateLoginID,hmSubject["SubReportType"].ToString());
                    }
                    else
                    blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", oISBLUser.ImpersonateLoginID);

                }
                else
                {
                    if (bolReportFlag1 == false)
                    {
                        blnStatus = SaveClientOrderNonJPMC(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", hmSubject["SubReportType"].ToString());
                    }
                    else
                    blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE");

                }
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    
                       if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                        File.Delete(strAttachmentFilePath + strAttachmentFile);
                    
                    return false;
                }
                try
                {
                    strDueDate = "";
                    flBudget = float.Parse(oISBLGen.CalculateBudget(oISBLUser.ClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                    dtReportDueDate = DateTime.Parse(strDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);

                }
                catch (Exception ex)
                {
                    //Start BI 34 - Budget recalculation
                    //string xmlData = "<input><reportType>" + strReportType + "</reportType></input>";
                    string xmlData = "<input><clientCode>" + oISBLUser.ClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                    //End BI 34 - Budget recalculation
                    oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Report Type Due Days: " + ex.Message);
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    return false;
                }
                blnStatus = UpdateBudgetDuedate(ref connPlaceOrder, OrderID, ref oISBLUser, dtReportDueDate, flBudget);
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    return false;
                }
               
                /////////////////////////////////////Client Order Email/////////////////////////////////////////////////////////////////////////////////////////////
                blnStatus = SaveClientOrderEmail(ref connPlaceOrder, OrderID, ref oISBLUser);
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;


                    //New Addition Start
                        //14/10/2015 on by sanjeeva start
                        if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                            File.Delete(strAttachmentFilePath + strAttachmentFile);
                        //end of sanjeeva
                    //New Addition End

                    return false;
                }
                    /////////////////////////////////////Client Order Subjects Insertion/////////////////////////////////////////////////////////////////////////////////////////////
                    DataTable rankDt = (oISBLGen.ConvertArrayListToDatatable(lstSubjects));

                    rankDt.Columns.Add("Rank");
                  int rank = 1;

                  for (int i = 0; i <= rankDt.Rows.Count - 1; i++)
                  {
                      rankDt.Rows[i]["Rank"] = rank;

                      if (rankDt.Rows.Count - 1 == i)
                      {
                          break; // TODO: might not be correct. Was : Exit For
                      }

                      DataRow thisRow = rankDt.Rows[i];
                      DataRow nextRow = rankDt.Rows[i + 1];

                      if (thisRow["SubjectCountryDesc"].ToString() != nextRow["SubjectCountryDesc"].ToString())
                      {
                          rank = 1;
                      }
                      else
                      {
                          rank += 1;
                      }
                  }
                               
                DataTable lstSubjectsDt = rankDt;
                #region commented old code
                // for (int x = 0; x < lstSubjects.Count; x++)
                //{
                //    OrderSubjectID = Guid.NewGuid();
                    
                //    int intcounter = 1;
                //    hmSubject = (Hashtable)lstSubjects[x];
                //    strSubjectName = hmSubject["SubjectName"].ToString();
                //    strSubjectType = hmSubject["SubjectType"].ToString();
                //    //Dim dsSubReport As DataSet
                //    DataSet dsSubReport = null;
                //    Boolean bolReportFlag = false;
                //    dsSubReport = oISBLGen.GetReportFlag(oISBLUser.ClientCode);
                //    if (dsSubReport.Tables[0].Rows.Count > 0)
                //    {
                //        if (dsSubReport.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString() == "")
                //        {
                //            bolReportFlag = false;
                //        }
                //        bolReportFlag = Convert.ToBoolean(dsSubReport.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString());
                //    }
                //    else
                //    {
                //        bolReportFlag = false;
                //    }
                //    if (bolReportFlag == true)
                //    {
                //        strSubReportType = hmSubject["SubreportType"].ToString();
                //        strSubjectAliases = hmSubject["SubjectAliases"].ToString();
                //        strDOB = hmSubject["DOB"].ToString();
                //        strID = hmSubject["ID"].ToString();
                //        strADDRESS = hmSubject["ADDRESS"].ToString();
                //        strAdditionalInformation = hmSubject["AdditionalInformation"].ToString();
                //        strOtherDetails = strAdditionalInformation;
                //    }
                //    else
                //    {
                //        strSubReportType = hmSubject["SubReportType"].ToString();
                //        strSubjectAliases = "";
                //        strDOB = "";
                //        strID = "";
                //        strADDRESS = "";
                //        strAdditionalInformation = "";
                //        strOtherDetails = hmSubject["OtherDetails"].ToString();
                //    }
                //    //strSubReportType = hmSubject["SubreportType"].ToString();
                //    //strSubjectAliases = hmSubject["SubjectAliases"].ToString();
                //    //strDOB = hmSubject["DOB"].ToString();
                //    //strID = hmSubject["ID"].ToString();
                //    //strADDRESS = hmSubject["ADDRESS"].ToString();
                //    //strAdditionalInformation = hmSubject["AdditionalInformation"].ToString();

                //    strSubjectTypeDesc = "";
                //    if (strSubjectType == "1")
                //    {
                //        intIndvCount++;
                //        strSubjectTypeDesc = "Individual";
                //    }
                //    if (strSubjectType == "2")
                //    {
                //        intCompanyCount++;
                //        strSubjectTypeDesc = "Company";
                //    }

                //    strCountry = hmSubject["SubjectCountry"].ToString();

                //    //New Addition Start
                //    strCountryDesc = hmSubject["SubjectCountryDesc"].ToString();
                //    //New Addition End

                   

                //    string mSubreporttype = strSubReportType;
                //    if (bolReportFlag == false)
                //        mSubreporttype = "";

                //        if (x == 0)
                //        {

                //            blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 1, mSubreporttype, strSubjectAliases, strDOB, strID, strADDRESS, strAdditionalInformation);
                            
                //        }
                //        else
                //        {
                //            blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 0, mSubreporttype, strSubjectAliases, strDOB, strID, strADDRESS, strAdditionalInformation);      
                     
                //        }
                //        if (!blnStatus)
                //        {
                //            oISBLGen.DisposeConnection();
                //            oISBLGen = null;

                //            return false;
                //        }
                //        ////////////////////////////////////////////////////////////Budget Calculation///////////////////////////////

                //        if (bolReportFlag == true)
                //        {
                //            try
                //            {
                //                strDueDate = "";
                //                strflsDueDate = "";
                //                // flBudget = float.Parse(oISBLGen.CalculateBudget(oISBLUser.ClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                //                flSBudget = float.Parse(oISBLGen.CalculateFlsBudget(oISBLUser.ClientCode, strReportType, strSubReportType, strCountry, strSubjectType, false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isflsCalculationRevert, ref strflsDueDate,intcounter));
                //                // SLBudget = oISBL.CalculateFlsBudget(strClientCode, strReportType, dr.Item("SubReportTypeCode"), StrCountry, subtype, chkExpressCase.Checked, False, Session("SubjectDataView"), blnIsCalculationRevert, strDueDate)
                //                //oISBLUser.ClientCode,strReportType,strSubReportType,strCountry,strSubjectType
                //                dtReportDueDate = DateTime.Parse(strflsDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);
                //                /////////////////////////////////////Update Budget/////////////////////////////////////////////////////////////////////////////////////////////
                //                blnStatus = UpdateflsBudgetDuedate(ref connPlaceOrder, OrderID, ref oISBLUser, dtReportDueDate, flSBudget, strSubReportType, OrderSubjectID, strSubjectName, strSubjectType);
                //                if (!blnStatus)
                //                {
                //                    oISBLGen.DisposeConnection();
                //                    oISBLGen = null;
                //                    return false;
                //                }
                //            }
                //            catch (Exception ex)
                //            {
                //                //Start BI 34  Budget recalculation
                //                string xmlData = "<input><clientCode>" + oISBLUser.ClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                //                //End BI 34  Budget recalculation
                //                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Budget and Due Date UpdateFlsBudget  : " + ex.Message);
                //                oISBLGen.DisposeConnection();
                //                oISBLGen = null;

                //                //New Addition Start
                //                //14/10/2015 on by sanjeeva start
                //                if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                //                    File.Delete(strAttachmentFilePath + strAttachmentFile);
                //                //end of sanjeeva
                //                //New Addition End

                //                return false;
                //            }
                //        }
                //        DataSet dsResearchElements = null; 
                       
                //        try
                //        {
                //            if (bolReportFlag == true)
                //            {

                //                dsResearchElements = oISBLGen.GetResearchElementBySubReportType(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc, strSubReportType);
                //            }
                //            else
                //            {
                //                if (strSubReportType != "")
                //                {
                //                    dsResearchElements = oISBLGen.GetResearchElementBySubReportType(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc, strSubReportType);
                //                }
                //                else
                //                {

                //                    dsResearchElements = oISBLGen.GetResearchElement(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc);
                //                }
                //            }
                //            intRECount = dsResearchElements.Tables[0].Rows.Count;
                //        }
                //        catch (Exception ex)
                //        {
                //            oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_6" : "BO_15", "Error Getting Research Elements: " + ex.Message);
                //            oISBLGen.DisposeConnection();
                //            oISBLGen = null;

                        
                //        return false;
                //       }

                //    for (int y = 0; y < intRECount; y++)
                //    {
                //        //Start Adam Jul 2009 IS FCPA
                //        if (strReportType == "IS FCPA")
                //        {
                //            if (x == 0) //Primary
                //                strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                //            else
                //            {
                //                if (strSubjectType == "1") //Individual
                //                    strResearchElementCode = "30"; //Individual
                //                else
                //                    strResearchElementCode = "14"; //Company

                //                blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                //                if (!blnStatus)
                //                {
                //                    oISBLGen.DisposeConnection();
                //                    oISBLGen = null;
                //                    return false;
                //                 }
                //                    break;
                //                }
                //            }
                //            else
                //            {
                //                strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                //            }
                //            //strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString(); //Adam Jul 2009 IS FCPA
                //            //End Adam Jul 2009 IS FCPA
                //            /////////////////////////////////////Client Order Subject RE       Insertion/////////////////////////////////////////////////////////////////////////////////////////////
                //                blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                            
                //            if (!blnStatus)
                //            {
                //                oISBLGen.DisposeConnection();
                //                oISBLGen = null;
                                                                
                //                return false;
                //            }
                //     }
                // }
                #endregion

                for (int x = 0; x < lstSubjectsDt.Rows.Count; x++)
                {
                    OrderSubjectID = Guid.NewGuid();

                    int intcounter = Convert.ToInt32(lstSubjectsDt.Rows[x]["Rank"]);
                    
                    strSubjectName = lstSubjectsDt.Rows[x]["SubjectName"].ToString();
                    strSubjectType = lstSubjectsDt.Rows[x]["SubjectType"].ToString();
                    //Dim dsSubReport As DataSet
                    DataSet dsSubReport = null;
                    Boolean bolReportFlag = false;
                    dsSubReport = oISBLGen.GetReportFlag(oISBLUser.ClientCode);
                    if (dsSubReport.Tables[0].Rows.Count > 0)
                    {
                        if (dsSubReport.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString() == "")
                        {
                            bolReportFlag = false;
                        }
                        bolReportFlag = Convert.ToBoolean(dsSubReport.Tables[0].Rows[0]["EnableRepSubRepTypeflag"].ToString());
                    }
                    else
                    {
                        bolReportFlag = false;
                    }
                    if (bolReportFlag == true)
                    {
                        strSubReportType = lstSubjectsDt.Rows[x]["SubreportType"].ToString();
                        strSubjectAliases = lstSubjectsDt.Rows[x]["SubjectAliases"].ToString();
                        strDOB = lstSubjectsDt.Rows[x]["DOB"].ToString();
                        strID = lstSubjectsDt.Rows[x]["ID"].ToString();
                        strADDRESS = lstSubjectsDt.Rows[x]["ADDRESS"].ToString();
                        strAdditionalInformation = lstSubjectsDt.Rows[x]["AdditionalInformation"].ToString();
                        strOtherDetails = strAdditionalInformation;
                    }
                    else
                    {
                        strSubReportType = lstSubjectsDt.Rows[x]["SubReportType"].ToString();
                        strSubjectAliases = "";
                        strDOB = "";
                        strID = "";
                        strADDRESS = "";
                        strAdditionalInformation = "";
                        strOtherDetails = lstSubjectsDt.Rows[x]["OtherDetails"].ToString();
                    }
                    //strSubReportType = hmSubject["SubreportType"].ToString();
                    //strSubjectAliases = hmSubject["SubjectAliases"].ToString();
                    //strDOB = hmSubject["DOB"].ToString();
                    //strID = hmSubject["ID"].ToString();
                    //strADDRESS = hmSubject["ADDRESS"].ToString();
                    //strAdditionalInformation = hmSubject["AdditionalInformation"].ToString();

                    strSubjectTypeDesc = "";
                    if (strSubjectType == "1")
                    {
                        intIndvCount++;
                        strSubjectTypeDesc = "Individual";
                    }
                    if (strSubjectType == "2")
                    {
                        intCompanyCount++;
                        strSubjectTypeDesc = "Company";
                    }

                    strCountry = lstSubjectsDt.Rows[x]["SubjectCountry"].ToString();

                    //New Addition Start
                    strCountryDesc = lstSubjectsDt.Rows[x]["SubjectCountryDesc"].ToString();
                    //New Addition End



                    string mSubreporttype = strSubReportType;
                    if (bolReportFlag == false)
                        mSubreporttype = "";

                    if (x == 0)
                    {

                        blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 1, mSubreporttype, strSubjectAliases, strDOB, strID, strADDRESS, strAdditionalInformation);

                    }
                    else
                    {
                        blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 0, mSubreporttype, strSubjectAliases, strDOB, strID, strADDRESS, strAdditionalInformation);

                    }
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;

                        return false;
                    }
                    ////////////////////////////////////////////////////////////Budget Calculation///////////////////////////////

                    if (bolReportFlag == true)
                    {
                        try
                        {
                            strDueDate = "";
                            strflsDueDate = "";
                            // flBudget = float.Parse(oISBLGen.CalculateBudget(oISBLUser.ClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                            flSBudget = float.Parse(oISBLGen.CalculateFlsBudget(oISBLUser.ClientCode, strReportType, strSubReportType, strCountry, strSubjectType, false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isflsCalculationRevert, ref strflsDueDate, intcounter));
                            // SLBudget = oISBL.CalculateFlsBudget(strClientCode, strReportType, dr.Item("SubReportTypeCode"), StrCountry, subtype, chkExpressCase.Checked, False, Session("SubjectDataView"), blnIsCalculationRevert, strDueDate)
                            //oISBLUser.ClientCode,strReportType,strSubReportType,strCountry,strSubjectType
                            dtReportDueDate = DateTime.Parse(strflsDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);
                            /////////////////////////////////////Update Budget/////////////////////////////////////////////////////////////////////////////////////////////
                            blnStatus = UpdateflsBudgetDuedate(ref connPlaceOrder, OrderID, ref oISBLUser, dtReportDueDate, flSBudget, strSubReportType, OrderSubjectID, strSubjectName, strSubjectType);
                            if (!blnStatus)
                            {
                                oISBLGen.DisposeConnection();
                                oISBLGen = null;
                                return false;
                            }
                        }
                        catch (Exception ex)
                        {
                            //Start BI 34  Budget recalculation
                            string xmlData = "<input><clientCode>" + oISBLUser.ClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                            //End BI 34  Budget recalculation
                            oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Budget and Due Date UpdateFlsBudget  : " + ex.Message);
                            oISBLGen.DisposeConnection();
                            oISBLGen = null;

                            //New Addition Start
                            //14/10/2015 on by sanjeeva start
                            if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                                File.Delete(strAttachmentFilePath + strAttachmentFile);
                            //end of sanjeeva
                            //New Addition End

                            return false;
                        }
                    }
                    DataSet dsResearchElements = null;

                    try
                    {
                        if (bolReportFlag == true)
                        {

                            dsResearchElements = oISBLGen.GetResearchElementBySubReportType(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc, strSubReportType);
                        }
                        else
                        {
                            if (strSubReportType != "")
                            {
                                dsResearchElements = oISBLGen.GetResearchElementBySubReportType(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc, strSubReportType);
                            }
                            else
                            {

                                dsResearchElements = oISBLGen.GetResearchElement(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc);
                            }
                        }
                        intRECount = dsResearchElements.Tables[0].Rows.Count;
                    }
                    catch (Exception ex)
                    {
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_6" : "BO_15", "Error Getting Research Elements: " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;


                        return false;
                    }

                    for (int y = 0; y < intRECount; y++)
                    {
                        //Start Adam Jul 2009 IS FCPA
                        if (strReportType == "IS FCPA")
                        {
                            if (x == 0) //Primary
                                strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                            else
                            {
                                if (strSubjectType == "1") //Individual
                                    strResearchElementCode = "30"; //Individual
                                else
                                    strResearchElementCode = "14"; //Company

                                blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                                if (!blnStatus)
                                {
                                    oISBLGen.DisposeConnection();
                                    oISBLGen = null;
                                    return false;
                                }
                                break;
                            }
                        }
                        else
                        {
                            strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                        }
                        //strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString(); //Adam Jul 2009 IS FCPA
                        //End Adam Jul 2009 IS FCPA
                        /////////////////////////////////////Client Order Subject RE       Insertion/////////////////////////////////////////////////////////////////////////////////////////////
                        blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);

                        if (!blnStatus)
                        {
                            oISBLGen.DisposeConnection();
                            oISBLGen = null;

                            return false;
                        }
                    }
                }
               
                /////////////////////////////////////Client Bulk  Order Insertion/////////////////////////////////////////////////////////////////////////////////////////////
                //New Addition Start
                blnStatus = SaveClientBulkOrder(ref connPlaceOrder, BulkSessionID, OrderID);
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    return false;
                }
                //New Addition End
            }
            else // Not Single CRN per Excel/////////////////////////////////////Main Insertion start/////////////////////////////////////////////////////////////////////////////////////////////
            {
                for (int x = 0; x < lstSubjects.Count; x++)
                {
                    OrderID = Guid.NewGuid();
                    dtOrderDateTime = DateTime.Now;

                    hmSubject = (Hashtable)lstSubjects[x];
                    strClientRefNum = hmSubject["ClientRefNum"].ToString();

                    if (oISBLUser.IsImpersonate)
                        blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", oISBLUser.ImpersonateLoginID);
                    else
                        blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, "", strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE");

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        
                        return false;
                    }
                    
                    try
                    {
                        strDueDate = "";
                        flBudget = float.Parse(oISBLGen.CalculateBudget(oISBLUser.ClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                        dtReportDueDate = DateTime.Parse(strDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);
                        
                    }
                    catch (Exception ex)
                    {
                        //Start BI 34 - Budget recalculation
                        //string xmlData = "<input><reportType>" + strReportType + "</reportType></input>";
                        string xmlData = "<input><clientCode>" + oISBLUser.ClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                        //End BI 34 - Budget recalculation
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Report Type Due Days: " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;                                               
                        return false;
                    }

                    blnStatus = SaveClientOrderEmail(ref connPlaceOrder, OrderID, ref oISBLUser);
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }

                    OrderSubjectID = Guid.NewGuid();
                    hmSubject = (Hashtable)lstSubjects[x];
                    strSubjectName = hmSubject["SubjectName"].ToString();
                    strSubjectType = hmSubject["SubjectType"].ToString();
                    strSubReportType = hmSubject["SubreportType"].ToString();

                    strSubjectTypeDesc = "";
                    intCompanyCount = 0;
                    intIndvCount = 0;
                    if (strSubjectType == "1")
                    {
                        intIndvCount++;
                        strSubjectTypeDesc = "Individual";
                    }
                    if (strSubjectType == "2")
                    {
                        intCompanyCount++;
                        strSubjectTypeDesc = "Company";
                    }

                    strCountry = hmSubject["SubjectCountry"].ToString();

                    //New Addition Start
                    strCountryDesc = hmSubject["SubjectCountryDesc"].ToString();
                    //New Addition End

                    strOtherDetails = hmSubject["OtherDetails"].ToString();

                    blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 1, strSubReportType, "", "", "", "", "");

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                       
                        return false;
                    }

                    DataSet dsResearchElements = oISBLGen.GetResearchElement(oISBLUser.ClientCode, strReportType, strSubjectTypeDesc);
                    try
                    {
                        intRECount = dsResearchElements.Tables[0].Rows.Count;
                    }
                    catch (Exception ex)
                    {
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_6" : "BO_15", "Error Getting Research Elements : " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;                        
                        return false;
                    }

                    for (int y = 0; y < intRECount; y++)
                    {
                        strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();//Adam Jul 2009 IS FCPA // FCPA checking not needed here
                                   
                        blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                        if (!blnStatus)
                        {
                            oISBLGen.DisposeConnection();
                            oISBLGen = null;


                            //New Addition Start
                            if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                                File.Delete(strAttachmentFilePath + strAttachmentFile);
                            //New Addition End

                            return false;
                        }
                    }

                    blnStatus = UpdateBudgetDuedate(ref connPlaceOrder, OrderID, ref oISBLUser, dtReportDueDate, flBudget);
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }

                    //New Addition Start
                    blnStatus = SaveClientBulkOrder(ref connPlaceOrder, BulkSessionID, OrderID);

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();

                        oISBLGen = null;
                        return false;
                    }
                    //New Addition End
                }

            }

            oISBLGen.DisposeConnection();
            oISBLGen = null;
            return true;
        }

        //Start BI 20 Refresh Order
        //Saves Case with multiple subjects.  
        private Boolean SaveCase(ref ISDL.Connect connPlaceOrder, ISBL.User oISBLUser, ArrayList lstSubjects, String strReportType, String strAttachmentFilePath, String strAttachmentFile, Boolean isExcel, Guid BulkSessionID, Boolean blnEnableRefreshOrder, String strClientCode, String strOriginalCRN, Guid BatchID)
        {
            ISBL.General oISBLGen = new ISBL.General();
            Hashtable hmSubject;
            int intCompanyCount, intIndvCount;
            Guid OrderID, OrderSubjectID;
            DateTime dtOrderDateTime, dtReportDueDate;
            Boolean blnEnableAutoAssignment, blnEnableSingleCRNPerExcel;
            String strOfficeAssignment, strDueDate, strResearchElementCode;
            int intDueDay;
            int intRECount;
            float flBudget = 0;
            String strSubjectName, strSubjectType, strCountry, strCountryDesc, strClientRefNum, strOtherDetails, strSubjectTypeDesc, strOriginalSubjectID,strSubReportType;
            Boolean blnStatus = true;
            Boolean isCalculationRevert = false; //BI 34
            blnEnableSingleCRNPerExcel = true;

            strClientRefNum = "";

            try
            {
                blnEnableAutoAssignment = ISEnableAutoAssignment(strClientCode);
            }
            catch (Exception ex)
            {
                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_2" : "BO_11", "Error Getting Enable Auto Assignment : " + ex.Message);
                oISBLGen.DisposeConnection();
                oISBLGen = null;
                               
                return false;
            }

            if (blnEnableSingleCRNPerExcel || isExcel)
            {
                OrderID = Guid.NewGuid();
                dtOrderDateTime = DateTime.Now;

                hmSubject = (Hashtable)lstSubjects[0];
                strClientRefNum = hmSubject["ClientRefNum"].ToString();

                intCompanyCount = 0;
                intIndvCount = 0;

                //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam 
                if (oISBLUser.IsImpersonate)
                    blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, strOriginalCRN, strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", oISBLUser.ImpersonateLoginID, blnEnableRefreshOrder, strClientCode, BatchID);
                else
                    blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, strOriginalCRN, strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", blnEnableRefreshOrder, strClientCode, BatchID);

                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                                      
                    return false;
                }
                
                try
                {
                    strDueDate = "";
                    flBudget = float.Parse(oISBLGen.CalculateBudget(strClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                    dtReportDueDate = DateTime.Parse(strDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);
                    
                }
                catch (Exception ex)
                {
                    string xmlData = "<input><clientCode>" + strClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                    oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Budget and Due Date after save Client Order : " + ex.Message);
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    //New Addition Start
                    if (File.Exists(strAttachmentFilePath + strAttachmentFile))
                        File.Delete(strAttachmentFilePath + strAttachmentFile);
                    //New Addition End

                    return false;
                }
                
                blnStatus = SaveClientOrderEmail(ref connPlaceOrder, OrderID, ref oISBLUser);
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                                       
                    return false;
                }

                for (int x = 0; x < lstSubjects.Count; x++)
                {
                    OrderSubjectID = Guid.NewGuid();
                    hmSubject = (Hashtable)lstSubjects[x];
                    strSubjectName = hmSubject["SubjectName"].ToString();
                    strSubjectType = hmSubject["SubjectType"].ToString();
                    strOriginalSubjectID = hmSubject["OriginalSubjectID"].ToString();
                    strSubReportType = hmSubject["SubreportType"].ToString();
                    strSubjectTypeDesc = "";
                    if (strSubjectType == "1")
                    {
                        intIndvCount++;
                        strSubjectTypeDesc = "Individual";
                    }
                    if (strSubjectType == "2")
                    {
                        intCompanyCount++;
                        strSubjectTypeDesc = "Company";
                    }

                    strCountry = hmSubject["SubjectCountry"].ToString();

                    //New Addition Start
                    strCountryDesc = hmSubject["SubjectCountryDesc"].ToString();
                    //New Addition End

                    strOtherDetails = hmSubject["OtherDetails"].ToString();
                    if (x == 0)
                    {
                        blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 1, strSubReportType,"","","","","");

                     }
                    else
                    {

                        blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 0, strSubReportType, "", "", "", "", "");

                     }
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }

                    String REParam = "<ClientCode>" + strClientCode  + "</ClientCode>";
                    REParam  += "<ReportType>" + strReportType + "</ReportType>";
                    REParam += "<SubjectTypeDesc>" + strSubjectTypeDesc + "</SubjectTypeDesc>";
                    REParam += "<OriginalSubjectID>" + strOriginalSubjectID + "</OriginalSubjectID>";
                    DataSet dsResearchElements =  new DataSet();
                    try
                    {
                        dsResearchElements = oISBLGen.GetResearchElement(strClientCode, strReportType, strSubjectTypeDesc, strOriginalSubjectID);
                        intRECount = dsResearchElements.Tables[0].Rows.Count;
                    }
                    catch (Exception ex)
                    {
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, REParam, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_6" : "BO_15", "Error Getting Research Elements: " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;                       
                        return false;
                    }

                    for (int y = 0; y < intRECount; y++)
                    {
                        //Start Adam Jul 2009 IS FCPA
                        if (strReportType == "IS FCPA")
                        {
                            if (x == 0) //Primary
                                strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                            else
                            {
                                if (strSubjectType == "1") //Individual
                                    strResearchElementCode = "30"; //Individual
                                else
                                    strResearchElementCode = "14"; //Company

                                blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                                if (!blnStatus)
                                {
                                    oISBLGen.DisposeConnection();
                                    oISBLGen = null;
                                    return false;
                                }
                                break;
                            }
                        }
                        else
                        {
                            strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();
                        }
                        
                        blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                        if (!blnStatus)
                        {
                            oISBLGen.DisposeConnection();
                            oISBLGen = null;

                            return false;
                        }
                    }
                }
               blnStatus = UpdateRefreshOrderBudgetDuedate(ref connPlaceOrder, OrderID, strClientCode, dtReportDueDate, flBudget);
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                                        
                    return false;
                }

                //New Addition Start
                blnStatus = SaveClientBulkOrder(ref connPlaceOrder, BulkSessionID, OrderID);
                if (!blnStatus)
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    return false;
                }
                //New Addition End
            }
            else // Not Single CRN per Excel
            {
                for (int x = 0; x < lstSubjects.Count; x++)
                {
                    OrderID = Guid.NewGuid();
                    dtOrderDateTime = DateTime.Now;

                    hmSubject = (Hashtable)lstSubjects[x];
                    strClientRefNum = hmSubject["ClientRefNum"].ToString();

                    if (oISBLUser.IsImpersonate)
                        blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, strOriginalCRN, strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", oISBLUser.ImpersonateLoginID, blnEnableRefreshOrder, strClientCode, BatchID);
                    else
                        blnStatus = SaveClientOrder(ref connPlaceOrder, ref oISBLUser, OrderID, strOriginalCRN, strReportType, strClientRefNum, dtOrderDateTime, "", 1, "CRE", blnEnableRefreshOrder, strClientCode, BatchID);

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }                    

                    try
                    {
                        strDueDate = "";
                        flBudget = float.Parse(oISBLGen.CalculateBudget(strClientCode, strReportType, "None", "All", "All", false, true, oISBLGen.ConvertSubjectArrayToDataView(lstSubjects), ref isCalculationRevert, ref strDueDate));
                        dtReportDueDate = DateTime.Parse(strDueDate, System.Globalization.CultureInfo.CreateSpecificCulture("en-AU").DateTimeFormat);
                        //End BI 34 - Budget recalculation
                    }
                    catch (Exception ex)
                    {
                        //Start BI 34 - Budget recalculation
                        string xmlData = "<input><clientCode>" + strClientCode + "</clientCode><reportType>" + strReportType + "</reportType><bulkOrder>True</bulkOrder></input>";
                        //End BI 34 - Budget recalculation
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, xmlData, "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_4" : "BO_13", "Error Getting Report Type Due Days: " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;

                        return false;
                    }
                    
                    blnStatus = SaveClientOrderEmail(ref connPlaceOrder, OrderID, ref oISBLUser);
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }

                    OrderSubjectID = Guid.NewGuid();
                    hmSubject = (Hashtable)lstSubjects[x];
                    strSubjectName = hmSubject["SubjectName"].ToString();
                    strSubjectType = hmSubject["SubjectType"].ToString();
                    strOriginalSubjectID = hmSubject["OriginalSubjectID"].ToString();
                    strSubReportType = hmSubject["SubreportType"].ToString();
                    strSubjectTypeDesc = "";
                    intCompanyCount = 0;
                    intIndvCount = 0;
                    if (strSubjectType == "1")
                    {
                        intIndvCount++;
                        strSubjectTypeDesc = "Individual";
                    }
                    if (strSubjectType == "2")
                    {
                        intCompanyCount++;
                        strSubjectTypeDesc = "Company";
                    }
                    // try catch added by sanjeeva on 18/06/2015
                    try
                    {
                        strCountry = hmSubject["SubjectCountry"].ToString();

                    //New Addition Start
                    strCountryDesc = hmSubject["SubjectCountryDesc"].ToString();
                    //New Addition End

                    strOtherDetails = hmSubject["OtherDetails"].ToString();

                    blnStatus = SaveJpmcClientOrderSubject(ref connPlaceOrder, OrderSubjectID, OrderID, strSubjectName, strSubjectType, strCountry, "", strOtherDetails, 1, strSubReportType, "", "", "", "", "");

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                          return false;
                    }
                    }
                    catch (Exception ex)
                    {
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, "", "", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_15" : "BO_2", "Error in Excel File Generation (Save Case subject country Linenumber-3045): " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;                    

                        return false;
                    }

                    String REParam = "<ClientCode>" + strClientCode + "</ClientCode>";
                    REParam += "<ReportType>" + strReportType + "</ReportType>";
                    REParam += "<SubjectTypeDesc>" + strSubjectTypeDesc + "</SubjectTypeDesc>";
                    REParam += "<OriginalSubjectID>" + strOriginalSubjectID + "</OriginalSubjectID>";
                    DataSet dsResearchElements = new DataSet();                    
                    try
                    {
                        dsResearchElements = oISBLGen.GetResearchElement(strClientCode, strReportType, strSubjectTypeDesc, strOriginalSubjectID);
                        intRECount = dsResearchElements.Tables[0].Rows.Count;
                    }
                    catch (Exception ex)
                    {
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelPlaceOrderMsg : WebPlaceOrderMsg, REParam, " Not Single CRN Per Excel", DateTime.Now, DateTime.Now, oISBLUser.LoginID, oISBLUser.UserType, isExcel ? "EBO_6" : "BO_15", "Error Getting Research Elements: " + ex.Message);
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;

                        return false;
                    }

                    for (int y = 0; y < intRECount; y++)
                    {
                        strResearchElementCode = dsResearchElements.Tables[0].Rows[y][0].ToString();//Adam Jul 2009 IS FCPA // FCPA checking not needed here
                        blnStatus = SaveClientOrderSubjectRE(ref connPlaceOrder, OrderID, OrderSubjectID, strResearchElementCode);
                        if (!blnStatus)
                        {
                            oISBLGen.DisposeConnection();
                            oISBLGen = null;
                            return false;
                        }
                    }

                    blnStatus = UpdateRefreshOrderBudgetDuedate(ref connPlaceOrder, OrderID, strClientCode, dtReportDueDate, flBudget);
                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }

                    //New Addition Start
                    blnStatus = SaveClientBulkOrder(ref connPlaceOrder, BulkSessionID, OrderID);

                    if (!blnStatus)
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        return false;
                    }
                    //New Addition End
                }

            }

            oISBLGen.DisposeConnection();
            oISBLGen = null;
            return true;
        }
        //End BI 20 Refresh Order

        //Gets Subject Type List
        public DataSet GetSubjectType()
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetSubjectType";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetSubjectType";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetSubReportType()
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetSubReportType";
            cmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetSubReportType";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }
        
        //Get budget Parameters for calculating budget
        private DataSet GetBudgetParams(ref ISBL.User oISBLUser, string strReportType, Guid OrderID)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetBudgetParams";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetBudgetParams";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }


        private DataSet GetBudgetParamsBasedOnVariantCountry(ref ISBL.User oISBLUser, string strReportType, string Variant, string Country, bool IsExpress)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetClientVariantSubjectCountryBudget";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
            //cmd.Parameters.Add("@Country", SqlDbType.VarChar, 15); Jul 2013 Adam - ISIS Atlas Data Sync
            cmd.Parameters.Add("@Country", SqlDbType.VarChar,4); //Jul 2013 Adam - ISIS Atlas Data Sync
            cmd.Parameters.Add("@IsExpress", SqlDbType.Bit);
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@Variant"].Value = Variant;
            cmd.Parameters["@Country"].Value = Country;
            cmd.Parameters["@IsExpress"].Value = IsExpress;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetBudgetParamsBasedOnVariantCountry";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return ds;
        }


        private float GetActualBudgetBasedOnVariantCountry(Guid OrderID, String strReportType, ref ISBL.User oISBLUser, int intIndvCount, int intCompanyCount, string Variant, string Country, bool IsExpress)
        {
            DataSet ds;
            ISBL.General oISBLGen = new ISBL.General();
            float flActualBudget, flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerIndividual;
            int intMaxCompanyForBasePrice, intMaxIndividualForBasePrice;
            ds = GetBudgetParamsBasedOnVariantCountry(ref oISBLUser, strReportType, Variant, Country, IsExpress);
            flBasePrice = Convert.ToSingle(ds.Tables[0].Rows[0]["BasePrice"]);
            flAdditionalCostPerCompany = Convert.ToSingle(ds.Tables[0].Rows[0]["AdditionalCostPerCompany"]);
            flAdditionalCostPerIndividual = Convert.ToSingle(ds.Tables[0].Rows[0]["AdditionalCostPerIndividual"]);
            intMaxCompanyForBasePrice = Convert.ToInt32(ds.Tables[0].Rows[0]["MaxCompanyForBasePrice"]);
            intMaxIndividualForBasePrice = Convert.ToInt32(ds.Tables[0].Rows[0]["MaxIndividualForBasePrice"]);
            //Fixing a bug for budget calculation. change flAdditionalCostPerCompany to flAdditionalCostPerIndividual on the 3rd parameter
            //Adam 3 Oct 2008
            //flActualBudget = oISBLGen.GetBudget(flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerCompany, intMaxCompanyForBasePrice, intMaxIndividualForBasePrice, intCompanyCount, intIndvCount);
            flActualBudget = oISBLGen.GetBudget(flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerIndividual, intMaxCompanyForBasePrice, intMaxIndividualForBasePrice, intCompanyCount, intIndvCount);
            oISBLGen.DisposeConnection();
            oISBLGen.DisposeConnection();
            oISBLGen = null;

            return flActualBudget;
        }

        //Gets actual budget by calling method from general class
        private float GetActualBudget(Guid OrderID, String strReportType, ref ISBL.User oISBLUser, int intIndvCount, int intCompanyCount)
        {
            DataSet ds;
            ISBL.General oISBLGen = new ISBL.General();
            float flActualBudget, flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerIndividual;
            int intMaxCompanyForBasePrice, intMaxIndividualForBasePrice;
            ds = GetBudgetParams(ref oISBLUser, strReportType, OrderID);
            flBasePrice = Convert.ToSingle(ds.Tables[0].Rows[0]["BasePrice"]);
            flAdditionalCostPerCompany = Convert.ToSingle(ds.Tables[0].Rows[0]["AdditionalCostPerCompany"]);
            flAdditionalCostPerIndividual = Convert.ToSingle(ds.Tables[0].Rows[0]["AdditionalCostPerIndividual"]);
            intMaxCompanyForBasePrice = Convert.ToInt32(ds.Tables[0].Rows[0]["MaxCompanyForBasePrice"]);
            intMaxIndividualForBasePrice = Convert.ToInt32(ds.Tables[0].Rows[0]["MaxIndividualForBasePrice"]);
            //Fixing a bug for budget calculation. change flAdditionalCostPerCompany to flAdditionalCostPerIndividual on the 3rd parameter
            //Adam 3 Oct 2008
            //flActualBudget = oISBLGen.GetBudget(flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerCompany, intMaxCompanyForBasePrice, intMaxIndividualForBasePrice, intCompanyCount, intIndvCount);
            flActualBudget = oISBLGen.GetBudget(flBasePrice, flAdditionalCostPerCompany, flAdditionalCostPerIndividual, intMaxCompanyForBasePrice, intMaxIndividualForBasePrice, intCompanyCount, intIndvCount);
            oISBLGen.DisposeConnection();
            oISBLGen.DisposeConnection();
            oISBLGen = null;

            return flActualBudget;
        }

        private Boolean ISEnableAutoAssignment(string ClientCode)
        {
            Boolean EnableAutoAssignment;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_ISEnableAutoAssignment";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.ISEnableAutoAssignment";
            EnableAutoAssignment = Convert.ToBoolean(conn.cmdScalarStoredProc(cmd));
            conn.Close();
            cmd.Dispose();
            return EnableAutoAssignment;
        }

        private string GetOfficeAssignment(ref ISBL.User oISBLUser, String strReportType, Boolean blnAutoAssignment, Boolean blnBulkOrder)
        {
            DataSet ds = new DataSet();
            String strOfficeAssignment;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetOfficeAssignment";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@BulkOrder"].Value = blnBulkOrder;
            cmd.Parameters["@AutoAssignment"].Value = blnAutoAssignment;
            SqlDataAdapter sda = new SqlDataAdapter();
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetOfficeAssignment";
            sda.SelectCommand = cmd;
            ds = conn.FillDataSet(sda);
            strOfficeAssignment = ds.Tables[0].Rows[0]["OfficeAssignment"].ToString();
            conn.Close();
            cmd.Dispose();
            sda.Dispose();
            return strOfficeAssignment;
        }

        private Boolean ISEnableSingleCRNPerExcel(ref ISBL.User oISBLUser)
        {
            Boolean EnableSingleCRNPerExcel;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_ISEnableSingleCRNPerExcel";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.ISEnableSingleCRNPerExcel";
            try
            {
                EnableSingleCRNPerExcel = Convert.ToBoolean(conn.cmdScalarStoredProc(cmd));
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
            return EnableSingleCRNPerExcel;
        }

        // OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        private Boolean SaveClientOrder(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode, String strImpersonateLoginID)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientOrderImpersonate";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ImpersonateBy", SqlDbType.VarChar, 15);
            
            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@CRN"].Value = strCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime; //dtOrderDateTime.Date; Adam 10 Apr 2008 R-042
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
            cmd.Parameters["@ImpersonateBy"].Value = strImpersonateLoginID;
            
            
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean SaveClientOrderNonJPMC(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode, String strImpersonateLoginID,string strSubReportType)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientOrderImpersonateNonJPMC";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ImpersonateBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@CRN"].Value = strCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime; //dtOrderDateTime.Date; Adam 10 Apr 2008 R-042
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
            cmd.Parameters["@ImpersonateBy"].Value = strImpersonateLoginID;
            cmd.Parameters["@SubReportType"].Value = strSubReportType;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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


        private Boolean SaveClientOrder(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientOrder";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            
            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@CRN"].Value = strCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime; 
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
           

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean SaveClientOrderNonJPMC(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode,string strSubReportType)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientOrderNonJPMC";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);

            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@CRN"].Value = strCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime;
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
            cmd.Parameters["@SubReportType"].Value = strSubReportType;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        //Start BI 20 ISIS Refresh Order
        private Boolean SaveClientOrder(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strOriginalCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode, String strImpersonateLoginID, Boolean blnEnableRefreshOrder, String strClientCode, Guid BatchID)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientRefreshOrderImpersonate";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@OriginalCRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ImpersonateBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@EnableRefreshOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier);

            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = strClientCode;
            cmd.Parameters["@OriginalCRN"].Value = strOriginalCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime; //dtOrderDateTime.Date; Adam 10 Apr 2008 R-042
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
            cmd.Parameters["@ImpersonateBy"].Value = strImpersonateLoginID;
            cmd.Parameters["@EnableRefreshOrder"].Value = blnEnableRefreshOrder;
            cmd.Parameters["@BatchID"].Value = BatchID;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean SaveClientOrder(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid OrderID, String strOriginalCRN, String strReportType, String strClientReferenceNumber, DateTime dtOrderDateTime, String strSpecialInstruction, int intBulkOrder, String strStatusCode, Boolean blnEnableRefreshOrder, String strClientCode, Guid BatchID)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();

            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientRefreshOrder";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@OriginalCRN", SqlDbType.VarChar, 80);
            cmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@ClientReferenceNumber", SqlDbType.NVarChar, 300);
            cmd.Parameters.Add("@SpecialInstruction", SqlDbType.NText);
            cmd.Parameters.Add("@OrderReceiptDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreatedDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@CreatedBy", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@EnableRefreshOrder", SqlDbType.Bit);
            cmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier);

            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@ClientCode"].Value = strClientCode;
            cmd.Parameters["@OriginalCRN"].Value = strOriginalCRN;
            cmd.Parameters["@ReportType"].Value = strReportType;
            cmd.Parameters["@ClientReferenceNumber"].Value = strClientReferenceNumber;
            cmd.Parameters["@SpecialInstruction"].Value = strSpecialInstruction;
            cmd.Parameters["@OrderReceiptDate"].Value = dtOrderDateTime; //dtOrderDateTime.Date; Adam 10 Apr 2008 R-042
            cmd.Parameters["@BulkOrder"].Value = intBulkOrder;
            cmd.Parameters["@StatusCode"].Value = strStatusCode;
            cmd.Parameters["@CreatedDate"].Value = dtOrderDateTime;
            cmd.Parameters["@CreatedBy"].Value = oISBLUser.LoginID;
            cmd.Parameters["@EnableRefreshOrder"].Value = blnEnableRefreshOrder;
            cmd.Parameters["@BatchID"].Value = BatchID;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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
        //End BI 20 ISIS Refresh Order

        private Boolean SaveJpmcClientOrderSubject(ref ISDL.Connect connPlaceOrder, Guid OrderSubjectID, Guid OrderID, String strSubjectName, String strSubjectType, String strCountry, String strSubjectPosition, String strOtherDetails, int intPrimary,String strSubReportTypeID,String strSubjectAliases,String strDOB,String strID,String strADDRESS, String strAdditionalInformation)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveJpmcClientOrderSubject";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@SubjectName", SqlDbType.NVarChar, 250);
            cmd.Parameters.Add("@SubjectType", SqlDbType.TinyInt);
            cmd.Parameters.Add("@Country", SqlDbType.VarChar, 4);
            cmd.Parameters.Add("@SubjectPosition", SqlDbType.NVarChar, 150);
            cmd.Parameters.Add("@OtherDetails", SqlDbType.NText);
            cmd.Parameters.Add("@Primary", SqlDbType.Bit);
            cmd.Parameters.Add("@SubReportTypeID", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@SubjectAliases", SqlDbType.VarChar, 50);
            cmd.Parameters.Add("@DOB", SqlDbType.VarChar, 20);
            cmd.Parameters.Add("@BLK_DATA_ID", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@ADDRESS", SqlDbType.VarChar, 150);
            cmd.Parameters.Add("@AdditionalInformation", SqlDbType.VarChar, 1000);
       
            cmd.Parameters["@ID"].Value = OrderSubjectID;
            cmd.Parameters["@OrderID"].Value = OrderID;
            cmd.Parameters["@SubjectName"].Value = strSubjectName;
            cmd.Parameters["@SubjectType"].Value = Convert.ToInt16(strSubjectType); //Jul 2013 Adam ISIS - Atlas Data Synch
            cmd.Parameters["@Country"].Value = strCountry;
            cmd.Parameters["@SubjectPosition"].Value = strSubjectPosition;
            cmd.Parameters["@OtherDetails"].Value = strOtherDetails;
            cmd.Parameters["@Primary"].Value = intPrimary;
            cmd.Parameters["@SubReportTypeID"].Value = strSubReportTypeID;
            cmd.Parameters["@SubjectAliases"].Value = strSubjectAliases;
            cmd.Parameters["@DOB"].Value = strDOB;
            cmd.Parameters["@BLK_DATA_ID"].Value = strID;
            cmd.Parameters["@ADDRESS"].Value = strADDRESS;
            cmd.Parameters["@AdditionalInformation"].Value = strAdditionalInformation;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean UpdateBudgetDuedate(ref ISDL.Connect connPlaceOrder, Guid OrderID, ref ISBL.User oISBLUser, DateTime dtReportDueDate, float flBudget)
        {
            String strStatus;
            ISBL.General oISBLGen = new ISBL.General();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_UpdateBudgetDuedate";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
            cmd.Parameters.Add("@Budget", SqlDbType.Float);
            cmd.Parameters.Add("@ReportDueDate", SqlDbType.DateTime);
            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@Currency"].Value = oISBLGen.GetClientCurrency(oISBLUser.ClientCode).Tables[0].Rows[0][0].ToString();
            cmd.Parameters["@Budget"].Value = flBudget;
            cmd.Parameters["@ReportDueDate"].Value = dtReportDueDate;
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return true;
                }
                else
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                oISBLGen.DisposeConnection();
                oISBLGen = null;

                cmd.Dispose();
                return false;
            }
        }
        //ref connPlaceOrder, OrderID, ref oISBLUser, dtReportDueDate, flSBudget, "USD", strSubReportType, OrderSubjectID, strSubjectName);
        private Boolean UpdateflsBudgetDuedate(ref ISDL.Connect connPlaceOrder, Guid OrderID, ref ISBL.User oISBLUser, DateTime dtReportDueDate, float flBudget, string strSubreportType, Guid OrderSubjectID,string strSubjectName,string strSubjectType)
        {
            String strStatus;
            ISBL.General oISBLGen = new ISBL.General();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_UpdateflsBudgetDuedate";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
            cmd.Parameters.Add("@Budget", SqlDbType.Float);
            cmd.Parameters.Add("@ReportDueDate", SqlDbType.DateTime);
            cmd.Parameters.Add("@SubreportType", SqlDbType.VarChar, 15);
            cmd.Parameters.Add("@OrderSubjectID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@SubjectName", SqlDbType.NVarChar, 450);
            cmd.Parameters.Add("@SubjectType", SqlDbType.TinyInt);

            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@Currency"].Value = oISBLGen.GetClientCurrency(oISBLUser.ClientCode).Tables[0].Rows[0][0].ToString();
            cmd.Parameters["@Budget"].Value = flBudget;
            cmd.Parameters["@ReportDueDate"].Value = dtReportDueDate;
            cmd.Parameters["@SubreportType"].Value = strSubreportType;
            cmd.Parameters["@OrderSubjectID"].Value = OrderSubjectID;
            cmd.Parameters["@SubjectName"].Value =strSubjectName;
            cmd.Parameters["@SubjectType"].Value = Convert.ToInt16(strSubjectType);
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return true;
                }
                else
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                oISBLGen.DisposeConnection();
                oISBLGen = null;

                cmd.Dispose();
                return false;
            }
        }
        //Start BI 20 ISIS Refresh Order
        private Boolean UpdateRefreshOrderBudgetDuedate(ref ISDL.Connect connPlaceOrder, Guid OrderID, string strClientCode, DateTime dtReportDueDate, float flBudget)
        {
            String strStatus;
            ISBL.General oISBLGen = new ISBL.General();
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_UpdateBudgetDuedate";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@Currency", SqlDbType.VarChar, 5);
            cmd.Parameters.Add("@Budget", SqlDbType.Float);
            cmd.Parameters.Add("@ReportDueDate", SqlDbType.DateTime);
            cmd.Parameters["@ID"].Value = OrderID;
            cmd.Parameters["@Currency"].Value = oISBLGen.GetClientCurrency(strClientCode).Tables[0].Rows[0][0].ToString();
            cmd.Parameters["@Budget"].Value = flBudget;
            cmd.Parameters["@ReportDueDate"].Value = dtReportDueDate;
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
            if (strStatus.Length != 0)
            {
                if (strStatus == "True")
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return true;
                }
                else
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                    cmd.Dispose();
                    return false;
                }
            }
            else
            {
                oISBLGen.DisposeConnection();
                oISBLGen = null;

                cmd.Dispose();
                return false;
            }
        }
        //End BI 20 ISIS Refresh Order

        private Boolean SaveClientOrderSubjectRE(ref ISDL.Connect connPlaceOrder, Guid OrderID, Guid OrderSubjectID, String strResearchElementCode)
        {
            String strStatus;

            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientOrderSubjectRE";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@OrderSubjectID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ResearchElementCode", SqlDbType.VarChar, 100);
            cmd.Parameters["@OrderID"].Value = OrderID;
            cmd.Parameters["@OrderSubjectID"].Value = OrderSubjectID;
            cmd.Parameters["@ResearchElementCode"].Value = strResearchElementCode;

            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean SaveClientOrderEmail(ref ISDL.Connect connPlaceOrder, Guid OrderID, ref ISBL.User oISBLUser)
        {

                String strStatus;
                ISBL.General oISBLGen = new ISBL.General();
                SqlCommand cmd = new SqlCommand();
                cmd.Connection = connPlaceOrder.Connection;
                cmd.CommandText = "sp_SaveClientOrderEmail";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Transaction = connPlaceOrder.currentTransaction;
                cmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
                cmd.Parameters.Add("@ClientEmail", SqlDbType.VarChar, 200);
                cmd.Parameters["@OrderID"].Value = OrderID;
                //cmd.Parameters["@ClientEmail"].Value = oISBLGen.GetClientEmail(oISBLUser.ClientCode); 'Adam 18-August-2010, suppose to keep case creator email
                cmd.Parameters["@ClientEmail"].Value = oISBLGen.GetLoginIDEmail(oISBLUser.LoginID);

                strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
                if (strStatus.Length != 0)
                {
                    if (strStatus == "True")
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        cmd.Dispose();
                        return true;
                    }
                    else
                    {
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;
                        cmd.Dispose();
                        return false;
                    }
                }
                else
                {
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;
                    cmd.Dispose();
                    return false;
                }          

        }

        //New Addition Start

        private Boolean SaveClientBulkOrderMaster(ref ISDL.Connect connPlaceOrder, ref ISBL.User oISBLUser, Guid BulkOrderID)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientBulkOrderMaster";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
            cmd.Parameters.Add("@CreateDate", SqlDbType.DateTime);

            cmd.Parameters["@ID"].Value = BulkOrderID;
            cmd.Parameters["@ClientCode"].Value = oISBLUser.ClientCode;
            cmd.Parameters["@StatusCode"].Value = "Closed";
            cmd.Parameters["@CreateDate"].Value = DateTime.Now;
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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

        private Boolean SaveClientBulkOrder(ref ISDL.Connect connPlaceOrder, Guid BulkOrderID, Guid OrderID)
        {
            String strStatus;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = connPlaceOrder.Connection;
            cmd.CommandText = "sp_SaveClientBulkOrder";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = connPlaceOrder.currentTransaction;
            cmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@BulkMasterID", SqlDbType.UniqueIdentifier);
            cmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);

            cmd.Parameters["@OrderID"].Value = OrderID;
            cmd.Parameters["@BulkMasterID"].Value = BulkOrderID;
            cmd.Parameters["@StatusCode"].Value = "Closed";
            strStatus = connPlaceOrder.cmdScalarStoredProc(cmd);
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


        private void SendBulkOrderMails(String strFileName, String strClientCode, String strClientName, String strClientEmail, String strReportType, Guid BulkSessionID, bool isExcel)
        {
            ISBL.General oISBLGen = new ISBL.General();
            ISBL.BizLog oISBLBizLog = new ISBL.BizLog();//Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID
            String strNewXMLData;
            // excel flagging started for mail sending by sanjeeva
            if (blnExcelFailed)
            {
                strNewXMLData = "<Parent>" +
                                "<BulkOrderMasterID>" + BulkSessionID + "</BulkOrderMasterID>" +
                                "</Parent>";
                //Save the Error in Event Log
                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelEmailMsg : WebEmailMsg, "", strNewXMLData, DateTime.Now, DateTime.Parse("1/1/1900"), "Admin", "Admin", isExcel ? "EBO_15" : "BO_2", "Excel Creation Failed");
                //sendExcelErrorMail();
            }
            else
            {
                String strEmailSubject;
                String strEmailBody;
                String strBranchEmail, strCaseManagerEmail, strSenderEmail;

                //Get CaseMnager Email
                strCaseManagerEmail = oISBLGen.GetCaseManagerEmail(strClientCode, strReportType);

                //Get Branch Email
                strBranchEmail = oISBLGen.GetBranchEmail(strClientCode, strReportType, oISBLGen.ISEnableAutoAssignment(strClientCode));
                //oISBLGen.DisposeConnection();  Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.
                //oISBLGen = null;  Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.


                //Get Sender's Email Address
                strSenderEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];

                //Start Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.
                //Boolean blnEnableShowEmailSubjectName = false ; //Mitul Feb 13 - ISIS Custom Fields Enhancement
                int intEnableShowEmailSubjectName = 0; //Mitul Feb 13 - ISIS Custom Fields Enhancement

                string strPrimarySubjectName ="";

                DataSet dsClientOrderInfo = oISBLGen.GetInfoByBulkMasterID(strClientCode, BulkSessionID.ToString());
                if (dsClientOrderInfo.Tables[0].Rows.Count > 0)
                {
                    //blnEnableShowEmailSubjectName = Convert.ToBoolean(dsClientOrderInfo.Tables[0].Rows[0]["EnableShowEmailSubjectName"]);
                    intEnableShowEmailSubjectName = Convert.ToInt32(dsClientOrderInfo.Tables[0].Rows[0]["EnableShowEmailSubjectName"]);// Mitul Feb 13 - ISIS Custom Fields Enhancement

                    strPrimarySubjectName = Convert.ToString(dsClientOrderInfo.Tables[0].Rows[0]["SubjectName"]);
                }
                oISBLGen.DisposeConnection();
                oISBLGen = null;
                //START Mitul Feb 13 - ISIS Custom Fields Enhancement
                if (intEnableShowEmailSubjectName == 1 || intEnableShowEmailSubjectName == 2)
                {
                        strEmailSubject = "Confirmation of Order Received: Bulk Order (" + strPrimarySubjectName + ").";
                }
                else
                {
                    strEmailSubject = "Confirmation of Order Received: Bulk Order. ";
                }

                //if (blnEnableShowEmailSubjectName)
                //{
                //    strEmailSubject = "Confirmation of Order Received: Bulk Order (" + strPrimarySubjectName + ")";
                //}
                //else
                ////End Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.
                //{
                //    //Sent Email To Client
                //    strEmailSubject = "Confirmation of Order Received: Bulk Order";
                //}
                //END Mitul Feb 13 - ISIS Custom Fields Enhancement
                //email body changed by sanjeeva on 14/10/2015
                strEmailBody = "Hello,<br><br>" +
                                 "We are pleased to announce that the summary of your order:Standard is ready for download.<br><br>" +
                                 "Click on the link below to download the summary of order.<br><br>" +
                                  "<br><font size='3'><a href='https://uateddonline.thomsonreuters.com/DownloadBulkorderExcel.aspx?BulkID=" + BulkSessionID + "'>https://uateddonline.thomsonreuters.com/DownloadBulkorderExcel.aspx?BulkID=" + BulkSessionID + "</a></font><br><br>" +
                                 //"Attached is the summary of your order.<br><br>" +
                                 "Please do not hesitate to contact us if you have any queries by sending an email to " + strCaseManagerEmail + "." +
                                  "<br><br>Thank You.<br><br>" +
                                 "<br><br>Regards,<br><br>" +
                                //Start Change Email Signed - Feb 2013 Adam
                                //"Global World-Check";
                               
                                 "Thomson Reuters Risk" +
                                 "<br><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font><br><br>";
                                //End Change Email Signed - Feb 2013 Adam

                if (!oISBLBizLog.IsMercuryID(new Guid(BulkSessionID.ToString()), "", "", "BulkMasterID"))//Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID
                { 
                    if (!General.sendemail(strClientEmail, strSenderEmail, strEmailSubject, strEmailBody, true, strFileName))
                    {
                        oISBLGen = new ISBL.General();
                        strNewXMLData = "<Parent>" +
                                        "<RecipientType>Client</RecipientType>" +
                                        "<BulkOrderMasterID>" + BulkSessionID + "</BulkOrderMasterID>" +
                                        "<Email>" + strClientEmail + "</Email>" +
                                        "</Parent>";
                        oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelEmailMsg : WebEmailMsg, "", strNewXMLData, DateTime.Now.Date, DateTime.Parse("1/1/1900"), "Admin", "Admin", isExcel ? "EBO_10" : "BO_1", "Emailing Failed");
                        oISBLGen.DisposeConnection();
                        oISBLGen = null;

                    }
                }
                //START Mitul Feb 13 - ISIS Custom Fields Enhancement
                if (intEnableShowEmailSubjectName == 1 || intEnableShowEmailSubjectName == 2)
                {
                    strEmailSubject = strClientName + " placed a bulk order - " + strPrimarySubjectName + ".";
                }
                else
                {
                    strEmailSubject = strClientName + " placed a bulk order.";
                }
                //Start Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.
                //if (blnEnableShowEmailSubjectName)
                //{
                //    strEmailSubject = strClientName + " placed a bulk order - " + strPrimarySubjectName + ".";
                //}
                //else
                ////End Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement.
                //{
                //    //Sent Email to Case Manager
                //    strEmailSubject = strClientName + " placed a bulk order.";
                //}
                //END Mitul Feb 13 - ISIS Custom Fields Enhancement
                oISBLGen = new ISBL.General();
                strEmailBody = "Hi " + oISBLGen.GetClientCaseManager(strClientCode, strReportType) + "," +
                                 "<br><br>There is a Bulk Order placed by  " + strClientName + "." +
                                 //sanjeeva changes started on 14/10/2015
                                 "<br><br>We are pleased to announce that the summary of your order:Standard is ready for download.<br><br>" +
                                 "Click on the link below to download the summary of order.<br><br>" +
                                  "<br><font size='3'><a href='https://uateddonline.thomsonreuters.com/DownloadBulkorderExcel.aspx?BulkID=" + BulkSessionID + "'>https://uateddonline.thomsonreuters.com/DownloadBulkorderExcel.aspx?BulkID=" + BulkSessionID + "</a></font><br><br>" +
                                 //end of sanjeeva changes
                                 //"<br><br>Please find the bulk order excel file attached.<br>" +
                                 "The cases are currently being automatically created by Savvion.<br>" +
                                  "<br><br>Thank You.<br><br>" +
                                 "<br><br>Regards,<br><br>" +
                                //Start Change Email Signed - Feb 2013 Adam
                                //"Global World-Check"; +
                                 "Thomson Reuters Risk" +
                                 "<br><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font><br><br>";
                                //End Change Email Signed - Feb 2013 Adam
                oISBLGen.DisposeConnection();
                oISBLGen = null;


                if (strClientCode == "J001")
                {
                    strCaseManagerEmail = strCaseManagerEmail + ";" + ConfigurationManager.AppSettings["JPMC_Ops"];
                }

                if (!General.sendemail(strCaseManagerEmail, strSenderEmail, strEmailSubject, strEmailBody, true, strFileName))
                {
                    oISBLGen = new ISBL.General();
                    strNewXMLData = "<Parent>" +
                                    "<RecipientType>CaseManager</RecipientType>" +
                                    "<BulkOrderMasterID>" + BulkSessionID + "</BulkOrderMasterID>" +
                                    "<Email>" + strCaseManagerEmail + "</Email>" +
                                    "</Parent>";
                    oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), strClientCode, "", isExcel ? ExcelEmailMsg : WebEmailMsg, "", strNewXMLData, DateTime.Now, DateTime.Parse("1/1/1900"), "Admin", "Admin", isExcel ? "EBO_10" : "BO_1", "Emailing Failed");
                    oISBLGen.DisposeConnection();
                    oISBLGen = null;

                }
            } // end of Excel falg checking sanjeeva
          
            //Delete the Excel File
            if (File.Exists(strFileName))
                File.Delete(strFileName);
        }
        //email generation for failed cases in bulk order by sanjeeva
        public void sendExcelErrorMail()
        {
            
            String strClientEmail1 = System.Configuration.ConfigurationManager.AppSettings["SendToStandardEmail"]; 
            String strSenderEmail1 = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
            String strEmailSubject1 = "Error in Excel File Generation for client : " + oISBLUser.ClientCode.ToString();
            String strEmailBody1 = "Hello,<br><br>" +
                             "Excel Creation failed during bulk order process<br><br>" +
                             "Please do not hesitate to contact us if you have any queries by sending an email to " + strClientEmail1 + "." +
                             "<br><br>Regards,<br><br>" +
                             "Thomson Reuters Risk" +
                             "<br><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font><br><br>";
            String strFileName1 = "";
            if (!General.sendemail(strClientEmail1, strSenderEmail1, strEmailSubject1, strEmailBody1, true, strFileName1))
            {
                oISBLGen = new ISBL.General();
                String strNewXMLData = "<Parent>" +
                                "<RecipientType>Client</RecipientType>" +
                                "<BulkOrderMasterID>" + Guid.NewGuid().ToString() + "</BulkOrderMasterID>" +
                                "<Email>" + strClientEmail1 + "</Email>" +
                                "</Parent>";
                oISBLGen.SaveEventLog(Guid.NewGuid().ToString(), oISBLUser.ClientCode, "", isExcel ? ExcelEmailMsg : WebEmailMsg, "", strNewXMLData, DateTime.Now.Date, DateTime.Parse("1/1/1900"), "Admin", "Admin", isExcel ? "EBO_10" : "BO_1", "Emailing Failed");
                oISBLGen.DisposeConnection();
                oISBLGen = null;

            }
            

        }
        //end
       // sanjeeva reddy for bulk order mail sending
        public DataSet GetBulkExcelOrderDetails(Guid BulkOrderID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkExcelOrderDetails";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@BulkID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@BulkID"].Value = BulkOrderID;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetBulkExcelOrderDetails";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetJpmcBulkExcelOrderDetails(Guid BulkOrderID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetJpmcBulkExcelOrderDetails";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@BulkID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@BulkID"].Value = BulkOrderID;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetJpmcBulkExcelOrderDetails";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetBulkExcelOrderFilenames(string versiontype)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkExcelOrderFilenames";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@Fversiontype", SqlDbType.VarChar,10);
            myCmd.Parameters["@Fversiontype"].Value = versiontype;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetBulkExcelOrderFilenames";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetBulkExcelOrderClientFilenames(string ClientCode )
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkExcelOrderClientFilenames";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 10);
            myCmd.Parameters["@ClientCode"].Value = ClientCode;


            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.BulkOrder.GetBulkExcelOrderClientFilenames";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }      


        //end of bulk order email sending
    }



        #endregion
}

