using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;
using System.Xml;
using System.IO;
using System.Configuration;

namespace ISBL
{
    public class WSAPI
    {
        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        #endregion

        #region Constructors
        public WSAPI()
        {
            conn.setConnection("ocrsConnection");
        }
        #endregion


        #region Public Member


        public DataSet listCases(string strLoginID, DateTime fromDate, DateTime toDate, string ReportType, out string strMessage, bool blnSaveEventLog)
        {
            Boolean blnStatus = false;
            DataSet ds = new DataSet();

            try
            {
                
                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_WSAPIlistCases";
                myCmd.CommandType = CommandType.StoredProcedure;

                myCmd.Parameters.Add("@fromDate", SqlDbType.DateTime );
                myCmd.Parameters["@fromDate"].Value = fromDate;

                myCmd.Parameters.Add("@toDate", SqlDbType.DateTime );
                myCmd.Parameters["@toDate"].Value = toDate;

                myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 50);
                myCmd.Parameters["@ReportType"].Value = ReportType;


                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                myCmd.Parameters["@LoginID"].Value = strLoginID;

                SqlDataAdapter sda = new SqlDataAdapter();
                sda.SelectCommand = myCmd;
                conn.Open();
                conn.callingMethod = "ISBL.WSAPI.sp_WSAPIlistCases";
                ds = conn.FillDataSet(sda);
                conn.Close();
                myCmd.Dispose();

                if (ds.Tables[0].Rows.Count > 0)
                {
                    blnStatus = true;
                    strMessage = "<Result>true</Result>";
                    strMessage += "<MessageOut>";
                    strMessage += "<FromDate>" + fromDate + "</FromDate><ToDate>" + toDate + "</ToDate><ReportType>" + ReportType + "</ReportType><Error/>";
                    strMessage += "</MessageOut>";

                }
                else
                {
                    strMessage = "<Result>false</Result>";
                    strMessage += "<MessageOut>";
                    strMessage += "<FromDate>" + fromDate + "</FromDate><ToDate>" + toDate + "</ToDate><ReportType>" + ReportType + "</ReportType><Error>No record exists.</Error>";
                    strMessage += "</MessageOut>";
                    blnStatus = false;
                    ds = null;
                }
            }
            catch (Exception ex)
            {
                strMessage = "<Result>false</Result>";
                strMessage += "<MessageOut>";
                strMessage += "<FromDate>" + fromDate + "</FromDate><ToDate>" + toDate + "</ToDate><ReportType>" + ReportType + "</ReportType><Error>" + ex.Message + "</Error>";
                strMessage += "</MessageOut>";
                blnStatus = false;
                ds = null;
            }

            if (blnSaveEventLog)
            {
                string strOldXMLData = "";
                string strErrorMessage = "";
                string strErrorCode = "APILC_";
                string strLogID = System.Guid.NewGuid().ToString();

                strOldXMLData += "<root>";
                strOldXMLData += "<FromDate>" + fromDate + "</FromDate>";
                strOldXMLData += "<ToDate>" + toDate + "</ToDate>";
                strOldXMLData += "<ReportType>" + ReportType  + "</ToDate>";
                strOldXMLData += "</root>";


                if (!blnStatus)
                {
                    strErrorMessage = "List Cases Failed.";
                    strErrorCode += "1";
                }
                else
                    strErrorCode += "0";

                ISBL.General oGen = new ISBL.General();
                oGen.SaveEventLog(strLogID, "", "", "API List Cases", strOldXMLData, "<root>" + strMessage + "</root>", System.DateTime.Now, System.DateTime.Now, strLoginID, "", strErrorCode, strErrorMessage);
                oGen = null;
            }

            return ds;
        }


        public DataSet checkOrderStatus(string strLoginID, string strPassword, string strCRN, bool blnSaveEventLog)
        {
            Boolean blnStatus;
            blnStatus = false;
            ISBL.General oGen = new ISBL.General();
            DataSet ds = new DataSet();
            string strErrorMessage = "";
            try
            {
                
                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_WSAPIcheckOrderStatus";
                myCmd.CommandType = CommandType.StoredProcedure;

                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                myCmd.Parameters["@LoginID"].Value = strLoginID;

                myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
                myCmd.Parameters["@Password"].Value = strPassword;

                myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
                myCmd.Parameters["@CRN"].Value = strCRN;

                SqlDataAdapter sda = new SqlDataAdapter();
                sda.SelectCommand = myCmd;
                conn.Open();
                conn.callingMethod = "ISBL.WSAPI.checkOrderStatus";
                ds = conn.FillDataSet(sda);
                conn.Close();
                myCmd.Dispose();

                if (ds.Tables[0].Rows.Count > 0)
                {
                    blnStatus = true;
                }
                else
                {
                    blnStatus = false;

                }
            }
            catch (Exception ex)
            {

                strErrorMessage = ex.Message;
                blnStatus = false;
            }

            if (blnSaveEventLog)
            {
                string strOldXMLData = "";
                
                string strErrorCode = "APICOS_";
                string strLogID = System.Guid.NewGuid().ToString();

                strOldXMLData += "<root>";
                strOldXMLData += "<LoginID>" + strLoginID + "</LoginID>";
                strOldXMLData += "<Password>********************</Password>";
                strOldXMLData += "<CRN>" + strCRN + "</CRN>";
                strOldXMLData += "</root>";


                if (!blnStatus)
                {
                    strErrorMessage = "Check Case Status Failed";
                    strErrorCode += "1";
                }
                else
                    strErrorCode += "0";

                oGen.SaveEventLog(strLogID, "", strCRN, "API Check Status", strOldXMLData, "", System.DateTime.Now, System.DateTime.Now, strLoginID, "", strErrorCode, strErrorMessage);

            }
            oGen = null;
            return ds;
        }


        public Boolean checkLogin(string strLoginID, string strPassword)
        {

            Boolean blnStatus;

            try
            {

                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_WSAPILogin";
                myCmd.CommandType = CommandType.StoredProcedure;

                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                myCmd.Parameters["@LoginID"].Value = strLoginID;

                myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
                myCmd.Parameters["@Password"].Value = strPassword;

                conn.Open();
                conn.callingMethod = "ISBL.WSAPI.checkLogin";
                blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
                conn.Close();
                myCmd.Dispose();

                return blnStatus;
            }
            catch
            {
                return false;
            }
        }


        //public Boolean DownloadFinalReport(string strLoginID, string strPassword, string strCRN, out string strMessage, out byte[] bytAttach, out string strFilename, bool blnSaveEventLog)
        //{

        //    Boolean blnStatus = false;
        //    Byte[] b1 = null;

        //    string strVersion = "";
        //    strFilename = "";

        //    General objGen = new General();
        //    BizLog objBizLog = new BizLog();

        //    string strStatusCode="";
            
        //    try
        //    {

        //        if (this.checkOrderStatus(strLoginID, strPassword, strCRN, out strMessage,false))
        //        {

        //            XmlDocument doc = new XmlDocument();
        //            doc.LoadXml("<root>" + strMessage + "</root>");
        //            strStatusCode = doc.GetElementsByTagName("Status")[0].InnerText;

        //            if (strStatusCode == "Completed")
        //            {
        //                DataSet dsReport = objGen.GetFinalReport(null, strCRN);
        //                if (dsReport.Tables[0].Rows.Count > 0)
        //                {
        //                    strFilename = dsReport.Tables[0].Rows[0]["FileName"].ToString();
        //                    strVersion = dsReport.Tables[0].Rows[0]["Version"].ToString();

        //                    string strPreYear = "";
        //                    string strPreCRN = "";
        //                    string[] arr = strCRN.Split('\\');

        //                    strPreCRN = arr[0].ToString();
        //                    strPreYear = arr[3].ToString().Trim();

        //                    if (File.Exists(ConfigurationManager.AppSettings["LocalFinalReportPath"] + strPreYear + "\\" + strPreCRN + "\\" + strFilename))
        //                    {
        //                        b1 = File.ReadAllBytes(ConfigurationManager.AppSettings["LocalFinalReportPath"] + strPreYear + "\\" + strPreCRN + "\\" + strFilename);
        //                    }
        //                    else
        //                    {
        //                        //Download Final Report from Savvion
        //                        WSResearchBPM.RBPMWebServiceInterfaceService oWSBPM = new WSResearchBPM.RBPMWebServiceInterfaceService();
        //                        b1 = oWSBPM.downloadOnlineReport(strCRN, strFilename, float.Parse(strVersion));
        //                    }

        //                    if (b1 != null)
        //                        blnStatus = true;
        //                    else
        //                    {
        //                        blnStatus = false;
        //                        strMessage = "File not fouind.";
        //                    }
        //                }
        //                else
        //                {
        //                    blnStatus = false;
        //                    strMessage = "File not fouind.";
        //                }
        //            }
        //            else
        //            {
        //                blnStatus = false;
        //                strMessage = "Final Report not yet ready.";
        //            }
        //        }
        //        else
        //        {
        //            blnStatus = false;

        //            XmlDocument doc = new XmlDocument();
        //            doc.LoadXml("<root>" + strMessage + "</root>");
        //            strMessage = doc.GetElementsByTagName("Error")[0].InnerText;
        //        }
        //    }
        //    catch (Exception Ex)
        //    {
        //        blnStatus = false;
        //        strMessage = Ex.Message;
        //    }
        //    finally
        //    {
        //        objGen = null;
        //        objBizLog = null;
        //    }

        //    bytAttach = b1;

        //    if (blnSaveEventLog)
        //    {

        //        string strOldXMLData = "";
        //        string strErrorMessage = "";
        //        string strErrorCode = "APIDFR_";
        //        string strLogID = System.Guid.NewGuid().ToString();

        //        strOldXMLData += "<root>";
        //        strOldXMLData += "<LoginID>" + strLoginID + "</LoginID>";
        //        strOldXMLData += "<Password>" + strPassword + "</Password>";
        //        strOldXMLData += "<CRN>" + strCRN + "</CRN>";
        //        strOldXMLData += "</root>";


        //        if (!blnStatus)
        //        {
        //            strErrorMessage = "Download Final Report Failed";
        //            strErrorCode += "1";
        //        }
        //        else
        //            strErrorCode += "0";

        //        ISBL.General oGen = new ISBL.General();
        //        oGen.SaveEventLog(strLogID, "", strCRN, "API Download Final Report", strOldXMLData, "<root>" + strMessage + "</root>", System.DateTime.Now, System.DateTime.Now, strLoginID, "Client", strErrorCode, strErrorMessage);
        //        oGen = null;
        //    }

        //    return blnStatus;

        //}

        #endregion
    }
}
