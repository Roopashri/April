using System;
using System.Collections.Generic;
using System.Text;
using System.Data.SqlClient;
using System.Data;
using System.Net.Mail;
using System.IO;
using System.Xml;
using System.Configuration;
using System.Diagnostics;
using System.Security.Principal;


namespace ISBL
{
    public class RefreshOrder
    {

#region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
#endregion

#region Constructors
        public RefreshOrder()
        {
            conn.setConnection("ocrsConnection");
        }
#endregion

        //Closes db connection.
        public void DisposeConnection()
        {
            conn.Dispose();
        }

        public DataSet GetRefreshOrderCriteriaMaster()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderCriteriaMaster";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderCriteriaMaster";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRefreshOrderPDFMaster()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderPDFMaster";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderPDFMaster";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRefreshOrderReportTypePDFByClient(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderReportTypePDFByClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderReportTypePDFByClient";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public string InsertClientReportPDF(string srClientCode, string strReportType, string strPDF)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_InsertClientReportPDF";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = srClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters.Add("@PDF", SqlDbType.VarChar, 10);
            myCmd.Parameters["@PDF"].Value = strPDF;

            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.InsertClientReportPDF";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }

        public DataSet GetRefreshOrderCountryPDFByClient(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderCountryPDFByClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderCountryPDFByClient";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public string GetCountryMasterPDF()
        {
            string strPDF;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryMasterPDF";
            myCmd.CommandType = CommandType.StoredProcedure;

            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetCountryMasterPDF";
            strPDF = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strPDF;
        }

        public string InsertClientCountryPDF(string srClientCode, string strCountryCode, string strPDF)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_InsertClientCountryPDF";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = srClientCode;
            myCmd.Parameters.Add("@CountryCode", SqlDbType.VarChar, 4);
            myCmd.Parameters["@CountryCode"].Value = strCountryCode;
            myCmd.Parameters.Add("@PDF", SqlDbType.VarChar, 10);
            myCmd.Parameters["@PDF"].Value = strPDF;

            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.InsertClientCountryPDF";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strStatus;
        }

        public string GetDecryptedRefreshOrderNotificationBatchID(string strEncryptedBatchID, string strLoginID)
        {
            string strDecryptedBatchID;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetDecryptedRefreshOrderNotificationBatchID";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@BatchID", SqlDbType.VarChar, 100);
            myCmd.Parameters["@BatchID"].Value = strEncryptedBatchID;

            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;

            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetDecryptedRefreshOrderNotificationBatchID";
            strDecryptedBatchID = conn.cmdScalarStoredProc(myCmd);
            conn.Close();

            myCmd.Dispose();

            return strDecryptedBatchID;
        }

        public DataSet GetRefreshOrderDetails(string strBatchID, string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderDetails";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@BatchID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@BatchID"].Value = new Guid(strBatchID);
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderDetails";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRefreshOrderDetailsByCRN(string strXMLCRN, string strLoginID, Boolean blnReturnResult)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderDetailsByCRN";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@XMLCRN", SqlDbType.Text);
            myCmd.Parameters["@XMLCRN"].Value = strXMLCRN;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters.Add("@ReturnResult", SqlDbType.Bit);
            myCmd.Parameters["@ReturnResult"].Value = blnReturnResult;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.RefreshOrder.GetRefreshOrderDetailsByCRN";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

    }
}
