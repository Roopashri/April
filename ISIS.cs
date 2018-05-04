using System;
using System.Collections.Generic;
using System.Text;

using System.Configuration;
using ISDL;
using System.Data.SqlClient;
using System.Data;
using System.Threading; 
using System.IO;

using System.Web;

namespace ISBL
{
    public class ISIS
    {
        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        // private SqlDataReader myDataReader;
        //private Boolean blnMLValidation;
        #endregion

        #region Constructors
        public ISIS()
        {
            conn.setConnection("ocrsConnection");
        }
        #endregion

        #region public function
        public String GetDataSetID(String strDSName)
        {
            String strDBID;
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GetDSID";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@strDSName", SqlDbType.VarChar, 1000);
                cmd.Parameters["@strDSName"].Value = strDSName;
                conn.Open();
                conn.callingMethod = "ISBL.ISIS.GetDataSetID";
                strDBID = cmd.ExecuteScalar().ToString();
                return strDBID;
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
           

        }
        public String InsertSearchBeforeLog(String strLoginName, String strDBIDs, String strName, String strSponsor, String strUniversity, String strPassport, String strCountry)
        {
            SqlCommand cmd = new SqlCommand();
            String strSearchID = "";
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_isis_goc_before_searchlog_Insert";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@search_by_login_name", SqlDbType.VarChar, 15);
                cmd.Parameters.Add("@search_at_date_time", SqlDbType.DateTime);
                cmd.Parameters.Add("@search_on_dataset_ids", SqlDbType.VarChar, 100);
                cmd.Parameters.Add("@parameter_name", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@parameter_sponsor_name", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@parameter_university_name", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@parameter_passport_no", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@parameter_country", SqlDbType.VarChar, 50);
                cmd.Parameters["@search_by_login_name"].Value = strLoginName;
                cmd.Parameters["@search_at_date_time"].Value = DateTime.Now;
                cmd.Parameters["@search_on_dataset_ids"].Value = strDBIDs;
                cmd.Parameters["@parameter_name"].Value = strName;
                cmd.Parameters["@parameter_sponsor_name"].Value = strSponsor;
                cmd.Parameters["@parameter_university_name"].Value = strUniversity;
                cmd.Parameters["@parameter_passport_no"].Value = strPassport;
                cmd.Parameters["@parameter_country"].Value = strCountry;
                conn.Open();
                conn.callingMethod = "ISBL.ISIS.InsertSearchBeforeLog";
                strSearchID = cmd.ExecuteScalar().ToString();
                return strSearchID;
            }
            catch (Exception  ex)
            {
                ISBL.User oISBLUser = new ISBL.User();
                string strErrorType = "Error in InsertSearchBeforeLog";
                InsertErrorLog("0", ex.Message, strErrorType);
                string strTo = System.Configuration.ConfigurationManager.AppSettings["GOCRecepientEmail"].ToString();
                string strFrom = "globalonecheck@integrascreen.com";
                string strSubject = "Error Occured in GOC";
                string strBody = "Hello GOC Team </br></br>The following error encountered for the user : " + oISBLUser.LoginName + "</br></br>The details of the error are as follows :</br></br>" + ex.Message;
                SendGOCErrorMail(strTo, strFrom, strSubject, strBody, true);
                return "false";
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
        }

        public String InsertSuccessLog(String strSearchSerialNo, String strSearchType)
        {
            SqlCommand cmd = new SqlCommand();
            String status = "";
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_isis_goc_final_searchlog_insert";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@search_serial_no", SqlDbType.Int);
                cmd.Parameters.Add("@search_type", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@created_date_time", SqlDbType.DateTime);
                cmd.Parameters["@search_serial_no"].Value = Convert.ToInt32(strSearchSerialNo);
                cmd.Parameters["@search_type"].Value = strSearchType;
                cmd.Parameters["@created_date_time"].Value = DateTime.Now;
                conn.Open();
                conn.callingMethod = "ISBL.ISIS.InsertSuccessLog";
                status = cmd.ExecuteScalar().ToString();
                return status;
            }
            catch(Exception ex)
            {
                ISBL.User oISBLUser = new ISBL.User();
                string strErrorType="Error in InsertSuccessLog";
                InsertErrorLog(strSearchSerialNo, ex.Message, strErrorType);
                string strTo = System.Configuration.ConfigurationManager.AppSettings["GOCRecepientEmail"].ToString();   
                string strFrom = "globalonecheck@integrascreen.com";
                string strSubject = "Error Occured in GOC";
                string strBody = "Hello GOC Team </br></br>The following error encountered for the user : " + oISBLUser.LoginName + "</br></br>The details of the error are as follows :</br></br>" + ex.Message;
                SendGOCErrorMail(strTo, strFrom, strSubject, strBody, true);
                return "false";
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
        }

        public String InsertErrorLog(String strSearchSerialNo, String strErrorMesg, String strErrorType)
        {
            SqlCommand cmd = new SqlCommand();
            String status = "";
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_isis_goc_errorlog_Insert";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@search_serial_no", SqlDbType.Int);
                cmd.Parameters.Add("@error_type", SqlDbType.VarChar, 50);
                cmd.Parameters.Add("@error_message", SqlDbType.VarChar, 200);
                cmd.Parameters.Add("@error_date_time", SqlDbType.DateTime);
                cmd.Parameters["@search_serial_no"].Value = Convert.ToInt32(strSearchSerialNo);
                cmd.Parameters["@error_type"].Value = strErrorType;
                cmd.Parameters["@error_message"].Value = strErrorMesg;
                cmd.Parameters["@error_date_time"].Value = DateTime.Now;
                conn.Open();
                conn.callingMethod = "ISBL.ISIS.InsertErrorLog";
                status = cmd.ExecuteScalar().ToString();
                return status;
            }
            catch(Exception ex)
            {
                ISBL.User oISBLUser = new ISBL.User();
                string strTo = System.Configuration.ConfigurationManager.AppSettings["GOCRecepientEmail"].ToString();
                string strFrom = "globalonecheck@integrascreen.com";
                string strSubject = "Error Occured in GOCErrorLog";
                string strBody = "Hello GOC Team </br></br>The following error encountered for the user : " + oISBLUser.LoginName + "</br></br>The details of the error are as follows :</br></br>" + ex.Message;
                SendGOCErrorMail(strTo, strFrom, strSubject, strBody, true);
                return "false";
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
            }
        }
        public void SendGOCErrorMail(String strTo, string strFrom, string strSubject, string strBody, bool IsBodyHTML)
        {
            ISBL.General oISBLGen = new ISBL.General();
            Thread threadSendMails = new Thread(delegate() { oISBLGen.sendemail(strTo, strFrom, strSubject, strBody, IsBodyHTML); });
            threadSendMails.IsBackground = true;
            threadSendMails.Start();
            //oISBLGen = null;
        }
        #endregion
    }


}
