//Codes added by Pravesh on 30/05/08 for GOC part


using System;
using System.Collections.Generic;
using System.Text;

using System.Configuration;
using ISDL;
using System.Data.SqlClient;
using System.Data;

using System.IO;

using System.Web;

namespace ISBL
{
    public class GOC
    {


        #region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
        //private SqlDataReader myDataReader;
        //private Boolean blnMLValidation;
        #endregion

        #region Constructors
        public GOC()
        {
            conn.setConnection("GOCConnection");
        }
        #endregion

        #region public function



        public void DisposeConnection()
        {
            conn.Dispose();
        }


        public DataSet GetCountryList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();

            try
            {
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "GetCountryList";
                myCmd.CommandType = CommandType.StoredProcedure;
                sda.SelectCommand = myCmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetCountryList";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                myCmd.Dispose();
                // ds.Dispose();
            }

        }

        public DataSet FreeSearch(String strName)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Free";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@FullName", SqlDbType.VarChar, 2000);
                cmd.Parameters["@FullName"].Value = strName;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.FreeSearch";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Name(String strFullName, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Name";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@FullName", SqlDbType.VarChar, 2000);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@FullName"].Value = strFullName;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Name";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Global(String strFullName, String strDBID, String strCountry)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Global";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@FullName", SqlDbType.VarChar, 2000);
                cmd.Parameters.Add("@Country", SqlDbType.VarChar, 100);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@FullName"].Value = strFullName;
                cmd.Parameters["@Country"].Value = strCountry;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Global";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Domestic(String strFullName, String strCountry, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Domestic";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@FullName", SqlDbType.VarChar, 2000);
                cmd.Parameters.Add("@Country", SqlDbType.VarChar, 100);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@FullName"].Value = strFullName;
                cmd.Parameters["@Country"].Value = strCountry;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Domestic";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Pass(String strFullName, String strPassport, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Pass";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@FullName", SqlDbType.VarChar, 2000);
                cmd.Parameters.Add("@PassportNo", SqlDbType.VarChar, 100);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@FullName"].Value = strFullName;
                cmd.Parameters["@PassportNo"].Value = strPassport;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Pass";
                ds = conn.FillDataSet(sda);
                conn.Close();
                cmd.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Sponsor(String strSponsor, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Sponsor";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@SponsorName", SqlDbType.VarChar, 1000);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@SponsorName"].Value = strSponsor;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Sponsor";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GlobalSearch_Uni(String strUniversityName, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GlobalSearch_Uni";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@UniversityName", SqlDbType.VarChar, 2000);
                cmd.Parameters.Add("@strDBID", SqlDbType.VarChar, 1000);
                cmd.Parameters["@UniversityName"].Value = strUniversityName;
                cmd.Parameters["@strDBID"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GlobalSearch_Uni";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }
        }

        public DataSet GetBasicDetails(String strEntityID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "stpGetBasicDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntityID", SqlDbType.Int);
                cmd.Parameters["@intEntityID"].Value = Convert.ToInt32(strEntityID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetBasicDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetAssociatedEntities(String strEntityID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "GOCstpGet_AssociatedEntities";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntityID", SqlDbType.Int);
                cmd.Parameters["@intEntityID"].Value = Convert.ToInt32(strEntityID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetAssociatedEntities";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetAdditionalInformation(String strEntityID, String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "stpGetAdditional_Information";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntity_ID", SqlDbType.Int);
                cmd.Parameters.Add("@intDB_ID", SqlDbType.Int);
                cmd.Parameters["@intEntity_ID"].Value = Convert.ToInt32(strEntityID);
                cmd.Parameters["@intDB_ID"].Value = Convert.ToInt32(strDBID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetAdditionalInformation";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetURLDetails(String strEntityID, String strCRN)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "stpGetURLDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntity_ID", SqlDbType.BigInt);
                cmd.Parameters.Add("@strCRN", SqlDbType.VarChar, 255);
                cmd.Parameters["@intEntity_ID"].Value = Convert.ToInt64(strEntityID);
               
                cmd.Parameters["@strCRN"].Value = strCRN;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetURLDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }


        public DataSet GetEntityOtherDetails(String strEntityID, String strDBID, String strCRN)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "GOCstpGetEntityOtherDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntityID", SqlDbType.Int);
                cmd.Parameters.Add("@intDBID", SqlDbType.Int);
                cmd.Parameters.Add("@strCRN", SqlDbType.VarChar, 255);
                cmd.Parameters["@intEntityID"].Value = Convert.ToInt32(strEntityID);
                cmd.Parameters["@intDBID"].Value = Convert.ToInt32(strDBID);
                cmd.Parameters["@strCRN"].Value = strCRN;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetEntityOtherDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                cmd.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetAddressDetails(String strEntityID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GetAddressDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@intEntityID", SqlDbType.Int);
                cmd.Parameters["@intEntityID"].Value = Convert.ToInt32(strEntityID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetAddressDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }


        public DataSet GetPositionDetails(String strEntityID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GetPositionDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@EntityID", SqlDbType.Int);
                cmd.Parameters["@EntityID"].Value = Convert.ToInt32(strEntityID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetPositionDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetEducationDetails(String strEntityID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GetEducationDetails ";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@EntityID", SqlDbType.Int);
                cmd.Parameters["@EntityID"].Value = Convert.ToInt32(strEntityID);
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetEducationDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet GetSourceDetails(String strEntityID, String strDBID, String strCRN)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_GetSourceDetails";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@DBID", SqlDbType.VarChar, 100);
                cmd.Parameters.Add("@CRN", SqlDbType.VarChar, 500);
                cmd.Parameters.Add("@EntityID", SqlDbType.Int);
                cmd.Parameters["@EntityID"].Value = Convert.ToInt32(strEntityID);
                cmd.Parameters["@DBID"].Value = strDBID;
                cmd.Parameters["@CRN"].Value = strCRN;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetSourceDetails";
                ds = conn.FillDataSet(sda);
                conn.Close();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }

        }

        public DataSet ExecuteQuery(String strQuery)
        {
            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandType = CommandType.Text;
                cmd.CommandText = strQuery;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.ExecuteQuery";
                ds = conn.FillDataSet(sda);
                conn.Close();
                cmd.Dispose();
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                // ds.Dispose();
                sda.Dispose();
            }
        }

        public DataSet GetDB_Details(String strDBID)
        {

            DataSet ds = new DataSet();
            SqlCommand cmd = new SqlCommand();
            SqlDataAdapter sda = new SqlDataAdapter();
            try
            {
                cmd.Connection = conn.Connection;
                cmd.CommandText = "sp_ISIS_DBDetail";
                cmd.CommandType = CommandType.StoredProcedure;
                cmd.Parameters.Add("@db_id", SqlDbType.NVarChar, 100);
                cmd.Parameters["@db_id"].Value = strDBID;
                sda.SelectCommand = cmd;
                conn.Open();
                conn.callingMethod = "ISBL.GOC.GetDB_Details";
                ds = conn.FillDataSet(sda);
                return ds;
            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {
                conn.Close();
                cmd.Dispose();
                //ds.Dispose();
                sda.Dispose();
            }

        }


        #endregion
    }
}
