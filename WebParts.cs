using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.Data.SqlClient;

namespace ISBL
{
    public class WebParts 
    {

#region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config

#endregion

#region Constructors
        public WebParts()
        {
            conn.setConnection("ocrsConnection");
        }
#endregion

#region Public Member
        public DataSet GetOrderSummary()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOrderSummary";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.WebParts.GetOrderSummary";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetEmailStatus()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEmailStatus";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.WebParts.GetOrderSummary";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
#endregion

    }
}
