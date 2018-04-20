/****************************************************
 *	File Name:	ISDL.dll
 *	Class Name: Connect
 *	Description:
 *		Public Members of IsOpen(), Open() and Close() - 
 *		each of these items will do it's seperate function
 *		IsOpen():
 *			Determins if the Connection State of m_conn is Open returns
 *			a boolean for determination.
 * 
 *		Open():
 *			Firsts check to see if the connection state of m_conn is open,
 *			then if it's not it opens the connection of m_conn,
 *			returns true to say that it is open.
 * 
 *		Close():
 *			First determins if the connection state is open or closed,
 *			If it's open then it will Close the connection
 *			Returns true to say that the connection has been closed
 * 
 ****************************************************/

using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Permissions;
using System.Security.Principal;
using System.Configuration;



namespace ISDL
{
	/// <summary>
	/// Database connection manipulation.
	/// </summary>
	public class Connect
	{
		#region Private Variables

		private string database;		// Database name inputed from program.
		private string server;			// Server name - Is set up in the setup program, and is recalled by the program
		private string sapwd;			// The Sql Authentication Password for user SA.
		private SqlConnection m_conn;   // Sql Database Connection parameter.
		private string connectionString; // The string used by m_conn for the DB Connection

		#endregion

        #region Public Properties
        /* Mudassar Start */
        public SqlTransaction currentTransaction = null;
        /* Mudassar End */

        /* Adam Start 12 Sep 2008- To include class name when calling ISDL.connect */
        /* If an exception is raised, it will be recorded in the Event Viewer, which class is making the connection*/
        public String callingMethod = null;
        /* Adam End*/
        #endregion

        #region Constructors

        /// <summary>
		/// Standard Constructor - No Inputs, allows to set inputs by calling ConDb for\n
		/// Connect Database; ConServ for Connection.server; ConSA for Connection.sapwd;
		/// </summary>
		public Connect()
		{

		}

		#endregion

		#region Public Members

		/// <summary>
		/// Determine if the Connection State of m_conn is open.\n
		/// No Inputs.
		/// </summary>
		/// <returns>true if state = open; false if state = closed;</returns>
		public bool IsOpen()
		{
			if (m_conn.State == ConnectionState.Open)
			{
				return true;
			}
			else
			{
				return false;
			}
		}

		/// <summary>
		/// Opens the database connection.
		/// </summary>
		/// <returns>boolean true</returns>
		public bool Open()
		{
			bool Opened = false;
			try
			{
				if (IsOpen())
				{
					Opened = true;
				}
				else
				{
					m_conn.Open();
					Opened = true;
				}
				return Opened;
			}
			catch(Exception ex)
			{
				this.errTrack(ex);
			}
			return Opened;
		}

		/// <summary>
		/// Closes the connection to the database.
		/// </summary>
		/// <returns>boolean true;</returns>
		public bool Close()
		{
			bool Closed = false;
			try
			{
				if (IsOpen())
				{
					m_conn.Close();
					Closed = true;
				}
				else
				{
					Closed = true;
				}
			}
			catch(Exception ex)
			{
               this.errTrack(ex);
			}
			return Closed;
		}


        public void Dispose()
        {
            try
            {
                m_conn.Dispose();
            }
            catch (Exception ex)
            {
                this.errTrack(ex);
            }
        }

		public void setConnection(string connName)
		{
            connectionString = ConfigurationManager.ConnectionStrings[connName].ConnectionString;
   		    m_conn = new SqlConnection(connectionString);
		}

        /// <summary>
        /// Use cmdNoneQuery to execute Insert, Update and Delete Statement.
        /// </summary>
        /// <param name="myCmd"></param>
        /// <returns>Boolean</returns>
        public Boolean cmdNoneQuery(SqlCommand myCmd)
        {
            try
            {
                myCmd.CommandTimeout = 300; //5 minutes
                myCmd.ExecuteNonQuery();
                return true;
            }
            catch (Exception ex)
            {
                this.errTrack(ex);
                return false ;
            }
        }

        /// <summary>
        /// Use cmdNoneQuery to execute Insert, Update and Delete Statement.
        /// </summary>
        /// <param name="myCmd"></param>
        /// <returns>Boolean</returns>
        public String  cmdScalarStoredProc(SqlCommand myCmd)
        {
            try
            {
                string strResult;//Added By Adam 18 Mar 08
                myCmd.CommandTimeout = 300; //5 minutes
                strResult = myCmd.ExecuteScalar().ToString();
                myCmd.Dispose();//Added By Adam 18 Mar 08
                return strResult;

            }
            catch (Exception ex)
            {
                this.errTrack(ex);
                myCmd.Dispose(); //Added By Adam 18 Mar 08
                return "";
            }
        }

        /// <summary>
        ///Use DataSetStoredProcWthParam where Parameter can be come from user input. 
        /// </summary>
        /// <returns>returns DataSet</returns>
        public DataSet FillDataSet(SqlDataAdapter myDataAdapter)
        {
            DataSet myDataSet = new DataSet();
            try
            {
                myDataAdapter.Fill(myDataSet, "Table");
                myDataAdapter.Dispose(); //Added by Adam 18 Mar 08
                return myDataSet;
            }
            catch (Exception ex)
            {
                this.errTrack(ex);
                myDataAdapter.Dispose(); //Added by Adam 18 Mar 08
                myDataSet.Dispose();
                return null;
            }
        }
        
        /// <summary>
        ///Use DataSetStoredProc where commandText is not from user input. 
        ///The stored procedure does not require input parameter.
        /// </summary>
        /// <returns>returns DataSet</returns>
        public DataSet DataSetStoredProc(string commandText, ref SqlConnection sqlConnect)
        {   
            DataSet myDataSet = new DataSet();
            SqlDataAdapter myDataAdapter = new SqlDataAdapter();
            try
            {
                myDataAdapter = new SqlDataAdapter(commandText, sqlConnect);
                myDataAdapter.Fill(myDataSet, "Table");
                myDataAdapter.Dispose(); //Added by Adam 18 Mar 08
                return myDataSet;
            }
            catch (Exception ex)
            {
                this.errTrack(ex);
                myDataSet.Dispose();
                myDataAdapter.Dispose();
                return null;
            }
        }

        /// <summary>
        ///Use cmdScalar where commandText is not from user input and only return one row and one column.
        /// </summary>
        /// <returns>returns string</returns>
        public string cmdScalar(string commandText, ref SqlConnection sqlConnect)
        {

                string tmpScalarValue;
                try
                {
                    SqlCommand cmd = new SqlCommand(commandText, sqlConnect);
                    cmd.CommandTimeout = 300; //5 minutes
                    tmpScalarValue = cmd.ExecuteScalar().ToString();
                    cmd.Dispose();
                    return tmpScalarValue;
                }
                catch (Exception ex)
                {
                    this.errTrack(ex);
                    return "";
                }

        }
        /// <summary>
        ///Use cmdReader where commandText is not from user input. It returns multiple rows and columns.
        /// </summary>
        /// <returns>returns SqlDataReader</returns>
        public SqlDataReader cmdReader(string commandText, ref SqlConnection sqlConnect)
        {
            SqlDataReader tmpDataReader = null;

            try
            {
                    SqlCommand cmd = new SqlCommand(commandText, sqlConnect);
                    cmd.CommandTimeout = 300; //5 minutes
                    tmpDataReader = cmd.ExecuteReader(CommandBehavior.CloseConnection);
                    cmd.Dispose();
                    return tmpDataReader;
                }
                catch (Exception ex)
                {
                    this.errTrack(ex);
                    return tmpDataReader;
                }

        }

        /// <summary>
        ///Use cmdReader where commandText is not from user input. It returns multiple rows and columns.
        /// </summary>
        /// <returns>returns SqlDataReader</returns>
        public SqlDataReader cmdReader(SqlCommand cmd)
        {
            SqlDataReader tmpDataReader = null;

            try
            {
                cmd.CommandTimeout = 300; //5 minutes
                tmpDataReader = cmd.ExecuteReader();
                cmd.Dispose();
                return tmpDataReader;
            }
            catch (Exception ex)
            {
                this.errTrack(ex);
                return tmpDataReader;
            }

        }

        public string SafeSqlLiteral(string inputSQL)
        {
            return inputSQL.Replace("'", "''");
        }


        //-------------------Mudassar---Start---------------------------//
        public void BeginTransaction()
        {
            currentTransaction = m_conn.BeginTransaction(IsolationLevel.ReadCommitted);
        }
        public void CommitTransaction()
        {
            currentTransaction.Commit();
        }
        public void RollBackTransaction()
        {
            currentTransaction.Rollback();
        }
        //-------------------Mudassar---End-----------------------------//


		#endregion

		#region Private Members

        //private void errTrack(Exception ex)
        //{
        //    string[] computer = WindowsIdentity.GetCurrent().Name.ToString().Split('\\');

        //    if (!EventLog.SourceExists("OCRS_sqlConnection"))
        //    {
        //        EventLog.CreateEventSource("OCRS_sqlConnection", "OCRS DB Error");
        //    }

        //    EventLog evntLog = new EventLog();
        //    evntLog.Source = "OCRS_sqlConnection";
        //    evntLog.WriteEntry("Error 1005\n\n" + ex.Message, EventLogEntryType.Error);
        //    evntLog.Close();
        //}

        private void errTrack(Exception ex)
        {
            string[] computer = WindowsIdentity.GetCurrent().Name.ToString().Split('\\');
            /* Abhijit Start Change Apr 1, 2008 */
            string appName = "";
            string sectionName = "";
            //Read Settings from web.config
            try
            {
                appName = System.Configuration.ConfigurationSettings.AppSettings["appName"].ToString().Trim();
                sectionName = System.Configuration.ConfigurationSettings.AppSettings["sectionName"].ToString().Trim();
            }
            catch //If settings not present in app.config then load default
            {
                appName = "OCRS_sqlConnection";
                sectionName = "OCRS DB Error";
            }
            if (!EventLog.SourceExists(appName))
            {
                EventLog.CreateEventSource(appName, sectionName);
            }
            EventLog evntLog = new EventLog();
            evntLog.Source = appName;
            evntLog.WriteEntry("Error 1005\n\n" + ex.Message + " (" + callingMethod + ")", EventLogEntryType.Error);
            evntLog.Close();
            /* Abhijit End Change Apr 1, 2008 */
        }


		#endregion

		#region Public Delegates

		/// <summary>
		/// Returns the SqlConnection m_conn
		/// </summary>
		public SqlConnection Connection
		{
			get{ return m_conn; }
		}
		
		/// <summary>
		/// Gets or sets the value for the database name
		/// </summary>
		public string Database
		{
			get{ return database; }
			set{ database = value; }
		}

		/// <summary>
		/// gets or sets the value for the server name
		/// </summary>
		public string Server
		{
			get{ return server; }
			set{ server = value; }
		}

		/// <summary>
		/// Gets or sets the value for the SAPWD password
		/// </summary>
		public string Sapwd
		{
			get{ return sapwd; }
			set{ sapwd = value; }
		}

		#endregion
	}
}
