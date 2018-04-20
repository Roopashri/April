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
    public class General
    {

#region Private Variables
        private ISDL.Connect conn = new ISDL.Connect(); //Return the connection string from web config
#endregion

#region Constructors
        public General()
        {
            conn.setConnection("ocrsConnection");
        }
#endregion


#region Public Member

        public String EncryptData(string ClearText)
        {
            return DataEncryption.Encrypt(ClearText);
        }


        public String DecryptData(string CipherText)
        {
            return DataEncryption.Decrypt(CipherText);
        }

        public Boolean ValidateUser(string LoginID, string Password)
        {
                        return true;
        }

        public Boolean LoginInstanceExist(string LoginID)
        {
            return true;
        }

        public Boolean CreateNewLoginInstance(string LoginID, System.Guid InstanceGUID)
        {
            return true;
        }
        
        //Closes db connection.
        public void DisposeConnection()
        {
            conn.Dispose();
        }

        public String convertCurrency(String strCurrency)
        {
            string strResult = strCurrency;
            switch (strCurrency)
            {
                case "SND":
                    strResult =  "SGD";
                    break;
            }
            return strResult;
        }

        public Boolean UpdateLoginPassword(string strLoginID, string strPassword, Boolean blnSendEmail)
        {
            string strClientContactEmail;
            string strBody;
            string strEncPassword;
            ISBL.BizLog oISBLBizLog = new ISBL.BizLog(); //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID

            strEncPassword = this.EncryptData(strPassword);

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateLoginPassword";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
            myCmd.Parameters["@LoginID"].Value = this.SafeSqlLiteral(strLoginID);
            myCmd.Parameters["@Password"].Value = strEncPassword;
            conn.Open();
            conn.callingMethod = "ISBL.General.UpdateLoginPassword";
            strClientContactEmail = conn.cmdScalarStoredProc(myCmd);
      
            conn.Close();

            if (strClientContactEmail.Length != 0)
            {
                if (blnSendEmail)
                {
                    strBody = "<Table><Tr><Td>Your password has been changed to: " + strPassword + ".</Td></Tr>";
                    strBody += "<Tr><Td>LoginID: " + strLoginID + "</Td></Tr>";
                    strBody += "<Tr><Td height='10'></Td></Tr>";
                    strBody += "<Tr><Td height='50'>Thank you.</Td></Tr>";
                    //Start Change Email Signed - Feb 2013 Adam
                    //strBody += "<Tr><Td height='30'>Global World-Check</Td></Tr>";
                    strBody += "<Tr><Td height='30'>Thomson Reuters Risk";
                    strBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></Td></Tr>";
                    strBody += "<Tr><Td height='20'></Td></Tr>";
                    //End Change Email Signed - Feb 2013 Adam
                    strBody += "</Table>";

                    //Email to Login User
                    myCmd.Dispose();

                    if (ISADFSLoginID(strLoginID)) //check if ADFS Login ID, do not send email
                        return true;
                    else
                    {
                        if (!oISBLBizLog.IsMercuryID(new Guid(Guid.NewGuid().ToString()), "", strLoginID, "LoginID")) //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID
                            return sendemail(strClientContactEmail, ConfigurationManager.AppSettings["SendFromStandardEmail"].ToString(), "Password Changed", strBody, true);
                        else
                            return true;
                    }
                }
                else
                {
                    myCmd.Dispose();
                    return true;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }

       //Password Enhancement OCT 2012 - overload
        public Boolean UpdateLoginPassword(string strLoginID, string strPassword, Boolean blnSendEmail, Boolean blnForceChangePassword)
        {
            string strClientContactEmail;
            string strBody;
            string strEncPassword;
            ISBL.BizLog oISBLBizLog = new ISBL.BizLog(); //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID

            strEncPassword = this.EncryptData(strPassword);

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateLoginPasswordEnhanced";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 250);
            myCmd.Parameters.Add("@ForceChangePassword", SqlDbType.Bit);
            myCmd.Parameters["@LoginID"].Value = this.SafeSqlLiteral(strLoginID);
            myCmd.Parameters["@Password"].Value = strEncPassword;
            myCmd.Parameters["@ForceChangePassword"].Value = blnForceChangePassword;
            conn.Open();
            conn.callingMethod = "ISBL.General.UpdateLoginPasswordEnhanced";
            strClientContactEmail = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strClientContactEmail.Length != 0)
            {
                if (blnSendEmail)
                {
                    //strBody = "<Table><Tr><Td>Your password has been changed to: " + strPassword + ".</Td></Tr>";
                    //strBody += "<Tr><Td>LoginID: " + strLoginID + "</Td></Tr>";
                    //strBody += "<Tr><Td height='10'></Td></Tr>";
                    //strBody += "<Tr><Td height='50'>Thank you</Td></Tr>";
                    //strBody += "<Tr><Td height='30'>Global World-Check</Td></Tr>";
                    //strBody += "</Table>";

                    strBody = "<Table><Tr><Td>Your password for the Login ID " + strLoginID + " has been changed by the Administrator to: " + strPassword + "</Td></Tr>";
                    strBody += "<Tr><Td height='10'></Td></Tr>";
                    strBody += "<Tr><Td>You will need to choose a new password when you login for the first time. Click <a href='" + ConfigurationManager.AppSettings["ISISURL"].ToString() + "login.aspx?open=true'>here</a> to choose a new password.</Td></Tr>";
                    strBody += "<Tr><Td height='10'></Td></Tr>";
                    strBody += "<Tr><Td height='50'>Thank you.</Td></Tr>";
                    //Start Change Email Signed - Feb 2013 Adam
                    //strBody += "<Tr><Td height='30'>Global World-Check</Td></Tr>";
                    strBody += "<Tr><Td height='30'>Thomson Reuters Risk";
                    strBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></Td></Tr>";
                    strBody += "<Tr><Td height='20'></Td></Tr>";
                    //End Change Email Signed - Feb 2013 Adam
                    strBody += "</Table>";

                    //Email to Login User
                    myCmd.Dispose();

                    if (ISADFSLoginID(strLoginID)) //check if ADFS Login ID, do not send email
                        return true;
                    else
                    {
                        if (!oISBLBizLog.IsMercuryID(new Guid(Guid.NewGuid().ToString()), "", strLoginID, "LoginID")) //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID
                            return sendemail(strClientContactEmail, ConfigurationManager.AppSettings["SendFromStandardEmail"].ToString(), "Password Changed", strBody, true);
                        else
                            return true;
                    }
                }
                else
                {
                    myCmd.Dispose();
                    return true;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }

        public Boolean ForgotPassword(string strLoginID)
        {

            string strPassword;
            string strClientContactEmail;
            string strBody;
            string strWebServiceURL;
            string strQueryString;
            string strEncQueryString;
            string strEncPassword;
            Boolean blnSaveEvent;
            Boolean blnSendEmail;
            string strAction;
            string strActionBy;
            string strUserType;
            string strErrorCode;
            string strErrorMessage;
            string strNewXMLData;
            ISBL.BizLog oISBLBizLog = new ISBL.BizLog(); //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID

            strPassword = RandomizeCharacters(7);
            
            //Get LoginID Email Address
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetLoginIDEmail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = this.SafeSqlLiteral(strLoginID);

          
            conn.Open();
            conn.callingMethod = "ISBL.General.ForgotPassword.sp_GetLoginIDEmail";


            strClientContactEmail = conn.cmdScalarStoredProc(myCmd);

           
            conn.Close();
            if (strClientContactEmail.Length != 0)
            {
                try
                {
                    //Get WebServie URL
                    SqlCommand myCmdWS = new SqlCommand();
                    myCmdWS.Connection = conn.Connection;
                    myCmdWS.CommandText = "sp_GetWebServiceURL";
                    myCmdWS.CommandType = CommandType.StoredProcedure;
                    myCmdWS.Parameters.Add("@WebServiceID", SqlDbType.VarChar, 50);
                    myCmdWS.Parameters["@WebServiceID"].Value = "UFP";
                    conn.Open();
                    conn.callingMethod = "ISBL.General.ForgotPassword.sp_GetWebServiceURL";
                    strWebServiceURL = conn.cmdScalarStoredProc(myCmdWS);
                    conn.Close();

                    //Encrypt the Web Service Query String
                    strQueryString = strLoginID + "&" + strPassword;
                    strEncQueryString = this.EncryptData(strQueryString);

                    strEncQueryString = strEncQueryString.Replace("+", "~");

                    //while (strEncQueryString.Contains("+"))
                    //{
                    //    strPassword = RandomizeCharacters(4);
                    //    strQueryString = strLoginID + "&" + strPassword;
                    //    strEncQueryString = this.EncryptData(strQueryString);
                    //}

                    strEncPassword = this.EncryptData(strPassword);                  

                    strBody = "<Table><Tr><Td colspan=2>You have requested for a password reset for your Login ID " + strLoginID + ".</Td></Tr>";
                    strBody += "<Tr><Td height='10'colspan=2></Td></Tr>";
                    strBody += "</Table>";
                    strBody += "<Table>";
                    strBody += "<Tr><Td colspan=2>Click on the link below to initiate the password reset process.</Td></Tr>";
                    strBody += "<Tr><Td colspan=2><a href='" + strWebServiceURL + strEncQueryString + "'>" + strWebServiceURL + strEncQueryString + "</a></Td></Tr>";
                    strBody += "<Tr><Td colspan=2>(If the link does not work, please copy the full address above and paste it to your Internet browser)</Td></Tr>";
                    strBody += "<Tr><Td height='10'colspan=2></Td></Tr>";
                    strBody += "<Tr><Td height='10'colspan=2></Td></Tr>";
                    strBody += "<Tr><Td colspan=2>If you did not request for your password to be reset, please ignore this email.</Td></Tr>";
                    strBody += "<Tr><Td colspan=2>Be assured that your login details are very secure.</Td></Tr>";
                    strBody += "<Tr><Td height='50' colspan=2>Thank you.</Td></Tr>";
                    //Start Change Email Signed - Feb 2013 Adam
                    //strBody += "<Tr><Td height='30' colspan=2>Global World-Check</Td></Tr>";
                    strBody += "<Tr><Td height='30' colspan=2>Thomson Reuters Risk";
                    strBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></Td></Tr>";
                    strBody += "<Tr><Td height='20'colspan=2></Td></Tr>";
                    //End Change Email Signed - Feb 2013 Adam
                    strBody += "</Table>";

                    //Email to Login User

                    if (!oISBLBizLog.IsMercuryID(new Guid(Guid.NewGuid().ToString()), "", strLoginID, "LoginID")) //Mercury Enhancement  - Jan 2014 - Block notification if case creation user is Mercury ID
                        blnSendEmail = sendemail(strClientContactEmail, ConfigurationManager.AppSettings["SendFromStandardEmail"].ToString(), "Confirm Reset Password", strBody, true);
                    else
                        blnSendEmail = true;
                    
                    if (blnSendEmail)
                    {
                        //Write to Event Log

                        StringWriter swNewXMLData = new StringWriter();
                        XmlTextWriter xwNewXMLData = new XmlTextWriter(swNewXMLData);

                        //Write the root element
                        xwNewXMLData.WriteStartElement("Parent");

                        //Write sub-elements
                        xwNewXMLData.WriteElementString("LoginID", strLoginID);
                        //xwNewXMLData.WriteElementString("Password", strEncPassword);
                        // end the root element
                        xwNewXMLData.WriteEndElement();
                        strNewXMLData = swNewXMLData.ToString();
                        //Close the writer
                        xwNewXMLData.Close();


                        strAction = "Forgot Password";

                        DateTime dtActionStartDate = System.DateTime.Now;
                        DateTime dtActionEndDate = dtActionStartDate;
                        strActionBy = "Admin";
                        strUserType = "Admin";
                        strErrorCode = "FP_0 ";
                        strErrorMessage = "";
                        string strLogID = System.Guid.NewGuid().ToString();

                        myCmdWS.Dispose();
                        myCmd.Dispose();
                        blnSaveEvent = this.SaveEventLog(strLogID, "", "", strAction, "", strNewXMLData, dtActionStartDate, dtActionEndDate, strActionBy, strUserType, strErrorCode, strErrorMessage);
                        return (blnSaveEvent);
                    }
                    else
                    {
                        strAction = "Forgot Password";
                        DateTime dtActionStartDate = System.DateTime.Now;
                        DateTime dtActionEndDate = dtActionStartDate;
                        strActionBy = "Admin";
                        strUserType = "Admin";
                        strErrorCode = "FP_2 ";
                        strErrorMessage = "Email Service is Down.";
                        string strLogID = System.Guid.NewGuid().ToString();
                        blnSaveEvent = this.SaveEventLog(strLogID, "", "", strAction, "", "", dtActionStartDate, dtActionEndDate, strActionBy, strUserType, strErrorCode, strErrorMessage);
                        myCmd.Dispose();
                        return false;
                    }
                
                }
                catch (Exception)
                {
                    strAction = "Forgot Password";
                    DateTime dtActionStartDate = System.DateTime.Now; //System.DateTime.Parse(System.DateTime.Now.ToString("dd/MMM/yyyy"));
                    DateTime  dtActionEndDate = dtActionStartDate;
                    strActionBy = "Admin";
                    strUserType = "Admin";
                    strErrorCode = "FP_1 ";
                    strErrorMessage = "Forgot password request not instantiated.";
                    string strLogID = System.Guid.NewGuid().ToString();
                    blnSaveEvent = this.SaveEventLog(strLogID, "", "", strAction, "", "", dtActionStartDate, dtActionEndDate, strActionBy, strUserType, strErrorCode, strErrorMessage);
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }

        //Send email used by OCRS Web Service : Robust Email Enhancement 19 Jan 2009 Adam.
        public Boolean sendemail(String strTo, string strFrom, string strSubject, string strBody, bool IsBodyHTML)
        {

            Boolean blnSendemail = false;
            string strFileName = "";
            Boolean HasAttachment = false;          

            try
            {

                ISBL.General oISBLGen = new ISBL.General();
                blnSendemail = oISBLGen.AddEmailStorage(strTo, strFrom, strSubject, strBody, IsBodyHTML, strFileName, HasAttachment);
                oISBLGen.DisposeConnection();
                oISBLGen = null;
                return blnSendemail;

            }
            catch (Exception ex)
            {
                this.errTrack(ex);                
                return false;
            }
        }

        public string GetDueDate(DateTime dtStartDate, int intDueday, string strOfficeAssignment, Boolean blnAutoAssignment)
        {
            string strDueDate;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetDueDate";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@StartDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@Dueday", SqlDbType.Int);
            myCmd.Parameters.Add("@OfficeAssignment", SqlDbType.VarChar, 20);
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@StartDate"].Value = dtStartDate;
            myCmd.Parameters["@Dueday"].Value = intDueday;
            myCmd.Parameters["@OfficeAssignment"].Value = strOfficeAssignment;
            myCmd.Parameters["@AutoAssignment"].Value = blnAutoAssignment;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetDueDate";
            strDueDate =conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strDueDate;
        }

        public string GetResearchDueDate(DateTime dtEndDate, string strOfficeAssignment, Boolean blnAutoAssignment)
        {
            string strResearchDueDate;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchDueDate";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@EndDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@OfficeAssignment", SqlDbType.VarChar, 20);
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@EndDate"].Value = dtEndDate;
            myCmd.Parameters["@OfficeAssignment"].Value = strOfficeAssignment;
            myCmd.Parameters["@AutoAssignment"].Value = blnAutoAssignment;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchDueDate";
            strResearchDueDate = this.FormatDDMMYYYY(System.DateTime.Parse(conn.cmdScalarStoredProc(myCmd)));
            conn.Close();
            myCmd.Dispose();
            return strResearchDueDate;
        }

        public string GetResearchDueDateMMM(DateTime dtEndDate, string strOfficeAssignment, Boolean blnAutoAssignment)
        {
            string strResearchDueDate;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchDueDate";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@EndDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@OfficeAssignment", SqlDbType.VarChar, 20);
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@EndDate"].Value = dtEndDate;
            myCmd.Parameters["@OfficeAssignment"].Value = strOfficeAssignment;
            myCmd.Parameters["@AutoAssignment"].Value = blnAutoAssignment;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchDueDate";
            strResearchDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(conn.cmdScalarStoredProc(myCmd)));
            conn.Close();
            myCmd.Dispose();
            return strResearchDueDate;
        }

        public Boolean GetClientSingleCRNPerExcelFlag(string strClientCode)
        {
            Boolean  blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientSingleCRNPerExcelFlag";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientSingleCRNPerExcelFlag";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        public float GetBudget(float flBasePrice, float IncPriceCompany, float IncPriceIndividual, int intBaseCompany, int intBaseIndividual, int intActualCompany, int intActualIndividual)
            {
                int intCompanyDiff;
                int intIndividualDiff;
                float flCompanyBudget;
                float flIndividualBudget;
                float flActualBudget;

                if(intActualCompany == 0)
                {
                    intCompanyDiff = 0;
                }
                else
                {
                    if (intBaseCompany >= intActualCompany)
                    {
                        intCompanyDiff = 0;
                    }
                    else
                    {
                        intCompanyDiff = intActualCompany - intBaseCompany;
                    }
                }
               
                if (intActualIndividual == 0)
                {
                    intIndividualDiff = 0;
                }
                else
                {
                    if (intBaseIndividual >= intActualIndividual)
                    {
                        intIndividualDiff = 0;
                    }
                    else
                    {
                        intIndividualDiff = intActualIndividual - intBaseIndividual;
                    }
                }

                flCompanyBudget = intCompanyDiff * IncPriceCompany;
                flIndividualBudget = intIndividualDiff * IncPriceIndividual;

                flActualBudget = flBasePrice + flCompanyBudget + flIndividualBudget;
                return flActualBudget;
            }


        public float GetJPMCBudget(float flBasePrice, float IncPriceCompany, float IncPriceIndividual, int intBaseCompany, int intBaseIndividual, int intActualCompany, int intActualIndividual)
        {
            float flCompanyBudget = 0;
            float flIndividualBudget = 0;
            float flActualBudget = 0;

            int intCompanyCount = intActualCompany;
            int intIndividualCount = intActualIndividual;

            while (intCompanyCount >= intBaseCompany && intIndividualCount >= intBaseIndividual)
            {
                flActualBudget = flActualBudget + flBasePrice;
                intCompanyCount = intCompanyCount - intBaseCompany;
                intIndividualCount = intIndividualCount - intBaseIndividual;
            }

            flCompanyBudget = intCompanyCount * IncPriceCompany;
            flIndividualBudget = intIndividualCount * IncPriceIndividual;

            flActualBudget = flActualBudget + flCompanyBudget + flIndividualBudget;
            return flActualBudget;
        }

        public String GetCRNByOrderID(string strOrderID)
        {
            string strCRN;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCRNByOrderID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@OrderID", SqlDbType.VarChar, 80);
            myCmd.Parameters["@OrderID"].Value = strOrderID;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCRNByOrderID";
            strCRN = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strCRN;
        }

        public DataSet GetEventLogDataByCRN(string prmStrCRN)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetEventLogDataByCRN";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters["@CRN"].Value = prmStrCRN;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetEventLogDataByCRN";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        public Boolean SaveEventLog(string strLogID, string strClientCode, string strCRN, string strAction, string strOldXMLData, string strNewXMLData, DateTime dtActionStartDate, DateTime dtActionEndDate, string strActionBy, string strUserType, string strErrorCode, string strErrorMessage, string strImpersonateLoginID)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_SaveEventLogImpersonate";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LogID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters.Add("@Action", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@OldData", SqlDbType.NText);
            myCmd.Parameters.Add("@NewData", SqlDbType.NText);
            myCmd.Parameters.Add("@ActionStartDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@ActionEndDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@ActionBy", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 500);
            myCmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 50);
            myCmd.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 500);
            myCmd.Parameters.Add("@ImpersonateLoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LogID"].Value = new System.Guid(strLogID);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters["@Action"].Value = strAction;
            myCmd.Parameters["@OldData"].Value = strOldXMLData;
            myCmd.Parameters["@NewData"].Value = strNewXMLData;
            myCmd.Parameters["@ActionStartDate"].Value = dtActionStartDate;
            myCmd.Parameters["@ActionEndDate"].Value =  dtActionEndDate;
            myCmd.Parameters["@ActionBy"].Value = strActionBy;
            myCmd.Parameters["@UserType"].Value = strUserType;
            myCmd.Parameters["@ErrorCode"].Value = strErrorCode;
            myCmd.Parameters["@ErrorMessage"].Value = strErrorMessage;
            myCmd.Parameters["@ImpersonateLoginID"].Value = strImpersonateLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.SaveEventLog.sp_SaveEventLogImpersonate";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }

        }

        public Boolean SaveEventLog(string strLogID, string strClientCode, string strCRN, string strAction, string strOldXMLData, string strNewXMLData, DateTime dtActionStartDate, DateTime dtActionEndDate, string strActionBy, string strUserType, string strErrorCode, string strErrorMessage)
        {
            string strStatus;
            if (strErrorCode == null)
                strErrorCode = "";
            if (strErrorMessage == null)
                strErrorMessage = "";

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_SaveEventLog";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LogID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters.Add("@Action", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@OldData", SqlDbType.NText);
            myCmd.Parameters.Add("@NewData", SqlDbType.NText);
            myCmd.Parameters.Add("@ActionStartDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@ActionEndDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@ActionBy", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 500);
            myCmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 50);
            myCmd.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 500);
            myCmd.Parameters["@LogID"].Value = new System.Guid(strLogID);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters["@Action"].Value = strAction;
            myCmd.Parameters["@OldData"].Value = strOldXMLData;
            myCmd.Parameters["@NewData"].Value = strNewXMLData;
            myCmd.Parameters["@ActionStartDate"].Value = dtActionStartDate;
            myCmd.Parameters["@ActionEndDate"].Value = dtActionEndDate;
            myCmd.Parameters["@ActionBy"].Value = strActionBy;
            myCmd.Parameters["@UserType"].Value = strUserType;
            myCmd.Parameters["@ErrorCode"].Value = strErrorCode;
            myCmd.Parameters["@ErrorMessage"].Value = strErrorMessage;
            conn.Open();
            conn.callingMethod = "ISBL.General.SaveEventLog";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }

        }

        public Boolean UpdateEventLog(string strLogID, string strCRN, DateTime dtActionEndDate, string strErrorCode, string strErrorMessage)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateEventLog";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LogID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters.Add("@ActionEndDate", SqlDbType.DateTime);
            myCmd.Parameters.Add("@ErrorCode", SqlDbType.VarChar, 50);
            myCmd.Parameters.Add("@ErrorMessage", SqlDbType.VarChar, 500);
            myCmd.Parameters["@LogID"].Value = new System.Guid(strLogID);
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters["@ActionEndDate"].Value = dtActionEndDate;
            myCmd.Parameters["@ErrorCode"].Value = strErrorCode;
            myCmd.Parameters["@ErrorMessage"].Value = strErrorMessage;
            conn.Open();
            conn.callingMethod = "ISBL.General.UpdateEventLog";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
            
        }


        public Boolean IsLoginInstanceValid(string strLoginID, string strInstanceID)
        {
            string strStatus;
            Guid gString = new System.Guid(strInstanceID);

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsLoginInstanceValid";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@InstanceID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters["@InstanceID"].Value = gString;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsLoginInstanceValid";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
            
        }

        public Boolean IsLoginLegacy(string strLoginID)
        {
            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsLoginLegacy";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsLoginLegacy";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }

        }
        //Mitul - 13May2013 - Capture Client Browser and OS details and log them - START
        public DataSet ValidateClientLogin(string strLoginID, string strPassword, string strIPAddress, string strClientBrowser, string strClientBrowserVersion, string strUserAgent, String strJSEnabled, Boolean blnRemovePreviousInstance)
        {

            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_ValidateClientLoginBrowserInfo", conn.Connection);
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
            conn.callingMethod = "ISBL.General.ValidateClientLogin.sp_ValidateClientLogin";
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }
        //Mitul - 13May2013 - Capture Client Browser and OS details and log them - END

        public DataSet ValidateClientLogin(string strLoginID, string strPassword, string strIPAddress, Boolean blnRemovePreviousInstance)
        {

            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_ValidateClientLogin", conn.Connection);
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
            conn.callingMethod = "ISBL.General.ValidateClientLogin.sp_ValidateClientLogin";
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }

        //Used for Excel (Bulk Order) Login
        public DataSet ValidateClientLogin(string strLoginID, string strPassword)
        {
            DataSet myDataSet = new DataSet();
            SqlDataAdapter myCmd = new SqlDataAdapter("sp_ExcelValidateLogin", conn.Connection);
            myCmd.SelectCommand.CommandType = CommandType.StoredProcedure;
            myCmd.SelectCommand.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.SelectCommand.Parameters.Add("@Password", SqlDbType.VarChar, 100);
            myCmd.SelectCommand.Parameters["@LoginID"].Value = strLoginID.Trim();
            myCmd.SelectCommand.Parameters["@Password"].Value = strPassword.Trim();
            conn.Open();
            conn.callingMethod = "ISBL.General.ValidateClientLogin.sp_ExcelValidateLogin";
            myDataSet = conn.FillDataSet(myCmd);
            conn.Close();
            myCmd.Dispose();
            return myDataSet;
        }

        public Boolean Logout(string strLoginID, string strInstanceID)
        {
            string strStatus;
            try
            {
                Guid gString = new System.Guid(strInstanceID);

                SqlCommand myCmd = new SqlCommand();
                myCmd.Connection = conn.Connection;
                myCmd.CommandText = "sp_Logout";
                myCmd.CommandType = CommandType.StoredProcedure;
                myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                myCmd.Parameters.Add("@InstanceID", SqlDbType.UniqueIdentifier);
                myCmd.Parameters["@LoginID"].Value = strLoginID;
                myCmd.Parameters["@InstanceID"].Value = gString;
                conn.Open();
                conn.callingMethod = "ISBL.General.Logout";
                strStatus = conn.cmdScalarStoredProc(myCmd);
                conn.Close();
                if (strStatus.Length != 0)
                {
                    if (strStatus == "1")
                    {
                        myCmd.Dispose();
                        return true;
                    }
                    else
                    {
                        myCmd.Dispose();
                        return false;
                    }
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            catch
            {
                return false;
            }

        }

        public DataSet GetStatusList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetStatusList";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetStatusList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }

        public DataSet GetBDMClient(string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBDMClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetBDMClient";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }

        //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        public DataSet GetClientUserID(string strClientCode) 
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientUserID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientUserID";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }

        //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        public DataSet GetClientPersonateUserInfo(string strOldLoginID, string strOldLoginInstanceID, string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientPersonateUserInfo";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@OldLoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@OldLoginID"].Value = strOldLoginID;
            myCmd.Parameters.Add("@OldLoginInstanceID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@OldLoginInstanceID"].Value = new Guid(strOldLoginInstanceID);

            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientPersonateUserInfo";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }



        // Added By Stev, 11 Sept 2008
        // Purpose : To return a dataset by calling Stored Proc with parameters pass in
        // Not recommended to use - Adam 21 Jan 2009
        public DataSet GetDataByProc(string sProcName, string[] arParamName, string[] arParamValue, SqlDbType[] arDbTypes, int[] arSize)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = sProcName;
                sqlCmd.CommandType = CommandType.StoredProcedure;

                for (int i = 0; i < arParamName.Length; i++)
                {
                    sqlCmd.Parameters.Add(arParamName[i], arDbTypes[i], arSize[i]);
                    sqlCmd.Parameters[arParamName[i]].Value = arParamValue[i];
                }
                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetDataByProc";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Added By Stev, 11 Sept 2008
        // Purpose : To return a dataset by calling Stored Proc with parameters pass in
        // Not recommended to use - Adam 21 Jan 2009
        public DataSet GetDataByProc(string sProcName)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = sProcName;
                sqlCmd.CommandType = CommandType.StoredProcedure;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetDataByProc";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet GetClientBillingIndicator(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBillingIndicator";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientBillingIndicator";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }

       #endregion

        // Created By Stev, 22 Sept 2008, for CmBdmView.aspx
        #region CmBdmView method

        public DataSet GetClientByUserType(string sUserType, string sLoginId)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_GetClientByUserType";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 10);
                sqlCmd.Parameters["@UserType"].Value = sUserType;
                sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                sqlCmd.Parameters["@LoginID"].Value = sLoginId;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetClientByUserType";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public DataSet GetLoginIDByClCode(string sClCode)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_GetLoginIDByClCode";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@CCode", SqlDbType.VarChar, 30);
                sqlCmd.Parameters["@CCode"].Value = sClCode;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetLoginIDByClCode";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Created to bind CmBdmView data grid.
        public DataSet GetCmBdmView(string sClCode, string sLoginId, string sStCode, string sCrn, string sStDate, string sEdDate, string sShowAll, string sUserLoginID, string sUserType, string sOrderID, string sBulkOrderMasterID)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_GetCmBdmView";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@ClCode", SqlDbType.VarChar, 30);
                sqlCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                sqlCmd.Parameters.Add("@StatusCode", SqlDbType.VarChar, 10);
                sqlCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
                //sqlCmd.Parameters.Add("@Subject", SqlDbType.NVarChar, 250);
                sqlCmd.Parameters.Add("@OrdStDate", SqlDbType.VarChar, 20);
                sqlCmd.Parameters.Add("@OrdEdDate", SqlDbType.VarChar, 20);
                sqlCmd.Parameters.Add("@ShowAll", SqlDbType.VarChar, 1);
                sqlCmd.Parameters.Add("@UserLoginID", SqlDbType.VarChar, 15);
                sqlCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 10);
                sqlCmd.Parameters.Add("@OrderID", SqlDbType.VarChar, 50); //OCRS Small Enhancement BRS Version 1.9.doc Section 6.D
                sqlCmd.Parameters.Add("@BulkOrderMasterID", SqlDbType.VarChar, 50);//OCRS Small Enhancement BRS Version 1.9.doc Section 6.D

                sqlCmd.Parameters["@ClCode"].Value = sClCode;
                sqlCmd.Parameters["@LoginId"].Value = sLoginId;
                sqlCmd.Parameters["@StatusCode"].Value = sStCode;
                sqlCmd.Parameters["@CRN"].Value = sCrn;
                //sqlCmd.Parameters["@Subject"].Value = sSubj;
                sqlCmd.Parameters["@OrdStDate"].Value = sStDate;
                sqlCmd.Parameters["@OrdEdDate"].Value = sEdDate;
                sqlCmd.Parameters["@ShowAll"].Value = sShowAll;
                sqlCmd.Parameters["@UserLoginID"].Value = sUserLoginID;
                sqlCmd.Parameters["@UserType"].Value = sUserType;
                sqlCmd.Parameters["@OrderID"].Value = sOrderID; //OCRS Small Enhancement BRS Version 1.9.doc Section 6.D
                sqlCmd.Parameters["@BulkOrderMasterID"].Value = sBulkOrderMasterID; //OCRS Small Enhancement BRS Version 1.9.doc Section 6.D

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetCmBdmView";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        // Created By Stev, 17 Oct 2008, for CreatedNewUser.aspx
        #region C4Client funtion
        // Return DataSet of ClientDetails which is finished set up
        public DataSet GetClientByUserType_Settled(string sUserType, string sLoginId)
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandText = "sp_GetClientByUserType_Settled";
                sqlCmd.CommandType = CommandType.StoredProcedure;

                sqlCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 10);
                sqlCmd.Parameters["@UserType"].Value = sUserType;
                sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                sqlCmd.Parameters["@LoginID"].Value = sLoginId;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetClientByUserType_Settled";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Return DataSet of C4 Client List that has been selected.
        public DataSet GetC4ClientSelected(string sLoginId)
        {
            DataSet dsClientSlt = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "Sp_GetC4ClientSelected";

                sqlCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                sqlCmd.Parameters["@LoginId"].Value = sLoginId;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetC4ClientSelected";
                dsClientSlt = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsClientSlt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Return DataSet of C5 Client List that has been selected.
        public DataSet GetC5ClientSelected(string sLoginId)
        {
            DataSet dsClientSlt = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "Sp_GetC5ClientSelected";

                sqlCmd.Parameters.Add("@LoginId", SqlDbType.VarChar, 15);
                sqlCmd.Parameters["@LoginId"].Value = sLoginId;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetC5ClientSelected";
                dsClientSlt = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsClientSlt;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion

        // Created By Stev, 26 Sept 08, for ClientMaintenance.aspx
        #region ClientMaintenance method

        // Return Currency detail to ClientMaintenance, for binding ddl
        public DataSet GetCurrency()
        {
            DataSet dsTemp = new DataSet();
            SqlCommand sqlCmd = new SqlCommand();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "sp_GetCurrency";

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetCurrency";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Return BPRegType detail to ClientMaintenance, for binding ddl
        public DataSet GetBPRegType()
        {
            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsTemp = new DataSet();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "Sp_GetBPRegType";

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetBPRegType";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlCmd.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Return BPRegTier detail to ClientMaintenance, for binding ddl
        public DataSet GetBPRegTier()
        {
            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsTemp = new DataSet();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "sp_GetBPRegTier";

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetBPRegTier";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlDa.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        // Return Client Master details to ClientMaintenance, for binding grid view
        public DataSet GetClientDetails(string sUserType, string sLoginId, string sClientCode)
        {
            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsTemp = new DataSet();

            try
            {
                sqlCmd.Connection = conn.Connection;
                sqlCmd.CommandType = CommandType.StoredProcedure;
                sqlCmd.CommandText = "sp_GetClientDetails";

                sqlCmd.Parameters.Add("@UserType", SqlDbType.VarChar, 10);
                sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                sqlCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                sqlCmd.Parameters["@UserType"].Value = sUserType;
                sqlCmd.Parameters["@LoginID"].Value = sLoginId;
                sqlCmd.Parameters["@ClientCode"].Value = sClientCode;

                SqlDataAdapter sqlDa = new SqlDataAdapter();
                sqlDa.SelectCommand = sqlCmd;
                conn.Open();
                conn.callingMethod = "ISBL.General.GetClientDetails";
                dsTemp = conn.FillDataSet(sqlDa);
                conn.Close();
                sqlDa.Dispose();
                return dsTemp;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        #endregion 



        //public  byte[] ConvertImageToByteArray(System.Drawing.Image imageToConvert, ImageFormat formatOfImage)
        //{
        //    byte[] Ret;

        //    try
        //    {
        //        using (MemoryStream ms = new MemoryStream())
        //        {
        //            imageToConvert.Save(ms, formatOfImage);
        //            Ret = ms.ToArray();
        //        }
        //    }
        //    catch (Exception) { throw; }

        //    return Ret;
        //} 

        public string FormatMMDDYYYY(DateTime dtDate) //eg. mm/dd/yyyy
        {
            return AddFrontZero(dtDate.Month.ToString()) + "/" + AddFrontZero(dtDate.Day.ToString()) + "/" + dtDate.Year.ToString();
        }

        public string FormatStringDDMMMYYYY(string strDDMMYYYY) //eg. dd-MMM-yyyy
        {


            string strReturn;        
            string[]  strArrParameter;
            char[] splitter = { '/' };

            strArrParameter = strDDMMYYYY.Split(splitter);

            strReturn = strArrParameter[0].ToString() + "-" + ShortDDMonth(strArrParameter[1].ToString()) + "-" + strArrParameter[2].ToString();
            return strReturn;
        }

        public string FormatDDMMMYYYY(DateTime dtDate)
        {
            return AddFrontZero(dtDate.Day.ToString()) + "-" + ShortMonth(dtDate.Month.ToString()) + "-" + dtDate.Year.ToString();
        }

        public string FormatDDMMYYYY(DateTime dtDate)
        {
            return AddFrontZero(dtDate.Day.ToString()) + "/" + AddFrontZero(dtDate.Month.ToString()) + "/" + dtDate.Year.ToString();
        }

        public string ShortMonth(string strMonth)
        {
            switch (strMonth)
            {
                case "1":
                    strMonth = "Jan";
                    break;

                case "2":
                    strMonth = "Feb";
                    break;

                case "3":
                    strMonth = "Mar";
                    break;

                case "4":
                    strMonth = "Apr";
                    break;

                case "5":
                    strMonth = "May";
                    break;

                case "6":
                    strMonth = "Jun";
                    break;

                case "7":
                    strMonth = "July";
                    break;

                case "8":
                    strMonth = "Aug";
                    break;

                case "9":
                    strMonth = "Sep";
                    break;

                case "10":
                    strMonth = "Oct";
                    break;

                case "11":
                    strMonth = "Nov";
                    break;

                case "12":
                    strMonth = "Dec";
                    break;

            }
            return strMonth;
        }

        public string ShortDDMonth(string strDoubleDigitMonth)
        {
            switch (strDoubleDigitMonth)
            {
                case "01":
                    strDoubleDigitMonth = "Jan";
                    break;

                case "02":
                    strDoubleDigitMonth = "Feb";
                    break;

                case "03":
                    strDoubleDigitMonth = "Mar";
                    break;

                case "04":
                    strDoubleDigitMonth = "Apr";
                    break;

                case "05":
                    strDoubleDigitMonth = "May";
                    break;

                case "06":
                    strDoubleDigitMonth = "Jun";
                    break;

                case "07":
                    strDoubleDigitMonth = "July";
                    break;

                case "08":
                    strDoubleDigitMonth = "Aug";
                    break;

                case "09":
                    strDoubleDigitMonth = "Sep";
                    break;

                case "10":
                    strDoubleDigitMonth = "Oct";
                    break;

                case "11":
                    strDoubleDigitMonth = "Nov";
                    break;

                case "12":
                    strDoubleDigitMonth = "Dec";
                    break;

            }
            return strDoubleDigitMonth;
        }

        public string AddFrontZero(string strNumber)
        {
            if (strNumber.Length == 1)
            {
                strNumber = "0" + strNumber;
            }
            return strNumber;
        }

        public string SafeSqlLiteral(string inputSQL)
        {
            return inputSQL.Replace("'", "''").Replace("[", "[[]");
        }

        /* Start Personalising ISIS - Adam 7 Aug 2008 */

        public DataSet GetClientAccessModule(string strSettingCode, string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientAccessModule";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@SettingCode", SqlDbType.VarChar, 20);
            myCmd.Parameters["@SettingCode"].Value = strSettingCode;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientAccessModule";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }


        public DataSet GetClientSettings(string strSettingCode, string strLoginID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientSettings";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@SettingCode", SqlDbType.VarChar, 20);
            myCmd.Parameters["@SettingCode"].Value = strSettingCode;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientSettings";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            return ds;
        }

        public Boolean  UpdatetClientSettings(string strSettingCode, string strSettingDescription, string strLoginID, string strValue)
        {

            string strStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdatetClientSettings";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@SettingCode", SqlDbType.VarChar, 20);
            myCmd.Parameters["@SettingCode"].Value = strSettingCode;
            myCmd.Parameters.Add("@SettingDescription", SqlDbType.VarChar, 100);
            myCmd.Parameters["@SettingDescription"].Value = strSettingDescription;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters.Add("@ValueCode", SqlDbType.VarChar, 20);
            myCmd.Parameters["@ValueCode"].Value = strValue;
            conn.Open();
            conn.callingMethod = "ISBL.General.UpdatetClientSettings";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }


        /* End Personalising ISIS - Adam 7 Aug 2008 */

        /* Added By Avinash  */

        public DataSet GetCountryList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryListAll";
            myCmd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCountryList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetCountryList(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryList";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCountryList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetSubReportTypeList(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetSubReportTypeList";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetSubReportTypeList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetReportType(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetReportType";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Start Adam Mar 2010 - Get list of Report Type for the existing placed orders
        public DataSet GetTrackOrderReportType(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetTrackOrderReportType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetTrackOrderReportType";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //added by snajeeva for JPMC Changes start
        public DataSet GetReportFlag(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportFlag";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetReportFlag";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetReportFormatFlag(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetReportFormatFlag";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetReportFormatFlag";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetTrackOrderSubReportType(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetTrackOrderSubReportType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetTrackOrderSubReportType";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        // end of jpmc changes
        public DataSet GetBulkCaseNoList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkCaseNoList";
            myCmd.CommandType = CommandType.StoredProcedure;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetBulkCaseNoList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //End Adam Mar 2010 - Get list of Report Type for the existing placed orders

        public DataSet GetResearchElement(string strClientCode, string strReportType, string strSubType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchElement";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@SubType"].Value = strSubType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchElement";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        /*Start BI 20 ISIS Refresh Order */
        public DataSet GetResearchElement(string strClientCode, string strReportType, string strSubType, string strSubjectID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRefreshOrderResearchElement";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubjectID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@SubType"].Value = strSubType;
            myCmd.Parameters["@SubjectID"].Value = new Guid(strSubjectID);
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchElement.sp_GetRefreshOrderResearchElement";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /*End BI 20 ISIS Refresh Order */

        public Boolean IsShowBudget(string strClientCode)
        {
            String strBudget;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsShowBudget"; ;
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsShowBudget";
            strBudget = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strBudget.Length != 0)
            {
                strBudget = strBudget.ToLower();

                if (strBudget.Equals("true"))
                {
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
        public Boolean IsShowDueDate(string strClientCode)
        {
            String strDueDate;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsShowDueDate"; ;
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsShowDueDate";
            strDueDate = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strDueDate.Length != 0)
            {
                strDueDate = strDueDate.ToLower();

                if (strDueDate.Equals("true"))
                {
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
        public DataSet GetClientCurrency(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientCurreny";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientCurrency";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetBudgetParamsBasedOnVariantCountry(string ClientCode, string strReportType, string Variant, string Country, bool IsExpress)
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
            cmd.Parameters.Add("@Country", SqlDbType.VarChar, 4); //Jul 2013 Adam - ISIS Atlas Data Sync
            cmd.Parameters.Add("@IsExpress", SqlDbType.Bit);
            cmd.Parameters["@ClientCode"].Value = ClientCode;
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


        public DataSet GetClientBudgetDetail(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBudgetDetail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientBudgetDetail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

         //Start Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement
        public DataSet GetClientBudgetDetailExpress(string strClientCode, string strReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBudgetDetailExpress";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientBudgetDetailExpress";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //CRE = Client / Report Type / Express Case
        public String GetClientDueDayByCRE(string strReportType, string strClientCode, Boolean blnExpress)
        {
            string strDuedate;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDueDayByCRE";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@Express", SqlDbType.Bit);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@Express"].Value = blnExpress;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientDueDayByCRE";
            strDuedate = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strDuedate;
        }
         //End Adam Nov 09 - OCRS Phase 4 2.8b Change Budget and TAT Express Case Enhancement

        public String GetClientEmail(string strClientCode)
        {
            string ClientEmail;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientEmail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientEmail";
            ClientEmail = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return ClientEmail;
        }
        public String GetResearchElementCode(string RElements)
        {
            string RECode;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchElementCode";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@RElements", SqlDbType.VarChar, 500);
            myCmd.Parameters["@RElements"].Value = RElements;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchElementCode";
            RECode = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return RECode;
        }

        public String GetClientDueDay(string strReportType)
        {
            string strDuedate;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDueDay";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientDueDay";
            strDuedate = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strDuedate;
        }

        public String GetClientDueDayBasedOnVariantCountry(string ClientCode, string strReportType, string Variant, string Country, bool IsExpress)
        {
            string strDuedate;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientVariantSubjectCountryTAT";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
            //myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 500);Jul 2013 Adam - ISIS Atlas Data Sync
            myCmd.Parameters.Add("@Country", SqlDbType.VarChar, 4); //Jul 2013 Adam - ISIS Atlas Data Sync
            myCmd.Parameters.Add("@IsExpress", SqlDbType.Bit);
            myCmd.Parameters["@ClientCode"].Value = ClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@Variant"].Value = Variant;
            myCmd.Parameters["@Country"].Value = Country;
            myCmd.Parameters["@IsExpress"].Value = IsExpress;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientDueDayBasedOnVariantCountry";
            strDuedate = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strDuedate;
        }

        public String GetCountry(string strCountry)
        {
            string strCCode;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CountryName", SqlDbType.VarChar, 100);
            myCmd.Parameters["@CountryName"].Value = strCountry;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCountry";
            strCCode = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strCCode;
        }
        public String GetSubReportTypeCode(string strSubreporttype,string strreporttype)
        {
            string strCCode;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetSubReportTypeCode";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 500);
            myCmd.Parameters["@SubReportType"].Value = strSubreporttype;
            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = strreporttype;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetSubReportTypeCode";
            strCCode = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strCCode;
        }
        public DataSet GetClientBranchDetail(string strClientCode, String strReportType, Boolean blAssignment, Boolean blBulk)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBranchDetail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters.Add("@AutoAssignment", SqlDbType.Bit);
            myCmd.Parameters["@AutoAssignment"].Value = blAssignment;
            myCmd.Parameters.Add("@BulkOrder", SqlDbType.Bit);
            myCmd.Parameters["@BulkOrder"].Value = blBulk;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientBranchDetail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public DataSet GetClientDetail(string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientDetail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientDetail";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        public String GetClientCaseManager(string strClientCode, String strReportType)
        {
            string strCasemanager;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientCaseManager";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientCaseManager";
            strCasemanager = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strCasemanager;
        }

        /* End Add Avinash*/

        /* Start Stephanie */
        public bool ISISRegistrationMail(string strCompanyName, string strAddress, string strCountry, string strContactName, string strContactEmail, string strContactNumber)
        {
            string strEmailSubject;
            string strEmailBody;
            string strToEmail;
            string strFromEmail;

            strEmailSubject = "Request for EDD Online Registration, from " + strCompanyName + "( " + strContactName + " )";
            strEmailBody = "<table><tr><td>Please follow up this request</td></tr>";
            strEmailBody += "<tr><td>Company Name </td><td>:</td><tr>" + strCompanyName + "</td></tr>";
            strEmailBody += "<tr><td>Address </td><td>:</td><tr>" + strAddress + "</td></tr>";
            strEmailBody += "<tr><td>Country </td><td>:</td><tr>" + strCountry + "</td></tr>";
            strEmailBody += "<tr><td>Contact (Name) </td><td>:</td><tr>" + strContactName + "</td></tr>";
            strEmailBody += "<tr><td>Contact (Email) </td><td>:</td><tr>" + strContactEmail + "</td></tr>";
            strEmailBody += "<tr><td>Contact (Number) </td><td>:</td><tr>" + strContactNumber + "</td></tr>";
            strEmailBody += "<tr><td height='50'>Regards,</td></tr>";
            //Start Change Email Signed - Feb 2013 Adam
            //strEmailBody += "<tr><td>Online Order System";
            strEmailBody += "<tr><td>Thomson Reuters Risk";
            strEmailBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></td></tr>";
            strEmailBody += "<tr><td height='20'></td></tr>";
            //End Change Email Signed - Feb 2013 Adam
            strEmailBody += "</table>";

            strToEmail = System.Configuration.ConfigurationManager.AppSettings["ISISRegistrationEmail"];
            strFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
            return sendemail(strToEmail, strFromEmail, strEmailSubject, strEmailBody, true);

        }
        public bool ContactUsSendMail(string strName, string strEmail, string strTitle, string strComment)
        {
            string strEmailSubject;
            string strEmailBody;
            string strContactUsToEmail;
            string strContactUsFromEmail;

            strEmailSubject = "Message from " + strName + " :" + strTitle;
            strEmailBody = "<table><tr><td>Please follow up this request</td></tr>";
            strEmailBody += "<tr><td>Name : " + strName + "</td></tr>";
            strEmailBody += "<tr><td>Email Address : " + strEmail + "</td></tr>";
            strEmailBody += "<tr><td>Comments :</td></tr>";
            strEmailBody += "<tr><td>" + strComment + "</td></tr>";
            strEmailBody += "<tr><td height='50'>Regards,</td></tr>";
            //Start Change Email Signed - Feb 2013 Adam
            //strEmailBody += "<tr><td>Online Order System";
            strEmailBody += "<tr><td>Thomson Reuters Risk";
            strEmailBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></td></tr>";
            strEmailBody += "<tr><td height='20'></td></tr>";
            //End Change Email Signed - Feb 2013 Adam
            strEmailBody += "</table>";

            strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactUsEmail"];
            strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
            return sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);

        }
        public bool ContactUsSendMail(string strFirstName, string strLastName, string strCompany, string strAddress, string strZipCode, string strCountry, string strTelephone, string strJobTitle, string strEmail, string strMessage)
        {
            string strEmailSubject;
            string strEmailBody;
            string strContactUsToEmail;
            string strContactUsFromEmail;

            strEmailSubject = "Message from " + strFirstName + " " + strLastName + " : " + strJobTitle;
            strEmailBody = "<table><tr><td colspan='3'>Please follow up this request</td></tr>";
            strEmailBody += "<tr><td width='25%'>Name </td><td>:</td><td width='70%'> " + strFirstName + " " + strLastName + "</td></tr>";
            strEmailBody += "<tr><td>Company </td><td>:</td><td> " + strCompany + "</td></tr>";
            strEmailBody += "<tr><td>Company Addres </td><td>:</td><td> " + strAddress + "</td></tr>";
            strEmailBody += "<tr><td>Zip Code </td><td>:</td><td> " + strZipCode + "</td></tr>";
            strEmailBody += "<tr><td>Country </td><td>:</td><td>" + strCountry + "</td></tr>";
            strEmailBody += "<tr><td>Telephone  </td><td>:</td><td> " + strTelephone + "</td></tr>";
            strEmailBody += "<tr><td>Job Title  </td><td>:</td><td>" + strJobTitle + "</td></tr>";
            strEmailBody += "<tr><td>Email </td><td>:</td><td>" + strEmail + "</td></tr>";
            strEmailBody += "<tr><td valign='Top'>Message </td><td valign='Top'>:</td><td> " + strMessage + "</td></tr>";
            strEmailBody += "<tr><td height='50'>Regards,</td></tr>";
            //Start Change Email Signed - Feb 2013 Adam
            //strEmailBody += "<tr><td>Online Order System";
            strEmailBody += "<tr><td>Thomson Reuters Risk";
            strEmailBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></td></tr>";
            strEmailBody += "<tr><td height='20'></td></tr>";
            //End Change Email Signed - Feb 2013 Adam
            strEmailBody += "</table>";

            strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactUsEmail"];
            strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
            return sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);

        }
        public bool ContactUsSendMail(string strFirstName, string strLastName, string strCompany, string strTelephone, string strEmail, bool bTechnical, bool bProduct, bool bGeneral, string strMessage)
        {
            string strEmailSubject;
            string strEmailBody;
            string strContactUsToEmail;
            string strContactUsFromEmail;
            bool bTechnicalMailSent =true ;
            bool bGeneralMailSent = true;

            strEmailSubject = "Message from " + strFirstName + " " + strLastName;
            strEmailBody = "<table><tr><td colspan='3'>Please follow up this request</td></tr>";
            strEmailBody += "<tr><td width='25%'>Name </td><td>:</td><td width='70%'> " + strFirstName + " " + strLastName + "</td></tr>";
            strEmailBody += "<tr><td>Company </td><td>:</td><td> " + strCompany + "</td></tr>";
            strEmailBody += "<tr><td>Telephone  </td><td>:</td><td> " + strTelephone + "</td></tr>";
            strEmailBody += "<tr><td>Email </td><td>:</td><td>" + strEmail + "</td></tr>";
            if (bTechnical == true || bProduct == true || bGeneral == true)
            {
                strEmailBody += "<br><tr><td>I would like to :- </td></tr>";
                if (bTechnical) { strEmailBody += "<tr><td colspan='3'>* enquire about technical issue.</td></tr>"; }
                if (bProduct) { strEmailBody += "<tr><td colspan='3'>* know more information about your products.</td></tr>"; }
                if (bGeneral) { strEmailBody += "<tr><td colspan='3'>* make a general enquiry.</td></tr>"; }
                strEmailBody += "<br>";
            }
            strEmailBody += "<tr><td valign='Top'>Message </td><td valign='Top'>:</td><td> " + strMessage + "</td></tr>";
            strEmailBody += "<tr><td height='50'>Regards,</td></tr>";
            //Start Change Email Signed - Feb 2013 Adam
            //strEmailBody += "<tr><td>Online Order System";
            strEmailBody += "<tr><td>Thomson Reuters Risk";
            strEmailBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></td></tr>";
            strEmailBody += "<tr><td height='20'></td></tr>";
            //End Change Email Signed - Feb 2013 Adam
            strEmailBody += "</table>";

           
            if (bTechnical)
            {
                strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactUsEmailInternalTechnical"];
                strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
                bTechnicalMailSent = sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);
            }

            if (bProduct || bGeneral)
            { 
                strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactUsEmailInternalGeneral"];
                strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
                bGeneralMailSent = sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);
            }
            else if (bTechnical == false & bProduct == false & bGeneral == false)
            {
                strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactUsEmailInternalGeneral"];
                strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
                bGeneralMailSent = sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);
            }


            if (bTechnicalMailSent == false || bGeneralMailSent == false)
            {
                return false;
            }
            else 
            {
                return true;
            }
     

        }

        public bool ContactTechSupportSendMail(string strName, string strLoginID, string strCompany, string strEmail, string strTelephone, string strCountry, string strSubject, string strDescription)
        {
            string strEmailSubject;
            string strEmailBody;
            string strContactUsToEmail;
            string strContactUsFromEmail;

            strEmailSubject = "Message from " + strName + " : " + strSubject;
            strEmailBody = "<table><tr><td>Please follow up this request</td></tr>";
            strEmailBody += "<tr><td>Name : " + strName + "</td></tr>";
            strEmailBody += "<tr><td>Login ID : " + strLoginID + "</td></tr>";
            strEmailBody += "<tr><td>Company: " + strCompany  + "</td></tr>";
            strEmailBody += "<tr><td>Email Address : " + strEmail + "</td></tr>";
            strEmailBody += "<tr><td>Telephone : " + strTelephone + "</td></tr>";
            strEmailBody += "<tr><td>Country : " + strCountry + "</td></tr>";
            strEmailBody += "<tr><td>Description :</td></tr>";
            strEmailBody += "<tr><td>" + strDescription + "</td></tr>";
            strEmailBody += "<tr><td height='50'>Regards,</td></tr>";
            //Start Change Email Signed - Feb 2013 Adam
            //strEmailBody += "<tr><td>Online Order System";
            strEmailBody += "<tr><td>Thomson Reuters Risk";
            strEmailBody += "<br /><font size='3'><a href='http://risk.thomsonreuters.com/'>risk.thomsonreuters.com</a></font></td></tr>";
            strEmailBody += "<tr><td height='20'></td></tr>";
            //End Change Email Signed - Feb 2013 Adam
            strEmailBody += "</table>";

            strContactUsToEmail = System.Configuration.ConfigurationManager.AppSettings["ContactTechSupportEmail"];
            strContactUsFromEmail = System.Configuration.ConfigurationManager.AppSettings["SendFromStandardEmail"];
            return sendemail(strContactUsToEmail, strContactUsFromEmail, strEmailSubject, strEmailBody, true);

        }
        public string GetPassword(string strLoginID)
        {
            string strPassword;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetPassword";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@UserID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@UserID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetPassword";
            strPassword = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strPassword;
        }

        public DataSet GetFinalReport(string strOrderID, string strCRN)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetFinalReport";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters["@CRN"].Value = strCRN;
            if (strOrderID != null)
            {
                myCmd.Parameters["@OrderID"].Value = new Guid(strOrderID);

            }
            else
            {
                myCmd.Parameters["@OrderID"].Value = null;
            }
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetFinalReport";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean CheckCRNvsLoginID(string strCRN, string strLoginID)
        {
            string strCheck = string.Empty;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckCRNvsLoginID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 50);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters["@CRN"].Value = strCRN;
            conn.Open();
            conn.callingMethod = "ISBL.General.CheckCRNvsLoginID";
            strCheck = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strCheck == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        // added by sanjeeva for Bulkupload Summary start on 19/10/2015
        public Boolean CheckBulkLoginID(string strBulkID, string strLoginID)
        {
            string strCheck = string.Empty;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckBulkidvsLoginID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@BID", SqlDbType.VarChar, 100);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters["@BID"].Value = strBulkID;
            conn.Open();
            conn.callingMethod = "ISBL.General.CheckBulkLoginID";
            strCheck = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strCheck == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //end of bulkorder summary

        /** Start Adam - Nov 2009 - OCRS Phase 4 2.8b - Master setting for English and Non-English speaking Country **/
        public Boolean IsEnglishSpeakingCountry(string strCountryDesc)
        {
            string strCheck = string.Empty;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsEnglishSpeakingCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CountryDesc", SqlDbType.VarChar, 100);
            myCmd.Parameters["@CountryDesc"].Value = strCountryDesc;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsEnglishSpeakingCountry";
            strCheck = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strCheck == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        /** End Adam - Nov 2009 - OCRS Phase 4 2.8b - Master setting for English and Non-English speaking Country **/

        /* End Stephanie */

        /*@@@ Start Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement. */
        public DataSet GetInfoByBulkMasterID(string strClientCode, string strBulkMasterID)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetInfoByBulkMasterID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@BulkMasterID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@BulkMasterID"].Value = new System.Guid(strBulkMasterID);
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetInfoByBulkMasterID";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /*@@@ End Adam Nov 09 - OCRS Phase 4 2.8b Adding Subject Name in email subject title Enhancement. */


        /*@@@ Start by Adam 20090310 ISIS3 - Report Variant */
        public DataSet GetResearchElementByVariant(string strClientCode, string strReportType, string strSubType, string strVariant)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchElementByVariant";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@SubType"].Value = strSubType;
            myCmd.Parameters["@Variant"].Value = strVariant;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchElementByVariant";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetResearchElementBySubReportType(string strClientCode, string strReportType, string strSubType, string strSubreportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetResearchElementBySubReportType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@SubType"].Value = strSubType;
            myCmd.Parameters["@SubReportType"].Value = strSubreportType;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetResearchElementBySubReportType";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /*@@@ End by Adam 20090310 ISIS3 - Report Variant */


        /*@@@ Start by Adam 20090505 - Password Hardening */
        public String IsPasswordHardened(string Password, string strLoginID)
        {
            String sRule = "";
            Boolean bAdd = false;
            const string lower = "abcdefghijklmnopqrstuvwxyz";
            const string upper = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";

            if (!(Password.Length >= 7)) //All passwords should be atleast 7 characters long
            {
                sRule += "1";
                bAdd = true;
            }


            if (!(Password.IndexOfAny(lower.ToCharArray()) >= 0)) //Check Lowercase if rule is enforced
            {
                if (bAdd)
                {
                    sRule += ", 2";
                }
                else
                {
                    sRule += "2";
                }
                bAdd = true;
            }
            else
            {
                if (!(Password.IndexOfAny(upper.ToCharArray()) >= 0)) //Check Uppercase if rule is enforced
                {
                    if (bAdd)
                    {
                        sRule += ", 2";
                    }
                    else
                    {
                        sRule += "2";
                    }
                    bAdd = true;
                }
            }


             //if (strLoginID.ToLower() == Password.ToLower()) //The password should not contain the userID.
            string lPassword = Password.ToLower();
            if (lPassword.Contains(strLoginID.ToLower()))
             {
                 if (bAdd)
                 {
                     sRule += ", 3";
                 }
                 else
                 {
                     sRule += "3";
                 }
                 bAdd = true;
             }

             if (CheckPasswordDictionary(Password)) //The password should not be based on a dictionary word.
              {
                  if (bAdd)
                  {
                      sRule += ", 4";
                  }
                  else
                  {
                      sRule += "4";
                  }
                  bAdd = true;
              }

              return sRule;
        }

        public Boolean CheckPasswordDictionary(string strText)
        {
            string strCheck = string.Empty;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_CheckPasswordDictionary";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Password", SqlDbType.VarChar, 20);
            myCmd.Parameters["@Password"].Value = strText;
            conn.Open();
            conn.callingMethod = "ISBL.General.CheckDictionary";
            strCheck = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (strCheck == "1")
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        /*@@@ End by Adam 20090505 - Password Hardening */

        /**********Mudassar Khan--Code---Start**************************************/
  #region Mudassar Khan Code

        public Boolean ISEnableAutoAssignment(string ClientCode)
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

        public String GetCaseManagerEmail(String strClientCode, String strReportType)
        {
            String strCaseManagerEmail;
            SqlCommand cmd = new SqlCommand();

            ISBL.General oISBLGen = new ISBL.General();
            String strCaseManager = oISBLGen.GetClientCaseManager(strClientCode, strReportType);
            oISBLGen = null;

            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetCaseManagerEmail";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;
            cmd.Parameters.Add("@CaseManager", SqlDbType.VarChar, 50);
            cmd.Parameters["@CaseManager"].Value = strCaseManager;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCaseManagerEmail";
            strCaseManagerEmail = conn.cmdScalarStoredProc(cmd);
            conn.Close();
            cmd.Dispose();
            return strCaseManagerEmail;
        }

        public String GetBranchEmail(String strClientCode, String strReportType, Boolean blnEnableAutoAssignment)
        {
            String strBranchEmail, strBranch;
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = conn.Connection;
            cmd.CommandText = "sp_GetBranchEmail";
            cmd.CommandType = CommandType.StoredProcedure;
            cmd.Transaction = conn.currentTransaction;

            ISBL.General oISBLGen = new ISBL.General();
            DataSet ds = new DataSet();
            conn.Open();
            conn.callingMethod = "ISBL.General.GetBranchEmail";
            ds = oISBLGen.GetClientBranchDetail(strClientCode, strReportType, blnEnableAutoAssignment, true);
            strBranch = ds.Tables[0].Rows[0]["Branch"].ToString();
            ds.Dispose();
            oISBLGen = null;

            cmd.Parameters.Add("@Branch", SqlDbType.VarChar, 10);
            cmd.Parameters["@Branch"].Value = strBranch;
            strBranchEmail = conn.cmdScalarStoredProc(cmd);
            conn.Close();
            cmd.Dispose();
            return strBranchEmail;
        }

        //for Robust Eamil enhancement 19 Jan 2009 Adam
        public Boolean AddEmailStorage(string strMailTo, string strMailFrom, string strSubject, string strBody, Boolean bHTML, string strFileName, Boolean bHasAttahcment)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_AddEmailStorage";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@MailTo", SqlDbType.VarChar, 1000);
            myCmd.Parameters.Add("@MailFrom", SqlDbType.VarChar, 250);
            myCmd.Parameters.Add("@Subject", SqlDbType.NVarChar, 500);
            myCmd.Parameters.Add("@Body", SqlDbType.NText);
            myCmd.Parameters.Add("@HTML", SqlDbType.Bit);
            myCmd.Parameters.Add("@FileName", SqlDbType.VarChar, 1000);
            myCmd.Parameters.Add("@HasAttahcment", SqlDbType.Bit);

            myCmd.Parameters["@MailTo"].Value = strMailTo;
            myCmd.Parameters["@MailFrom"].Value = strMailFrom;
            myCmd.Parameters["@Subject"].Value = strSubject;
            myCmd.Parameters["@Body"].Value = strBody;
            myCmd.Parameters["@HTML"].Value = bHTML;
            myCmd.Parameters["@FileName"].Value = strFileName;
            myCmd.Parameters["@HasAttahcment"].Value = bHasAttahcment;
            conn.Open();
            conn.callingMethod = "ISBL.General.AddEmailStorage";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        //This is used for bulk order to send email with attachment: Robust Email Enhancement 19 Jan 2009 Adam.
        public static bool sendemail(String strTo, string strFrom, string strSubject, string strBody, bool IsBodyHTML, string strAttachmentPath)
        {
            Boolean blnSendemail = false;
            string strSaveAttachmentTo = "";
            string strFileName = "";
            Boolean HasAttachment = false;           
            ISBL.General oISBLGen = new ISBL.General();
            try
            {

                strSaveAttachmentTo = ConfigurationManager.AppSettings["RobustEmailFTPFolder"];
                if (!Directory.Exists(strSaveAttachmentTo))
                {
                    Directory.CreateDirectory(strSaveAttachmentTo);
                }
                strFileName = System.IO.Path.GetFileName(strAttachmentPath);
                if (strFileName.Length > 0)
                {
                    HasAttachment = true;
                    if (!File.Exists(strSaveAttachmentTo + strFileName))
                    {
                        File.Copy(strAttachmentPath, strSaveAttachmentTo + strFileName);
                    }

                }
                blnSendemail = oISBLGen.AddEmailStorage(strTo, strFrom, strSubject, strBody, IsBodyHTML, strFileName, HasAttachment);

                return blnSendemail;
            }
            catch (Exception ex)
            {
                oISBLGen.errTrack(ex);
                return false;
            }
            finally
            {
                //Dispose the MailMessage Object
                //for Robust Eamil enhancement 19 Jan 2009 Adam 
                //-- mm.Dispose();
                oISBLGen.DisposeConnection();
                oISBLGen = null;
            }
        }

        //Adam - 18 August 2010
        public String GetLoginIDEmail(string LoginID)
        {
            string LoginIDEmail;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetLoginIDEmail";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = this.SafeSqlLiteral(LoginID);
            conn.Open();
            conn.callingMethod = "ISBL.General.GetLoginIDEmail";
            LoginIDEmail = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return LoginIDEmail;
        }


        /* Start ADFS */
        public Boolean ISADFSClient(string strClientCode)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsADFSClient";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.ISADFSClient";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }

        }

        public DataSet GetADFSLoginInfo(string strUPN)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetADFSLoginInfo";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@UPN", SqlDbType.VarChar, 100);
            myCmd.Parameters["@UPN"].Value = strUPN;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetADFSLoginInfo";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean ISADFSLoginID(string strLoginID)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsADFSLoginID";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.ISADFSLoginID";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
         }

        public Boolean ISADFSCRN(string strCRN)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsADFSCRN";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters["@CRN"].Value = strCRN;
            conn.Open();
            conn.callingMethod = "ISBL.General.ISADFSCRN";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }


        public Boolean ISUPNExist(string strUPN, string strLoginID)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsUPNExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@UPN", SqlDbType.VarChar, 100);
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@UPN"].Value = strUPN;
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.ISUPNExist";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }

        public Boolean IsAtlasCRN(string strCRN, DateTime AtlasLiveDate)
        {
            string strStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsAtlasCRN";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 80);
            myCmd.Parameters.Add("@AtlasLiveDate", SqlDbType.DateTime);
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters["@AtlasLiveDate"].Value = AtlasLiveDate;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsAtlasCRN";
            strStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            if (strStatus.Length != 0)
            {
                if (strStatus == "1")
                {
                    myCmd.Dispose();
                    return true;
                }
                else
                {
                    myCmd.Dispose();
                    return false;
                }
            }
            else
            {
                myCmd.Dispose();
                return false;
            }
        }


        public String GetOrderPrimaryVariant(string OrderID)
        {
            string PrimaryVariant;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOrderPrimaryVariant";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@OrderID"].Value = new System.Guid(OrderID);
            conn.Open();
            conn.callingMethod = "ISBL.General.GetOrderPrimaryVariant";
            PrimaryVariant = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return PrimaryVariant;
        }

        public String GetOrderPrimaryCountry(string OrderID)
        {
            string PrimaryCountry;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOrderPrimaryCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@OrderID"].Value = new System.Guid(OrderID);
            conn.Open();
            conn.callingMethod = "ISBL.General.GetOrderPrimaryCountry";
            PrimaryCountry = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return PrimaryCountry;
        }

        /* End ADFS */

        /* Start ISIS v2 Phase 1 Release 1*/
        public DataSet GetOrderedBy(string strClientCode, string strLoginID, string strLoginUserRole)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetOrderedBy";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters.Add("@LoginUserRole", SqlDbType.VarChar, 5);
            myCmd.Parameters["@LoginUserRole"].Value = strLoginUserRole;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetOrderedBy";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public Boolean IsBulkOrder(Guid gOrderID)
        {
            String sStatus;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsBulkOrder";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@OrderID"].Value = gOrderID;

            conn.Open();
            conn.callingMethod = "ISBL.General.IsBulkOrder";
            sStatus = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            if (sStatus.Length != 0)
            {
                if (sStatus == "1" || sStatus == "True")
                    return true;
                else
                    return false;
            }
            else
            {
                return false;
            }
        }
        /* Start ISIS v2 Phase 1 Release 1*/


        /* Start ISIS v2 Phase 1 Release 3*/
        public String GetClientBulkOrderDueDay(string strReportType, string strClientCode)
        {
            string strDuedate;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientBulkOrderDueDay";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientBulkOrderDueDay";
            strDuedate = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strDuedate;
        }
        /* End ISIS v2 Phase 1 Release 3*/

        /**********Mudassar Khan--Code---End**************************************/


        /* Start ISIS Password Enhancement Oct 2012*/
        public Boolean IsLoginIDBlock(string strLoginID)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsLoginIDBlock";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsLoginIDBlock";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        public Boolean IsPasswordExpired(string strLoginID)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsPasswordExpired";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsPasswordExpired";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }


        public Boolean IsValidLogin(string strLoginID)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsValidLogin";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsValidLogin";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        public string GetLoginUserType(string strLoginID)
        {
            string strPassword;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetLoginUserType";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetLoginUserType";
            strPassword = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strPassword;
        }

        public Boolean UpdateFailedLoginAttempt(string strLoginID, bool blnReset, int intIncrementValue)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_UpdateFailedLoginAttempt";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
            myCmd.Parameters["@LoginID"].Value = strLoginID;
            myCmd.Parameters.Add("@Reset", SqlDbType.Bit);
            myCmd.Parameters["@Reset"].Value = blnReset;
            myCmd.Parameters.Add("@IncrementValue", SqlDbType.Int);
            myCmd.Parameters["@IncrementValue"].Value = intIncrementValue;
            conn.Open();
            conn.callingMethod = "ISBL.General.UpdateFailedLoginAttempt";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }

        /* End ISIS Password Enhancement Oct 2012*/

        /* Start ISIS Custom Field Enhancement Feb 2013*/
        public DataSet GetClientCustomField(string strClientCode, string strActive)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientCustomField";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters.Add("@Active", SqlDbType.Char, 1);
            myCmd.Parameters["@Active"].Value = strActive.ToUpper();
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientCustomField";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetClientOrderCustomField(string strOrderID, string strClientCode)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetClientOrderCustomField";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@OrderID"].Value = new System.Guid(strOrderID);
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetClientOrderCustomField";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        public DataSet GetRegularExpression()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetRegularExpression";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetRegularExpression";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /* End ISIS Custom Field Enhancement Feb 2013*/

        //Mitul - 17May13 - Replace % sign with per cent word - START
        // Sunil - 4May15 - Added to replace & with and word
        public String ReplaceSignWithWord(String strInput)
        {
            return strInput.Replace("%", ConfigurationManager.AppSettings["PercentSignConversion"]).Replace("&"," and ");
        }
        //Mitul - 17May13 - Replace % sign with per cent word - END
        /*Mitul - 1July13 - Multi Language user Guide - START*/
        public DataSet GetUserGuideList()
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetUserGuide";
            myCmd.CommandType = CommandType.StoredProcedure;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetUserGuideList";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }
        /*Mitul - 1July13 - Multi Language user Guide - END*/

        /*Mitul - Aug13 - ExcelQuery Fix - START*/
        public String GetExcelQuery(string strTableName)
        {
            string strReturnQuery;
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetExcelQuery";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@TableName", SqlDbType.VarChar);
            myCmd.Parameters["@TableName"].Value = strTableName;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetExcelQuery";
            strReturnQuery = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strReturnQuery;
        }
        /*Mitul - Aug13 - ExcelQuery Fix - END*/

        /* Start Adam - Jul 2013 - ISIS and Atlas Sync*/
        public Boolean IsMasterCodeExist(string strCode, string strMaster)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsMasterCodeExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Code", SqlDbType.VarChar, 15);
            myCmd.Parameters["@Code"].Value = strCode;
            myCmd.Parameters.Add("@Master", SqlDbType.VarChar, 20);
            myCmd.Parameters["@Master"].Value = strMaster;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsReportTypeExist";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }
        /* End Adam - Jul 2013 - ISIS and Atlas Sync*/

        /* Start ISIS Enhancement - Oct 2014 NAGARAJ*/
        public Boolean IsMasterCodeExistValid(string strCode, string strMaster)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsMasterCodeExistValid";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Code", SqlDbType.VarChar, 15);
            myCmd.Parameters["@Code"].Value = strCode;
            myCmd.Parameters.Add("@Master", SqlDbType.VarChar, 20);
            myCmd.Parameters["@Master"].Value = strMaster;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsReportTypeExist";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }
        /* End ISIS Enhancement - Oct 2014 NAGARAJ*/

        /* Start Mercury Enhancement - Feb 2014*/
         public Boolean IsAtlasAlive()
         {
             DataSet dsCredential = new DataSet();
             ISBL.BizLog bizlog = new ISBL.BizLog();
             dsCredential = bizlog.GetSOAPCredential();


             ISBL.AtlasWS.UserDetails UserDetails = new ISBL.AtlasWS.UserDetails();
             ISBL.AtlasWS.AtlasWebServiceService WS = new ISBL.AtlasWS.AtlasWebServiceService();
             ISBL.AtlasWS.CaseResultVO ResultVO = new ISBL.AtlasWS.CaseResultVO();

             UserDetails.userName = dsCredential.Tables[0].Rows[0]["UserName"].ToString();
             UserDetails.password = dsCredential.Tables[0].Rows[0]["Password"].ToString();

             WS.UserDetailsValue = UserDetails;
             WS.Timeout = 120000;

             Boolean bpingAtlas = false;
             try
             {
                 if (WS.pingAtlas() == "Ping_0")
                     bpingAtlas = true;
                 else
                     bpingAtlas = false;

             }
             catch
             {
                 bpingAtlas = false;
             }

             return bpingAtlas;
         }

         public List<FinalReportObjects> DownloadFinalReport(string strCRN, string strLoginID) 
         {
             ISBL.General oISBL = new ISBL.General();
             ISBL.BizLog oBiz = new ISBL.BizLog();
             List<FinalReportObjects> listFinalReport = new List<FinalReportObjects> { };
             FinalReportObjects objFinalReport = new FinalReportObjects();

             System.Guid EventLogGUID;
             EventLogGUID = System.Guid.NewGuid();

             String NewData = "";
             NewData += "<Parent><Files>";

             try
             {

  
                 ISBL.AtlasWS.UserDetails UserDetails = new ISBL.AtlasWS.UserDetails();
                 ISBL.AtlasWS.AtlasWebServiceService WS = new ISBL.AtlasWS.AtlasWebServiceService();
                 ISBL.AtlasWS.DownloadOnlineReportResultVO ResultVO = new ISBL.AtlasWS.DownloadOnlineReportResultVO();
                 ISBL.AtlasWS.DownloadOnlineReportVO DownloadOnlineReportVO = new ISBL.AtlasWS.DownloadOnlineReportVO();


                 Boolean bPingAtlas = false;
 
                 DataSet dsFinalReport = new DataSet();

      
                 DataSet dsCrendetial = new DataSet();
                 dsCrendetial = oBiz.GetSOAPCredential();
                 foreach (DataRow row in dsCrendetial.Tables[0].Rows)
                 {
                     UserDetails.userName = row["UserName"].ToString();
                     UserDetails.password = row["Password"].ToString();
                 }
                 
                 dsFinalReport = oBiz.GetClientOrderFinalReportList(strCRN);

                 if (dsFinalReport != null)
                 {
                     if (dsFinalReport.Tables[0].Rows.Count > 0)
                     {
                         foreach (DataRow row in dsFinalReport.Tables[0].Rows)
                         {
                             NewData += "<File>";
                             NewData += "<CRN>" + row["CRN"].ToString() + "</CRN>";
                             NewData += "<FileName>" + row["FileName"].ToString() + "</FileName>";
                             NewData += "<Version>" + row["Version"].ToString() + "</Version>";
                             NewData += "</File>";

                             DownloadOnlineReportVO.crn = row["CRN"].ToString();
                             DownloadOnlineReportVO.fileName = row["FileName"].ToString();
                             DownloadOnlineReportVO.version = row["Version"].ToString();

                             WS.UserDetailsValue = UserDetails;
                             WS.Timeout = 120000; //2 minutes


                             try
                             {
                                 if (WS.pingAtlas() == "Ping_0")
                                 {
                                     bPingAtlas = true;
                                 }
                             }
                             catch (Exception ex)
                             {
                                 bPingAtlas = false;
                             }

                             if (bPingAtlas)
                             {
                                 ResultVO = WS.downloadOnlineReport(DownloadOnlineReportVO);
                                 if (strCRN == ResultVO.crn)
                                 {
                                     if (ResultVO.errorCode == "DOR_0")
                                     {
                                         objFinalReport.FinalReportFileName = DownloadOnlineReportVO.fileName;
                                         objFinalReport.FinalReportBase64StringContent = System.Convert.ToBase64String(ResultVO.fileContent, 0, ResultVO.fileContent.Length);
                                         listFinalReport.Add(objFinalReport);
                                     }
                                     else
                                     {
                                         System.Threading.Thread.Sleep(30000); //'wait for 30 seconds
                                         ResultVO = WS.downloadOnlineReport(DownloadOnlineReportVO);
                                         if (ResultVO.errorCode == "DOR_0")
                                         {
                                             objFinalReport.FinalReportFileName = DownloadOnlineReportVO.fileName;
                                             objFinalReport.FinalReportBase64StringContent = System.Convert.ToBase64String(ResultVO.fileContent, 0, ResultVO.fileContent.Length);
                                             listFinalReport.Add(objFinalReport);
                                         }
                                         else
                                         {
                                             NewData += "</Files></Parent>";
                                             oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_" + ResultVO.errorCode, ResultVO.errorMessage, "");
                                             return null;
                                         }
                                     }
                                 }
                                 else
                                 {
                                     NewData += "</Files></Parent>";
                                     oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_2", "Wrong CRN", "");
                                     return null;
                                 }
                             }
                             else
                             {
                                 NewData += "</Files></Parent>";
                                 oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_3", "Failed to Ping Atlas", "");
                                 return null;
                             }
                         }

                         NewData += "</Files></Parent>";
                         oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_0", "", "");
                         return listFinalReport;
                     }
                     else
                     {
                         NewData += "</Files></Parent>";
                         oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_4", "The order does not contain Final Report.", "");
                         return null;                     
                     }
                 }
                 else
                 {
                     NewData += "</Files></Parent>";
                     oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_4", "The order does not contain Final Report.", "");
                     return null;
                 }                 
             }
             catch (Exception ex)
             {
                 NewData += "</Files></Parent>";
                 oISBL.SaveEventLog(EventLogGUID.ToString(), "", strCRN, "Download Report", "", NewData, System.DateTime.Now, System.DateTime.Now, strLoginID, "", "DFR_1", ex.Message, "");
                 return null;
             }
             finally
             {
                 oISBL.DisposeConnection();
                 oBiz.DisposeConnection();
                 oISBL = null;
                 oBiz = null;
             }
         }
        
        /* End Mercury Enhancement - Feb 2014 */

         /* Start fixing OWASP Mar 2014*/
         public Boolean IsLoginIDBDMAccess(string strLoginID, string strBDM)
         {
             SqlCommand sqlCmd = new SqlCommand();
             Boolean bStatus;

             try
             {
                 sqlCmd.Connection = conn.Connection;
                 sqlCmd.CommandType = CommandType.StoredProcedure;
                 sqlCmd.CommandText = "sp_IsLoginIDBDMAccess";

                 sqlCmd.Parameters.Add("@LoginID", SqlDbType.VarChar, 15);
                 sqlCmd.Parameters["@LoginID"].Value = strLoginID;
                 sqlCmd.Parameters.Add("@BDM", SqlDbType.VarChar, 15);
                 sqlCmd.Parameters["@BDM"].Value = strBDM; 

                 conn.Open();
                 conn.callingMethod = "ISBL.BizLog.IsLoginIDBDMAccess";
                 bStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(sqlCmd));
                 sqlCmd.Dispose();
                 conn.Close();

                 return bStatus;
             }
             catch
             {
                 return false;
             }
         }
         /* End fixing OWASP Mar 2014*/

        /* Start BI 34 Budget & TAT recalculation */
         public String CalculateBudget(string strClientCode, string strReportType, string PrimaryVariant, string PrimarySubjectCountry, string strSubReportType, Boolean isExpress, Boolean isBulkOrder, DataView dvSubjectList, ref  Boolean isCalculationRevert, ref string strDueDate)
         {
             SqlCommand sqlCmd = new SqlCommand();
             try
             {
                 DataSet dsbuddet = new DataSet();
                 SqlCommand myCmd = new SqlCommand();
                 myCmd.Connection = conn.Connection;
                 myCmd.CommandText = "sp_GetBudgetDetails";
                 myCmd.CommandType = CommandType.StoredProcedure;
                 myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                 myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
                 myCmd.Parameters.Add("@SubjectCountry", SqlDbType.VarChar, 4);
                 myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubReportTypeForSpecialCalculation", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubjectTypeForSpecialCalculation", SqlDbType.TinyInt);
                 myCmd.Parameters.Add("@SubjectCountryForSpecialCalculation", SqlDbType.VarChar, 4);

                 myCmd.Parameters["@ClientCode"].Value = strClientCode;
                 myCmd.Parameters["@ReportType"].Value = strReportType;
                 myCmd.Parameters["@Variant"].Value = PrimaryVariant;
                 myCmd.Parameters["@SubjectCountry"].Value = PrimarySubjectCountry;
                 myCmd.Parameters["@SubReportType"].Value = strSubReportType;
                 myCmd.Parameters["@SubReportTypeForSpecialCalculation"].Value = "All";
                 myCmd.Parameters["@SubjectTypeForSpecialCalculation"].Value = 0;
                 myCmd.Parameters["@SubjectCountryForSpecialCalculation"].Value = "All";

                 SqlDataAdapter sda = new SqlDataAdapter();
                 sda.SelectCommand = myCmd;
                 conn.Open();
                 conn.callingMethod = "ISBL.General.CalculateBudget_sp_GetBudgetDetails";
                 dsbuddet = conn.FillDataSet(sda);
                 conn.Close();
                 myCmd.Dispose();
                 sda.Dispose();

                 String calculationType = "1";
                 if (dsbuddet.Tables[1].Rows.Count > 0){
                     calculationType = dsbuddet.Tables[1].Rows[0]["Calculation"].ToString();
                     isCalculationRevert = Convert.ToBoolean(dsbuddet.Tables[1].Rows[0]["IsCalculationRevert"].ToString());
                 }

                 String strBud = "";
                 float flBasePrice=0;
                 float IncPriceCompany = 0;
                 float IncPriceIndividual = 0;
                 int intBaseCompany = 0;
                 int intBaseIndividual = 0;
                 int intActualCompany = 0;
                 int intActualIndividual = 0;
                 int intTAT = 0;

                 if (calculationType != "3")
                 {
                     DataTable tblSubject = new DataTable();
                     DataView dvTemp = new DataView();
                     dvTemp = dvSubjectList;
                     tblSubject = dvTemp.Table;


                     foreach (DataRow dr in tblSubject.Rows)
                     {
                         if (dr["ResearchType"].ToString() == "Company") 
                         {
                              intActualCompany = intActualCompany + 1;
                         } 
                         else
                         {
                              intActualIndividual = intActualIndividual + 1;
                         }

                     }

                    if (dsbuddet.Tables[0].Rows.Count > 0){
                        if (isExpress) {
                            flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePriceExpress"]);
                            IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompanyExpress"]);
                            IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividualExpress"]);
                            intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePriceExpress"]);
                            intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePriceExpress"]);
                            if (isBulkOrder)
                                intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                            else
                                intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATExpress"]);
                        } 
                        else 
                        {
                            flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePrice"]);
                            IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompany"]);
                            IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividual"]);
                            intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePrice"]);
                            intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePrice"]);
                            if (isBulkOrder)
                                intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                            else
                                intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormal"]);
                        }
                    }
                 }
                 switch (calculationType)
                 {
                     case "1": case "2":
                         strBud = GetBudget(flBasePrice, IncPriceCompany, IncPriceIndividual, intBaseCompany, intBaseIndividual, intActualCompany, intActualIndividual).ToString();
                         if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                         {
                             int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                             if (intHour >= GMTTime(strClientCode))
                             {
                                 strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                             }
                             else
                             {
                                 if (intTAT <= 1)
                                 {
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                 }
                                 else
                                 {
                                     intTAT = intTAT - 1;
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                                 }
                             }
                         }
                         else
                         {
                             strDueDate =  this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                         }
                        break;
                     case "3":
                         string strSubjectCountry = "";
                         string strSubjectType = "";
                         DataSet dsBudDetCal = new DataSet();
                         double dblBudget = 0;
                         int intBudDetCalTAT = 0;
                         int intTempTAT = 0;
                        foreach (DataRowView rowView in dvSubjectList)
                        {
                            DataRow row = rowView.Row;
                            strSubjectType = row["ResearchType"].ToString();
                            strSubjectCountry = GetCountry(row["CountryName"].ToString().Trim());
                            dsBudDetCal = GetBudgetDetailsByCountry(strClientCode, strReportType, 3, strSubjectCountry, strSubjectType, isExpress);
                            if (dsBudDetCal.Tables[0].Rows.Count > 0)
                            {
                                dblBudget += Double.Parse(dsBudDetCal.Tables[0].Rows[0]["Cost"].ToString());
                                if (isBulkOrder)
                                    intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TATBulk"].ToString());
                                else
                                    intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TAT"].ToString()); 
                            }
                            else
                            {
                                dblBudget += 0;
                                intTempTAT = 0;
                            }
                            strSubjectCountry = "";
                            strSubjectType = "";
                            if (intTempTAT > intBudDetCalTAT)
                            {
                                intBudDetCalTAT = intTempTAT;
                                intTempTAT = 0;
                            }

                        }
                        strBud = dblBudget.ToString();
                        if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                        {
                            int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                            if (intHour >= GMTTime(strClientCode))
                            {
                                strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                            }
                            else
                            {
                                if (intBudDetCalTAT <= 1)
                                {
                                    strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                }
                                else
                                {
                                    intBudDetCalTAT = intBudDetCalTAT - 1;
                                    strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                                }
                            }
                        }
                        else
                        {
                            strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                        }

                        break;
                 }

                 return strBud;
             }
             catch 
             {
                 return "0";
             }
         }
         public String CalculateFlsBudget(string strClientCode, string strReportType, string strSubReportType, string PrimarySubjectCountry, string strSubjectType, Boolean isExpress, Boolean isBulkOrder, DataView dvSubjectList, ref  Boolean isflsCalculationRevert, ref string strflsDueDate, int intCountryCounter)                      
         {
             SqlCommand sqlCmd = new SqlCommand();
             try
             {
                 DataSet dsbuddet = new DataSet();
                 SqlCommand myCmd = new SqlCommand();
                 myCmd.Connection = conn.Connection;
                 myCmd.CommandText = "sp_GetFlsBudgetDetails";
                 myCmd.CommandType = CommandType.StoredProcedure;
                 myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                 myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubjectCountry", SqlDbType.VarChar, 10);
                 myCmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubReportTypeForSpecialCalculation", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubjectTypeForSpecialCalculation", SqlDbType.TinyInt);
                 myCmd.Parameters.Add("@SubjectCountryForSpecialCalculation", SqlDbType.VarChar, 4);

                 myCmd.Parameters["@ClientCode"].Value = strClientCode;
                 myCmd.Parameters["@ReportType"].Value = strReportType;
                 myCmd.Parameters["@SubReportType"].Value = strSubReportType;
                 myCmd.Parameters["@SubjectCountry"].Value = PrimarySubjectCountry;
                 myCmd.Parameters["@SubjectType"].Value = strSubjectType;
                 myCmd.Parameters["@SubReportTypeForSpecialCalculation"].Value = "All";
                 myCmd.Parameters["@SubjectTypeForSpecialCalculation"].Value = 0;
                 myCmd.Parameters["@SubjectCountryForSpecialCalculation"].Value = "All";

                 SqlDataAdapter sda = new SqlDataAdapter();
                 sda.SelectCommand = myCmd;
                 conn.Open();
                 conn.callingMethod = "ISBL.General.CalculateBudget_sp_GetFlsBudgetDetails";
                 dsbuddet = conn.FillDataSet(sda);
                 conn.Close();
                 myCmd.Dispose();
                 sda.Dispose();

                 String calculationType = "1";
                 if (dsbuddet.Tables[1].Rows.Count > 0)
                 {
                     calculationType = dsbuddet.Tables[1].Rows[0]["Calculation"].ToString();
                     isflsCalculationRevert = Convert.ToBoolean(dsbuddet.Tables[1].Rows[0]["IsCalculationRevert"].ToString());
                 }

                 String strBud = "";
                 float flBasePrice = 0;
                 float IncPriceCompany = 0;
                 float IncPriceIndividual = 0;
                 int intBaseCompany = 0;
                 int intBaseIndividual = 0;
                 int intActualCompany = 0;
                 int intActualIndividual = 0;
                 int intTAT = 0;

                 if (calculationType != "3")
                 {
                     DataTable tblSubject = new DataTable();
                     DataView dvTemp = new DataView();
                     dvTemp = dvSubjectList;
                     dvTemp.RowFilter = "CountryName = '" + GetCountryName(PrimarySubjectCountry) +"'";
                    
                     tblSubject = dvTemp.ToTable();

                     
                     foreach (DataRow dr in tblSubject.Rows)
                     {
                         //if (dr["ResearchType"].ToString() == "Company") // commented by deepak
                         if (dr["ResearchType"].ToString() == "2" || dr["ResearchType"].ToString() == "Company")
                         {
                            
                             intActualCompany = intActualCompany + 1;
                         }
                         else
                         {
                             intActualIndividual = intActualIndividual + 1;
                         }

                     }

                     //newly addded to calculate the budget
                     //if (strSubjectType == "2" || strSubjectType == "Company")
                     //{
                     //    intActualCompany = intActualCompany + 1;
                     //}
                     //else
                     //{
                     //    intActualIndividual = intActualIndividual + 1;
                     //}

                     if (dsbuddet.Tables[0].Rows.Count > 0)
                     {
                         if (isExpress)
                         {
                             flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePriceExpress"]);
                             IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompanyExpress"]);
                             IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividualExpress"]);
                             intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePriceExpress"]);
                             intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePriceExpress"]);
                             if (isBulkOrder)
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                             else
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATExpress"]);
                         }
                         else
                         {
                             flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePrice"]);
                             IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompany"]);
                             IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividual"]);
                             intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePrice"]);
                             intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePrice"]);
                             if (isBulkOrder)
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                             else
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormal"]);
                         }
                     }
                 }
                 switch (calculationType)
                 {
                     case "1":
                     case "2":
                         //check entity type first
                         if (strSubjectType == "2" || strSubjectType == "Company")
                         {
                             if (intBaseCompany == 0)
                                 strBud = IncPriceCompany.ToString();
                             else if (intActualCompany <= intBaseCompany)
                                 strBud = "0";
                             else if (intCountryCounter <= intBaseCompany)
                                 strBud = "0";
                             else
                                 strBud = IncPriceCompany.ToString();
                         }
                         else
                         {
                             if (intBaseIndividual == 0)
                                 strBud = IncPriceIndividual.ToString();
                             else if (intActualIndividual <= intBaseIndividual)
                                 strBud = "0";
                             else if (intCountryCounter <= intBaseIndividual)
                                 strBud = "0";
                             else 
                                 strBud = IncPriceIndividual.ToString();
                                      
                         }

                         
                         if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                         {
                             int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                             if (intHour >= GMTTime(strClientCode))
                             {
                                 strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                             }
                             else
                             {
                                 if (intTAT <= 1)
                                 {
                                     strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                 }
                                 else
                                 {
                                     intTAT = intTAT - 1;
                                     strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                                 }
                             }
                         }
                         else
                         {
                             strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                         }
                         break;
                     case "3":
                         string strSubjectCountry = "";
                         string strSubjectType1 = "";
                         DataSet dsBudDetCal = new DataSet();
                         double dblBudget = 0;
                         int intBudDetCalTAT = 0;
                         int intTempTAT = 0;
                         foreach (DataRowView rowView in dvSubjectList)
                         {
                             DataRow row = rowView.Row;
                             strSubjectType = row["ResearchType"].ToString();
                             strSubjectCountry = GetCountry(row["CountryName"].ToString().Trim());
                             strSubReportType = row["SubreportType"].ToString();
                             //////////////////////////////////////////////////////////////***************************************************///////////////////////////////////////////
                             dsBudDetCal = GetflsBudgetDetailsByCountry(strClientCode, strReportType,strSubReportType, 3, strSubjectCountry, strSubjectType, isExpress);
                             //////////////////////////////////////////////////////////////***************************************************///////////////////////////////////////////
                             if (dsBudDetCal.Tables[0].Rows.Count > 0)
                             {
                                 dblBudget += Double.Parse(dsBudDetCal.Tables[0].Rows[0]["Cost"].ToString());
                                 if (isBulkOrder)
                                     intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TATBulk"].ToString());
                                 else
                                     intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TAT"].ToString());
                             }
                             else
                             {
                                 dblBudget += 0;
                                 intTempTAT = 0;
                             }
                             strSubjectCountry = "";
                             strSubjectType = "";
                             if (intTempTAT > intBudDetCalTAT)
                             {
                                 intBudDetCalTAT = intTempTAT;
                                 intTempTAT = 0;
                             }

                         }
                         strBud = dblBudget.ToString();
                         if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                         {
                             int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                             if (intHour >= GMTTime(strClientCode))
                             {
                                 strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                             }
                             else
                             {
                                 if (intBudDetCalTAT <= 1)
                                 {
                                     strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                 }
                                 else
                                 {
                                     intBudDetCalTAT = intBudDetCalTAT - 1;
                                     strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                                 }
                             }
                         }
                         else
                         {
                             strflsDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                         }

                         break;
                 }

                 return strBud;
             }
             catch
             {
                 return "0";
             }
         }
        //CalculatetotBudget
         public String CalculatetotBudget(string strClientCode, Guid strOrderID ,string strReportType, string PrimaryVariant, string PrimarySubjectCountry, string strSubReportType, Boolean isExpress, Boolean isBulkOrder, DataView dvSubjectList, ref  Boolean isCalculationRevert, ref string strDueDate)
         {
             SqlCommand sqlCmd = new SqlCommand();
             try
             {
                 DataSet dsbuddet = new DataSet();
                 SqlCommand myCmd = new SqlCommand();
                 myCmd.Connection = conn.Connection;
                 myCmd.CommandText = "sp_GetBulkTotBudgetDetails";
                 myCmd.CommandType = CommandType.StoredProcedure;
                 myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
                 myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@OrderID", SqlDbType.UniqueIdentifier);
                 myCmd.Parameters.Add("@Variant", SqlDbType.VarChar, 250);
                 myCmd.Parameters.Add("@SubjectCountry", SqlDbType.VarChar, 4);
                 myCmd.Parameters.Add("@SubReportType", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubReportTypeForSpecialCalculation", SqlDbType.VarChar, 15);
                 myCmd.Parameters.Add("@SubjectTypeForSpecialCalculation", SqlDbType.TinyInt);
                 myCmd.Parameters.Add("@SubjectCountryForSpecialCalculation", SqlDbType.VarChar, 4);

                 myCmd.Parameters["@ClientCode"].Value = strClientCode;
                 myCmd.Parameters["@ReportType"].Value = strReportType;
                 myCmd.Parameters["@OrderID"].Value = strOrderID;
                 myCmd.Parameters["@Variant"].Value = PrimaryVariant;
                 myCmd.Parameters["@SubjectCountry"].Value = PrimarySubjectCountry;
                 myCmd.Parameters["@SubReportType"].Value = strSubReportType;
                 myCmd.Parameters["@SubReportTypeForSpecialCalculation"].Value = "All";
                 myCmd.Parameters["@SubjectTypeForSpecialCalculation"].Value = 0;
                 myCmd.Parameters["@SubjectCountryForSpecialCalculation"].Value = "All";

                 SqlDataAdapter sda = new SqlDataAdapter();
                 sda.SelectCommand = myCmd;
                 conn.Open();
                 conn.callingMethod = "ISBL.General.CalculateBudget_sp_GetBudgetDetails";
                 dsbuddet = conn.FillDataSet(sda);
                 conn.Close();
                 myCmd.Dispose();
                 sda.Dispose();

                 String calculationType = "1";
                 if (dsbuddet.Tables[1].Rows.Count > 0)
                 {
                     calculationType = dsbuddet.Tables[1].Rows[0]["Calculation"].ToString();
                     isCalculationRevert = Convert.ToBoolean(dsbuddet.Tables[1].Rows[0]["IsCalculationRevert"].ToString());
                 }

                 String strBud = "";
                 float flBasePrice = 0;
                 float IncPriceCompany = 0;
                 float IncPriceIndividual = 0;
                 int intBaseCompany = 0;
                 int intBaseIndividual = 0;
                 int intActualCompany = 0;
                 int intActualIndividual = 0;
                 int intTAT = 0;

                 if (calculationType != "3")
                 {
                     DataTable tblSubject = new DataTable();
                     DataView dvTemp = new DataView();
                     dvTemp = dvSubjectList;
                     tblSubject = dvTemp.Table;


                     foreach (DataRow dr in tblSubject.Rows)
                     {
                         if (dr["ResearchType"].ToString() == "Company")
                         {
                             intActualCompany = intActualCompany + 1;
                         }
                         else
                         {
                             intActualIndividual = intActualIndividual + 1;
                         }

                     }

                     if (dsbuddet.Tables[0].Rows.Count > 0)
                     {
                         if (isExpress)
                         {
                             flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePriceExpress"]);
                             IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompanyExpress"]);
                             IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividualExpress"]);
                             intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePriceExpress"]);
                             intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePriceExpress"]);
                             if (isBulkOrder)
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                             else
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATExpress"]);
                         }
                         else
                         {
                             flBasePrice = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["BasePrice"]);
                             IncPriceCompany = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerCompany"]);
                             IncPriceIndividual = Convert.ToSingle(dsbuddet.Tables[0].Rows[0]["AdditionalCostPerIndividual"]);
                             intBaseCompany = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxCompanyForBasePrice"]);
                             intBaseIndividual = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["MaxIndividualForBasePrice"]);
                             if (isBulkOrder)
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                             else
                                 intTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormal"]);
                         }
                     }
                 }
                 switch (calculationType)
                 {
                     case "1":
                     case "2":
                         strBud = GetBudget(flBasePrice, IncPriceCompany, IncPriceIndividual, intBaseCompany, intBaseIndividual, intActualCompany, intActualIndividual).ToString();
                         if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                         {
                             int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                             if (intHour >= GMTTime(strClientCode))
                             {
                                 strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                             }
                             else
                             {
                                 if (intTAT <= 1)
                                 {
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                 }
                                 else
                                 {
                                     intTAT = intTAT - 1;
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                                 }
                             }
                         }
                         else
                         {
                             strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intTAT, "", false)));
                         }
                         break;
                     case "3":
                         string strSubjectCountry = "";
                         string strSubjectType = "";
                         DataSet dsBudDetCal = new DataSet();
                         double dblBudget = 0;
                         int intBudDetCalTAT = 0;
                         int intTempTAT = 0;
                         foreach (DataRowView rowView in dvSubjectList)
                         {
                             DataRow row = rowView.Row;
                             strSubjectType = row["ResearchType"].ToString();
                             strSubjectCountry = GetCountry(row["CountryName"].ToString().Trim());
                             //dsBudDetCal = GetBudgetDetailsByCountry(strClientCode, strReportType, 3, strSubjectCountry, strSubjectType, isExpress);
                             dsBudDetCal = GetFinalBudgetDetails(strClientCode, strOrderID);
                             if (dsBudDetCal.Tables[0].Rows.Count > 0)
                             {
                                 dblBudget += Double.Parse(dsBudDetCal.Tables[0].Rows[0]["Cost"].ToString());
                                 if (isBulkOrder)
                                     intTempTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATNormalBulk"]);
                                 else
                                     intTempTAT = Convert.ToInt16(dsbuddet.Tables[0].Rows[0]["TATExpress"]);
                                 //if (isBulkOrder)
                                 //    intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TATBulk"].ToString());
                                 //else
                                 //    intTempTAT = int.Parse(dsBudDetCal.Tables[0].Rows[0]["TAT"].ToString());
                             }
                             else
                             {
                                 dblBudget += 0;
                                 intTempTAT = 0;
                             }
                             strSubjectCountry = "";
                             strSubjectType = "";
                             if (intTempTAT > intBudDetCalTAT)
                             {
                                 intBudDetCalTAT = intTempTAT;
                                 intTempTAT = 0;
                             }

                         }
                         strBud = dblBudget.ToString();
                         if (ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Contains(strClientCode))
                         {
                             int intHour = Convert.ToInt16(System.DateTime.Now.ToString("HH"));
                             if (intHour >= GMTTime(strClientCode))
                             {
                                 strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                             }
                             else
                             {
                                 if (intBudDetCalTAT <= 1)
                                 {
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(System.DateTime.Now.Date.ToString()));
                                 }
                                 else
                                 {
                                     intBudDetCalTAT = intBudDetCalTAT - 1;
                                     strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                                 }
                             }
                         }
                         else
                         {
                             strDueDate = this.FormatDDMMMYYYY(System.DateTime.Parse(GetDueDate(System.DateTime.Now.Date, intBudDetCalTAT, "", false)));
                         }

                         break;
                 }

                 DataSet dsBudDetCalnew = new DataSet();                
                 dsBudDetCalnew = GetFinalBudgetDetails(strClientCode, strOrderID);
                 if (dsBudDetCalnew.Tables[0].Rows.Count > 0)
                 {
                     strBud = dsBudDetCalnew.Tables[0].Rows[0]["Totbudget"].ToString();
                 }

                 return strBud;
             }
             catch
             {
                 return "0";
             }
         }
        public int GMTTime(string strClientCode) 
        {
            Array arrClientTimelist =  ConfigurationManager.AppSettings["ClientGMTTimeForDueDateCalculation"].ToString().Split(';');
            Array arrClientTime;
            int intGMTTime = 0;
            foreach (string strClientTime in arrClientTimelist) {
                if (strClientTime.Contains(strClientCode))
                {
                    arrClientTime = strClientTime.Split(':');
                    foreach (string strValue in arrClientTime)
                    {
                        if (strValue != strClientCode) 
                            intGMTTime = Convert.ToInt16(strValue);
                    }
                 }
            }    
            return intGMTTime;
        }

        public System.Data.DataView ConvertSubjectArrayToDataView(System.Collections.ArrayList arrList)
        {
            DataTable table = new DataTable("table");
            DataColumn column1;
            DataColumn column2;
            DataColumn column3;
            DataRow row;

            column1 = new DataColumn();
            column1.DataType = Type.GetType("System.String");
            column1.ColumnName = "ResearchType";
            table.Columns.Add(column1);

            column2 = new DataColumn();
            column2.DataType = Type.GetType("System.String");
            column2.ColumnName = "SubjectName";
            table.Columns.Add(column2);

            column3 = new DataColumn();
            column3.DataType = Type.GetType("System.String");
            column3.ColumnName = "CountryName";
            table.Columns.Add(column3);

            foreach(System.Collections.Hashtable hastable in arrList)
            {
                row = table.NewRow();
                foreach (System.Collections.DictionaryEntry item in hastable)
                {                      
                      switch (item.Key.ToString())
                      {
                          case "SubjectName":
                              row["SubjectName"] = item.Value;
                              break;
                          case "SubjectType":
                              row["ResearchType"] = Convert.ToInt16(item.Value) == 2 ? "Company" : "Individual";                              
                              break;
                          case "SubjectCountryDesc":
                              row["CountryName"] = item.Value;
                              break;
                      }
                }
                table.Rows.Add(row);
            }
            DataView view = new DataView(table);
            return view;
        }
        /// <summary>
        /// Created new method to convert list to datatable
        /// </summary>
        /// <param name="arrList"></param>
        /// <returns></returns>
        public System.Data.DataTable ConvertArrayListToDatatable(System.Collections.ArrayList arrList)
        {
            DataTable table = new DataTable("table");
           
            DataRow row;
            foreach (System.Collections.Hashtable hastable in arrList)
            {
                
                foreach (System.Collections.DictionaryEntry item in hastable)
                {
                    table.Columns.Add(item.Key.ToString());
                }
                break;
            }

            foreach (System.Collections.Hashtable hastable in arrList)
            {
                row = table.NewRow();
                foreach (System.Collections.DictionaryEntry item in hastable)
                {
                    //table.Columns.Add(item.Key.ToString());
                    row[item.Key.ToString()] = item.Value;
                }
                table.Rows.Add(row);
            }
           
            return table;
        }

        public DataSet GetBudgetDetailsByCountry(string strClientCode, string strReportType, int intCalculation, string strSubjectCountry, string strSubjectType, bool blnIsExpress)
        {

            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsBudDetCal = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBudgetDetailsByCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Calculation", SqlDbType.Int);
            myCmd.Parameters.Add("@SubjectCountry", SqlDbType.VarChar, 4);
            myCmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 10);
            myCmd.Parameters.Add("@IsExpress", SqlDbType.Bit);

            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@Calculation"].Value = intCalculation;
            myCmd.Parameters["@SubjectCountry"].Value = strSubjectCountry;
            myCmd.Parameters["@SubjectType"].Value = strSubjectType;
            myCmd.Parameters["@IsExpress"].Value = blnIsExpress;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.sp_GetBudgetDetailsByCountry";
            dsBudDetCal = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return dsBudDetCal;
        }
        public DataSet GetFinalBudgetDetails(string strClientCode, Guid strOrderID)
        {

            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsBudDetCal = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetFinalBudgetDetails";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ID", SqlDbType.UniqueIdentifier);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ID"].Value = strOrderID;
            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.sp_GetFinalBudgetDetails";
            dsBudDetCal = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return dsBudDetCal;
        }
        public DataSet GetflsBudgetDetailsByCountry(string strClientCode, string strReportType,string strSubreportType, int intCalculation, string strSubjectCountry, string strSubjectType, bool blnIsExpress)
        {

            SqlCommand sqlCmd = new SqlCommand();
            DataSet dsBudDetCal = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetflsBudgetDetailsByCountry";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters.Add("@ReportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@SubreportType", SqlDbType.VarChar, 15);
            myCmd.Parameters.Add("@Calculation", SqlDbType.Int);
            myCmd.Parameters.Add("@SubjectCountry", SqlDbType.VarChar, 4);
            myCmd.Parameters.Add("@SubjectType", SqlDbType.VarChar, 10);
            myCmd.Parameters.Add("@IsExpress", SqlDbType.Bit);

            myCmd.Parameters["@ClientCode"].Value = strClientCode;
            myCmd.Parameters["@ReportType"].Value = strReportType;
            myCmd.Parameters["@SubreportType"].Value = strSubreportType;
            myCmd.Parameters["@Calculation"].Value = intCalculation;
            myCmd.Parameters["@SubjectCountry"].Value = strSubjectCountry;
            myCmd.Parameters["@SubjectType"].Value = strSubjectType;
            myCmd.Parameters["@IsExpress"].Value = blnIsExpress;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.sp_GetflsBudgetDetailsByCountry";
            dsBudDetCal = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return dsBudDetCal;
        }
        public string GetCountryName(string strCountryCode)
        {
            string strCountryName;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetCountryName";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@CountryCode", SqlDbType.VarChar, 4);
            myCmd.Parameters["@CountryCode"].Value = strCountryCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetCountryName";
            strCountryName = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strCountryName;
        }
        public string GetBulkOrderFilename(string strVersion)
        {
            string strFileName;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkOrderFilename";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Version", SqlDbType.VarChar, 50);
            myCmd.Parameters["@Version"].Value = strVersion;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetBulkOrderFilename";
            strFileName = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strFileName;
        }
        /* End BI 34 Budget & TAT recalculation */

        /* End BI 38 Check Duplicate Subject Name */
        public string RemoveMultipleSpaces(string strInputText)
        {
            return System.Text.RegularExpressions.Regex.Replace(strInputText.Trim(), "\\s+", " ");
        }

        /* End BI 38 Check Duplicate Subject Name */

        /*
         * Code added for enhancement on Download Template
         * Code added by Deepak
         * Code added On 8 sep 2016
         */
       

        public string GetBulkOrderClientFilename(string strVersion,string ClientCode)
        {
            string strFileName;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetBulkOrderClientFilename";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Version", SqlDbType.VarChar, 50);
            myCmd.Parameters["@Version"].Value = strVersion;
            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 10);
            myCmd.Parameters["@ClientCode"].Value = ClientCode;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetBulkOrderClientFilename";
            strFileName = conn.cmdScalarStoredProc(myCmd);
            conn.Close();
            myCmd.Dispose();
            return strFileName;
        }



        public Boolean IsConfigurationExist(string strCode, string strCRN,string strFlag,string strSubrepcode,int entity)
        {
            Boolean blnStatus;

            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_IsConfigurationExist";
            myCmd.CommandType = CommandType.StoredProcedure;
            myCmd.Parameters.Add("@Code", SqlDbType.VarChar, 15);
            myCmd.Parameters["@Code"].Value = strCode;
            myCmd.Parameters.Add("@FLAG", SqlDbType.VarChar, 20);
            myCmd.Parameters["@FLAG"].Value = strFlag;
            myCmd.Parameters.Add("@CRN", SqlDbType.VarChar, 20);
            myCmd.Parameters["@CRN"].Value = strCRN;
            myCmd.Parameters.Add("@SUBREPORTTYPECODE", SqlDbType.VarChar, 20);
            myCmd.Parameters["@SUBREPORTTYPECODE"].Value = strSubrepcode;
            myCmd.Parameters.Add("@Entity", SqlDbType.Int);
            myCmd.Parameters["@Entity"].Value = entity;
            conn.Open();
            conn.callingMethod = "ISBL.General.IsConfigurationExist";
            blnStatus = Convert.ToBoolean(conn.cmdScalarStoredProc(myCmd));
            conn.Close();
            myCmd.Dispose();
            return blnStatus;
        }
        /*Code ended by Deepak*/
       
#endregion
#region Private Members
        private String  RandomizeCharacters(int intMaxLengthPassword)
        {
            char[] allowedCharsCap = "NAYBCRMDGPHIKOEQRSTJCUVWXFZL".ToCharArray();
            char[] allowedCharsSml = "naybcrmdgphikoeqrstjcuvwxfzl".ToCharArray();
            char[] allowedCharsNum = "123@4567!890#".ToCharArray();

            Random myRandom = new Random();
            System.Text.StringBuilder myStringBuilder = new System.Text.StringBuilder();

            //Capital 2 Letter
            for (int index = 0, roof = 2; index < roof; index++)
            {
                myStringBuilder.Append(allowedCharsCap[myRandom.Next(0, allowedCharsCap.Length - 1)]);
            }

            //Number and Special 3 Letter
            for (int index = 0, roof = 3; index < roof; index++)
            {
                myStringBuilder.Append(allowedCharsNum[myRandom.Next(0, allowedCharsNum.Length - 1)]);
            }
    
            //Small 2 Letter
            for (int index = 0, roof = 2; index < roof; index++)
            {
                myStringBuilder.Append(allowedCharsSml[myRandom.Next(0, allowedCharsSml.Length - 1)]);
            }


            // Now myStringBuilder contains eight random characters picked
            // from the allowedChars array.
            return myStringBuilder.ToString();
        }

        private void errTrack(Exception ex)
        {
            string[] computer = WindowsIdentity.GetCurrent().Name.ToString().Split('\\');

            if (!EventLog.SourceExists("OCRS_ISBL_General"))
            {
                EventLog.CreateEventSource("OCRS_ISBL_General", "OCRS General Error");
            }

            EventLog evntLog = new EventLog();
            evntLog.Source = "OCRS_ISBL_General";
            evntLog.WriteEntry("Error 1006\n\n" + ex.Message, EventLogEntryType.Error);
            evntLog.Close();
        }

        //Code Added By Deepak
        public DataSet GetTrackOrderSubReportTypeBYRT(string strClientCode,string ReportType)
        {
            DataSet ds = new DataSet();
            SqlCommand myCmd = new SqlCommand();
            myCmd.Connection = conn.Connection;
            myCmd.CommandText = "sp_GetTrackOrderSRTByRepType";
            myCmd.CommandType = CommandType.StoredProcedure;

            myCmd.Parameters.Add("@ClientCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ClientCode"].Value = strClientCode;

            myCmd.Parameters.Add("@ReportTypeCode", SqlDbType.VarChar, 30);
            myCmd.Parameters["@ReportTypeCode"].Value = ReportType;

            SqlDataAdapter sda = new SqlDataAdapter();
            sda.SelectCommand = myCmd;
            conn.Open();
            conn.callingMethod = "ISBL.General.GetTrackOrderSubReportTypeBYRT";
            ds = conn.FillDataSet(sda);
            conn.Close();
            myCmd.Dispose();
            sda.Dispose();
            return ds;
        }

        //Code Ended By Deepak

#endregion

    }
}
