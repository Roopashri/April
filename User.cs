using System;
using System.Collections.Generic;
using System.Text;

namespace ISBL
{
    public class User
    {
        private string _ClientCode;
        private string _InstanceID;
        private string _LoginID;
        private string _Role;
        private string _UserType;
        private string _LoginIDEmail;
        private string _ClientName;
        private string _EncPassword;
        private string _GOCAccess;
        private Boolean _DashBoardAccess;
        private Boolean _OrderManagementAccess;
        private Boolean _CertCheckAccess;
        private string _LoginName;
        private DateTime _LastLogin;
        private string _LogoPath;
        private string _LogoText;
        private Boolean _AllowBatchProcessing;
        private string _CurrentPath; //Adam Web Parts 18-Mar-09
        private string _ImpersonateLoginID; //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        private Boolean _IsImpersonate;//OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        private Boolean _allowCancelOrder; //Implement Cancel Order Oct 2011
        private int _DayLeftBeforeExpired; //Password Enhancement Oct 2012
        private Boolean _ShowExpiryNotification; //Password Enhancement Oct 2012
        private Object _RegularExpression; //Custom Field Enhancement - Feb 2013 Adam
        private Object _UserGuideList; //MultiLang UserGuide - July 2013 Mitul
        private Object _SubjectOthderDetails; //ISIS and Atlas Data Sync - Aug 2013 Adam
        private Boolean _IsRefreshOrder; //BI 20 ISIS Refresh Order
        private string _RefreshOrderBatchID; //BI 20 ISIS Refresh Order

        public Boolean IsImpersonate //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        {
            get
            {
                return _IsImpersonate;
            }
            set
            {
                _IsImpersonate = value;
            }
        }

        public string ImpersonateLoginID //OCRS Phase 4 6.G v 2.8b Emulate Client user login Enhancement Oct 2009 - Adam
        {
            get
            {
                return _ImpersonateLoginID;
            }
            set
            {
                _ImpersonateLoginID = value;
            }
        }

        public string CurrentPath //Adam Web Parts 18-Mar-09
        {
            get
            {
                return _CurrentPath;
            }
            set
            {
                _CurrentPath = value;
            }
        }

        public string ClientCode
        {
            get
            {
                return _ClientCode;
            }
            set
            {
                _ClientCode = value;
            }
        }

        public string InstanceID
        {
            get
            {
                return _InstanceID;
            }
            set
            {
                _InstanceID = value;
            }
        }

        public string LoginID
        {
            get
            {
                return _LoginID;
            }
            set
            {
                _LoginID = value;
            }
        }
        public string Role
        {
            get
            {
                return _Role;
            }
            set
            {
                _Role = value;
            }
        }
        public string UserType
        {
            get
            {
                return _UserType;
            }
            set
            {
                _UserType = value;
            }
        }
        public string LoginIDEmail
        {
            get
            {
                return _LoginIDEmail;
            }
            set
            {
                _LoginIDEmail = value;
            }
        }
        public string ClientName
        {
            get
            {
                return _ClientName;
            }
            set
            {
                _ClientName = value;
            }
        }
        public string EncPassword
        {
            get
            {
                return _EncPassword;
            }
            set
            {
                _EncPassword = value;
            }
        }

        public string GOCAccess
        {
            get
            {
                return _GOCAccess;
            }
            set
            {
                _GOCAccess = value;
            }
        }

        public Boolean DashBoardAccess
        {
            get
            {
                return _DashBoardAccess;
            }
            set
            {
                _DashBoardAccess = value;
            }
        }

        public Boolean OrderManagementAccess
        {
            get
            {
                return _OrderManagementAccess;
            }
            set
            {
                _OrderManagementAccess = value;
            }
        }

        public string LoginName
        {
            get
            {
                return _LoginName;
            }
            set
            {
                _LoginName = value;
            }
        }

        public DateTime LastLogin
        {
            get
            {
                return _LastLogin;
            }
            set
            {
                _LastLogin = value;
            }
        }

        public string LogoPath
        {
            get
            {
                return _LogoPath;
            }
            set
            {
                _LogoPath = value;
            }
        }

        public string LogoText
        {
            get
            {
                return _LogoText;
            }
            set
            {
                _LogoText = value;
            }
        }

        public Boolean AllowBatchProcessing
        {
            get
            {
                return _AllowBatchProcessing;
            }
            set
            {
                _AllowBatchProcessing = value;
            }
        }

        public Boolean CertCheckAccess
        {
            get
            {
                return _CertCheckAccess;
            }
            set
            {
                _CertCheckAccess = value;
            }
        }

        public Boolean AllowCancelOrder //Implement Cancel Order Oct 2011
        {
            get
            {
                return _allowCancelOrder;
            }
            set
            {
                _allowCancelOrder = value;
            }
        }

        public int DayLeftBeforeExpired //Pasword Enhancement 2012
        {
            get
            {
                return _DayLeftBeforeExpired;
            }
            set
            {
                _DayLeftBeforeExpired = value;
            }
        }

        public Boolean ShowExpiryNotification //Pasword Enhancement 2012
        {
            get
            {
                return _ShowExpiryNotification;
            }
            set
            {
                _ShowExpiryNotification = value;
            }
        }

        public Object RegularExpression //Pasword Enhancement 2012
        {
            get
            {
                return _RegularExpression;
            }
            set
            {
                _RegularExpression = value;
            }
        }

        
        
        public Object UserGuideList //MutliLang UserGuide Enhancement Jily 2013 Mitul
        {
            get
            {
                return _UserGuideList;
            }
            set
            {
                _UserGuideList = value;
            }
        }

        public Object SubjectOthderDetails //ISIS and Atlas Data Sync - Aug 2013 Adam
        {
            get
            {
                return _SubjectOthderDetails;
            }
            set
            {
                _SubjectOthderDetails = value;
            }
        }

        //Start BI 20 ISIS Refresh Order
        public Boolean IsRefreshOrder
        {
            get
            {
                return _IsRefreshOrder;
            }
            set
            {
                _IsRefreshOrder = value;
            }
        }

        public string RefreshOrderBatchID
        {
            get
            {
                return _RefreshOrderBatchID;
            }
            set
            {
                _RefreshOrderBatchID = value;
            }
        }

        //End BI 20 ISIS Refresh Order

    }
}
