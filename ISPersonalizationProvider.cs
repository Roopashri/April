using System;
using System.Collections.Generic;
using System.Text;
using System.Web.UI.WebControls.WebParts;
using System.Web;

namespace ISBL
{
    public class ISPersonalizationProvider :
        SqlPersonalizationProvider
    {

#region Private Variables

        private string _userName = "";
        private string _path = "";

        private string UserName         
        {         
            get         
            {
                return _userName;         
            }         
            set         
            {
                _userName = value;         
            }         
        }

        private string Path
        {
            get
            {
                return _path;
            }
            set
            {
                _path = value;
            }
        }
#endregion

#region Constructors

        public ISPersonalizationProvider()
        {
        }
#endregion

#region Private Member

        protected override void LoadPersonalizationBlobs(WebPartManager webPartManager, string path, string userName, ref byte[] sharedDataBlob, ref byte[] userDataBlob)
        {
            ISBL.User oUser = (ISBL.User)System.Web.HttpContext.Current.Session["Admin"];

            path = oUser.CurrentPath;
            userName = oUser.LoginID;
            base.LoadPersonalizationBlobs(webPartManager, path, userName, ref sharedDataBlob, ref userDataBlob);
            oUser = null;
        }


        protected override void SavePersonalizationBlob(WebPartManager webPartManager, string path, string userName, byte[] dataBlob)
        {
            ISBL.User oUser = (ISBL.User)System.Web.HttpContext.Current.Session["Admin"];

            path = oUser.CurrentPath;
            userName = oUser.LoginID;
            base.SavePersonalizationBlob(webPartManager, path, userName, dataBlob);
            oUser = null;
        }

        protected override void ResetPersonalizationBlob(WebPartManager webPartManager, string path, string userName) 
        {
            ISBL.User oUser = (ISBL.User)System.Web.HttpContext.Current.Session["Admin"];

            path = oUser.CurrentPath;
            userName = oUser.LoginID;
            base.ResetPersonalizationBlob(webPartManager, path, userName);
            oUser = null;
        }

#endregion

    }
}