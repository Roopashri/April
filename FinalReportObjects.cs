using System;
using System.Collections.Generic;
using System.Text;

/*Mercury Enhancement - Mar 2014 for supporting multiple final report downloads*/
namespace ISBL
{
    public class FinalReportObjects
    {
        private string _FinalReportFileName;
        private string _FinalReportBase64StringContent;

        public string FinalReportFileName 
        {
            get
            {
                return _FinalReportFileName;
            }
            set
            {
                _FinalReportFileName = value;
            }
        }

        public string FinalReportBase64StringContent 
        {
            get
            {
                return _FinalReportBase64StringContent;
            }
            set
            {
                _FinalReportBase64StringContent = value;
            }
        }
    }
}
