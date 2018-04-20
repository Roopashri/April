using System;
using System.Collections.Generic;
using System.Text;
using System.Collections;

namespace ISBL
{
    public class SearchCriteria
    {
        private string _CRN;
        private string _Status;
        private string _ReferenceNo;
        private string _ReportType;
        private string _SubReportType;
        private string _SubjectName;
        private string _Country;
        private ArrayList _Countries;   // Added By Stev 08Oct08 
        private string _StartDate;
        private string _EndDate;
        private string _FSDDate;
        private Boolean _ShowAll;       // Added By Stev, 11Oct08
        private Boolean _NotShowCNL;    // Added By Stev, 11Oct08
        private string _OrderedBy;
        private string _BulkOrderID;
        private System.Collections.ArrayList _CustomFieldsArrayList; //ISIS Custom Field Enhancement - Feb 2013 Mitul

        private Boolean _HasRisk; //Added by Nagaraj 28 Oct 2014 for ISIS Risk Enhancement

        public string CRN
        {
            get
            {
                return _CRN;
            }
            set
            {
                _CRN = value;
            }
        }
        public string Status
        {
            get
            {
                return _Status;
            }
            set
            {
                _Status = value;
            }
        }
        public string ReferenceNo
        {
            get
            {
                return _ReferenceNo;
            }
            set
            {
                _ReferenceNo = value;
            }
        }
        public string ReportType
        {
            get
            {
                return _ReportType;
            }
            set
            {
                _ReportType = value;
            }
        }
        public string SubReportType
        {
            get
            {
                return _SubReportType;
            }
            set
            {
                _SubReportType = value;
            }
        }
        public string SubjectName
        {
            get
            {
                return _SubjectName;
            }
            set
            {
                _SubjectName = value;
            }
        }
        public string Country
        {
            get
            {
                return _Country;
            }
            set
            {
                _Country = value;
            }
        }
        // Added By Stev, 09Oct08. To keep Country List
        public ArrayList Countries
        {
            get
            {
                return _Countries;
            }
            set
            {
                _Countries = value;
            }
        }
        // Added By Stev, 10Oct08. To Keep chkShowAll in TrackOrder.aspx
        public Boolean bShowAll
        {
            get
            {
                return _ShowAll;
            }
            set
            {
                _ShowAll = value;
            }
        }
        // Added By Stev, 10Oct08. To Keep chkNotCancel in TrackOrder.aspx
        public Boolean bNotShowCancel
        {
            get
            {
                return _NotShowCNL;
            }
            set
            {
                _NotShowCNL = value;
            }
        }

        public string StartDate
        {
            get
            {
                return _StartDate;
            }
            set
            {
                _StartDate = value;
            }
        }
        public string EndDate
        {
            get
            {
                return _EndDate;
            }
            set
            {
                _EndDate = value;
            }
        }
        public string FSDDate
        {
            get
            {
                return _FSDDate;
            }
            set
            {
                _FSDDate = value;
            }
        }
        /** Start ISIS v2 Phase 1 Release 1 **/
        public string OrderedBy
        {
            get
            {
                return _OrderedBy;
            }
            set
            {
                _OrderedBy = value;
            }
        }
        public string BulkOrderID
        {
            get
            {
                return _BulkOrderID;
            }
            set
            {
                _BulkOrderID = value;
            }
        }
        /** End ISIS v2 Phase 1 Release 1 **/

        /** Start ISIS Custom Field Enhancement - Feb 2013 Mitul **/
        public System.Collections.ArrayList CustomFieldsArrayList
        {
            get
            {
                return _CustomFieldsArrayList;
            }
            set
            {
                _CustomFieldsArrayList = value;
            }
        }
        /** End ISIS Custom Field Enhancement - Feb 2013 Mitul **/


        /** Start ISIS Custom Field Enhancement - OCT 2014 Nagaraj **/
        public Boolean HasRisk
        {
            get
            {
                return _HasRisk;
            }
            set
            {
                _HasRisk = value;
            }
        }
        /** END ISIS Custom Field Enhancement - OCT 2014 Nagaraj **/
    }
}
