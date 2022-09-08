//------------------------------------------------------------------------------
// <auto-generated>
//     This code was generated from a template.
//
//     Manual changes to this file may cause unexpected behavior in your application.
//     Manual changes to this file will be overwritten if the code is regenerated.
// </auto-generated>
//------------------------------------------------------------------------------

namespace ReportProcess.DBContexts
{
    using System;
    using System.Collections.Generic;
    
    public partial class SUSPLAN_REPORT_Q
    {
        public int RQ_ID { get; set; }
        public Nullable<int> TOC_ID { get; set; }
        public Nullable<int> NODE_ID { get; set; }
        public Nullable<int> SURVEY_ID { get; set; }
        public string USER_ID { get; set; }
        public string STATUS { get; set; }
        public string STD_PAGE_NUMBERS { get; set; }
        public Nullable<System.DateTime> STARTED_DATE { get; set; }
        public Nullable<System.DateTime> SUBMITTED_DATE { get; set; }
        public Nullable<System.DateTime> FINISHED_DATE { get; set; }
        public string TARGET_DIRECTORY { get; set; }
        public string TARGET_NAME { get; set; }
        public string RENAME_NAME { get; set; }
        public string CREATE_DOCUMENT_YN { get; set; }
        public string DOCUMENT_LABEL { get; set; }
        public string CONTROL_FILE { get; set; }
        public string ERROR_MESSAGE { get; set; }
        public Nullable<int> FIRST_SEC_ID { get; set; }
    
        public virtual SUSPLAN_TOC_HEADER SUSPLAN_TOC_HEADER { get; set; }
    }
}
