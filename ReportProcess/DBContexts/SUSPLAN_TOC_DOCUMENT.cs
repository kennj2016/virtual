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
    
    public partial class SUSPLAN_TOC_DOCUMENT
    {
        public int TOCDOC_ID { get; set; }
        public Nullable<int> SECTION_ID { get; set; }
        public Nullable<int> PREVIOUS_TOCDOC_ID { get; set; }
        public Nullable<int> NEXT_TOCDOC_ID { get; set; }
        public string DOCUMENT_LABEL { get; set; }
        public Nullable<int> DOC_OPTION { get; set; }
    }
}