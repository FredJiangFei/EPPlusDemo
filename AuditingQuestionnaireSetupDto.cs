using System;

public class AuditingQuestionnaireSetupDto
    {
        public long AuditingQuestionnaireId { get; set; }
        public string Title { get; set; }
        public string Facility { get; set; }
        public string Site { get; set; }
        public string Department { get; set; }
        public string CustomField { get; set; }
        public DateTime InspectionStartDate { get; set; }
        public DateTime InspectionEndDate { get; set; }
        public string Notes { get; set; }
        public long Id { get; set; }
        public string SiteManager { get; set; }
        public string SiteManagerTitle { get; set; }
        public string AuditManager { get; set; }
        public string AuditManagerTitle { get; set; }
        public string Managers { get; set; }
        public string ManagersTitle { get; set; }
        public string LeadInspector { get; set; }
        public string LeadInspectorTitle { get; set; }
        public string SiteInspector1 { get; set; }
        public string SiteInspector1Title { get; set; }
        public string SiteInspector2 { get; set; }
        public string SiteInspector2Title { get; set; }
        public string OtherSiteInspectors { get; set; }
        public string OtherSiteInspectorsTitle { get; set; }
    }