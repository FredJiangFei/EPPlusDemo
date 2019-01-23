using System;

namespace epplus_demo
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");

            var setup = new AuditingQuestionnaireSetupDto
            {
                Title ="SHE Test -  - Audit",
                Facility = "Fac 567",
                Site = "North Wacker Drive, Peoria, IL, USA",
                Department = "Department 1",
                CustomField = "CustomField fred add",
                SiteManager = "fred",
                SiteManagerTitle = "manager",
                AuditManager = "yan",
                AuditManagerTitle = "amt",
                InspectionStartDate = DateTime.Now,
                InspectionEndDate = DateTime.Now.AddDays(1),
                LeadInspector = "LeadInspector fred",
                LeadInspectorTitle = "lit",
                SiteInspector1 = "jack",
                SiteInspector1Title = "wang",
                // SiteInspector2 ="jia",
                SiteInspector2Title ="liu",
                OtherSiteInspectors ="randy",
                OtherSiteInspectorsTitle ="gong",
                Notes =@"To see how this works let’s do a short walkthrough of sample 6 that creates a report on a directory in the file system. 
                The spreadsheet is created without any template. 
                First sheet is a list of subdirectories and files, with an icon, name, size, and dates. The second sheet contains some statistics"
            };
            ExcelHelper.Create(setup);
        }
    }
}
