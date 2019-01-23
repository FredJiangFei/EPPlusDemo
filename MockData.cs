
using System;
using System.Collections.Generic;

public class MockData
{
    public static AuditingQuestionnaireSetupDto GetSetup()
    {
        var setup = new AuditingQuestionnaireSetupDto
        {
            Title = "SHE Test -  - Audit",
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
            SiteInspector2Title = "liu",
            OtherSiteInspectors = "randy",
            OtherSiteInspectorsTitle = "gong",
            Notes = @"To see how this works let’s do a short walkthrough of sample 6 that creates a report on a directory in the file system. 
                The spreadsheet is created without any template. 
                First sheet is a list of subdirectories and files, with an icon, name, size, and dates. The second sheet contains some statistics"
        };

        return setup;
    }

    public static List<AuditingQuestionnaireResultDto> GetResult()
    {
        var results = new List<AuditingQuestionnaireResultDto> {
            new AuditingQuestionnaireResultDto
            {
                Id = 2,
                ScoreSheet = "Rulebook 1",
                Complete = 0.33m,
                Compliance = 0.11m,
                Score = 15,
                MaxScore = 9,
                AuditingRatingCounts = new List<AuditingRatingCountDto> {
                    new AuditingRatingCountDto { Key = "Observations", Value = 0 },
                    new AuditingRatingCountDto { Key = "Compliance", Value = 1 },
                    new AuditingRatingCountDto { Key = "Needs Attention", Value = 0 },
                    new AuditingRatingCountDto { Key = "Needs Improve", Value = 7 },
                    new AuditingRatingCountDto { Key = "Partial compliance", Value = 2 }
                }
            },
             new AuditingQuestionnaireResultDto
            {
                Id = 3,
                ScoreSheet = "Rulebook 1",
                Complete = 0.73m,
                Compliance = 0.91m,
                Score = 54,
                MaxScore = 2,
                AuditingRatingCounts = new List<AuditingRatingCountDto> {
                    new AuditingRatingCountDto { Key = "Observations", Value = 0 },
                    new AuditingRatingCountDto { Key = "Compliance", Value = 1 },
                    new AuditingRatingCountDto { Key = "Needs Attention", Value = 0 },
                    new AuditingRatingCountDto { Key = "Needs Improve", Value = 7 },
                    new AuditingRatingCountDto { Key = "Partial compliance", Value = 2 }
                }
            },
             new AuditingQuestionnaireResultDto
            {
                Id = 4,
                ScoreSheet = "Rulebook 1",
                Complete = 0.33m,
                Compliance = 0.11m,
                Score = 15,
                MaxScore = 9,
                AuditingRatingCounts = new List<AuditingRatingCountDto> {
                    new AuditingRatingCountDto { Key = "Observations", Value = 0 },
                    new AuditingRatingCountDto { Key = "Compliance", Value = 1 },
                    new AuditingRatingCountDto { Key = "Needs Attention", Value = 0 },
                    new AuditingRatingCountDto { Key = "Needs Improve", Value = 7 },
                    new AuditingRatingCountDto { Key = "Partial compliance", Value = 2 }
                }
            }
        };


        return results;
    }
}