

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

class ExcelHelper
{

    public static void Create()
    {
        using (var package = new ExcelPackage())
        {
            package.Workbook.Properties.Created = DateTime.Now;

            var setup = MockData.GetSetup();
            var questions = MockData.GetQuestions();
            var results = MockData.GetResults();

            CreateSetupSheet(package, setup);
            CreateResultSheet(package, results, setup.Title);
            CreateCorrectiveActionSheet(package, questions, setup.Title);

            var fi = new FileInfo(@"F:\File.xlsx");
            package.SaveAs(fi);
        }
    }

    private static void CreateCorrectiveActionSheet(ExcelPackage package, List<QuestionDto> questions, string setupTitle)
        {
            var worksheet = package.Workbook.Worksheets.Add("Corrective_Action");
            var headers = new List<string>() {
                "Number",
                "Section",
                "Rank",
                "Status",
                "Score",
                "Observations",
                "Recommendations",
                "Person Assigned",
                "Start Date",
                "Date Complete"
            };

            int row = GetNewRow(worksheet);
            var table = worksheet.Cells[row, 1, row + questions.Count, headers.Count];
            worksheet.Tables.Add(table, "CorrectiveActionTable");

            SetTableHeader(worksheet, headers);

            foreach (var question in questions)
            {
                var values = new List<string>() {
                    "1",
                    "Section",
                    question.RankRating?.ToString(),
                    "Status",
                    question.AuditRating?.ToString(),
                    question.Observations,
                    question.Recommendations,
                    question.AssignAnaswerUserName,
                    DateTime.Now.ToShortDateString(),
                    DateTime.Now.ToShortDateString()
                };
                SetTableRowValue(worksheet, values, 1);
            }

            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 20;
            worksheet.Column(6).Width = 50;
            worksheet.Column(7).Width = 50;

            var title = setupTitle + " - Applicable Regulations - Corrective Action Report";
            SetTopTitle(worksheet, title);

            AddTopImages(worksheet, 2, -50);
        }

        private static void CreateResultSheet(ExcelPackage package, IEnumerable<AuditingQuestionnaireResultDto> results, string setupTitle)
        {
                var worksheet = package.Workbook.Worksheets.Add("Results_Dashboard");

                var firstResult = results.FirstOrDefault();
                if (firstResult != null)
                {
                    var headers = new List<string>() {
                    "% Completed", "Scoresheet", "% Compliance", "Score", "Max Score"
                };
                var dynamicHeaders = firstResult.AuditingRatingCounts.Select(x => x.Key);
                headers.AddRange(dynamicHeaders);
                SetTableHeader(worksheet, headers);

                int tableRow = 1;
                foreach (var result in results)
                {
                    var values = new List<string>() {
                    result.Complete.ToString("0.00%"),
                    result.ScoreSheet,
                    result.Compliance.ToString("0.00%"),
                    result.Score.ToString(),
                    result.MaxScore.ToString()
                };
                    var dynamicValues = result.AuditingRatingCounts.Select(x => x.Value.ToString());
                    values.AddRange(dynamicValues);

                    SetTableRowValue(worksheet, values, tableRow);
                    tableRow++;
                }
            }

            worksheet.Cells.AutoFitColumns();

            var title = setupTitle + " - Applicable Regulations - Results Dashboard";
            SetTopTitle(worksheet, title);

            AddTopImages(worksheet, 2, -30);
        }

        private static void SetTableHeader(ExcelWorksheet worksheet, List<string> headers)
        {
            int row = GetNewRow(worksheet);
            int col = 1;
            foreach (var item in headers)
            {
                var cell = worksheet.Cells[row, col];
                cell.Value = item;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.Font.Color.SetColor(Color.White);
                cell.Style.Font.Bold = true;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                SetBackgroundColor(cell, "#343896");
                col++;
            }
        }

        private static void SetTableRowValue(ExcelWorksheet worksheet, List<string> values, int tableRow)
        {
            int row = GetNewRow(worksheet);
            int col = 1;
            foreach (var value in values)
            {
                var cell = worksheet.Cells[row, col];
                SetBackgroundColor(cell, "#fff");
                cell.Value = value;
                cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
                cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
                cell.Style.WrapText = true;
                cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

                if (tableRow % 2 == 0)
                {
                    SetBackgroundColor(cell, "#D3D3D3");
                }
                col++;
            }
        }

        private static void CreateSetupSheet(ExcelPackage package, AuditingQuestionnaireSetupDto setup)
        {
            var worksheet = package.Workbook.Worksheets.Add("Setup");
            var mergedCells = new List<ExcelRange>();

            SetSetupCellValue(worksheet, "Facility", setup.Facility);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet, "Site", setup.Site);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet, "Department", setup.Department);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet, "CustomField", setup.CustomField);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet,
            "Site Manager", setup.SiteManager,
            "Site Manager Title", setup.SiteManagerTitle);
            SetSetupCellValue(worksheet,
            "Audit Manager", setup.AuditManager,
            "Audit Manager Title", setup.AuditManagerTitle);
            SetSetupCellValue(worksheet,
            "Managers", setup.Managers,
            "Managers Title", setup.ManagersTitle);
            NewRow(worksheet, mergedCells);

            SetTitle(worksheet, "Inspection", mergedCells);
            NewRow(worksheet, mergedCells);

            var startDate = setup.InspectionStartDate.ToShortDateString();
            var endDate = setup.InspectionEndDate.ToShortDateString();
            SetSetupCellValue(worksheet, "Start Date", startDate, "End Date", endDate);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet,
            "Lead Inspector", setup.LeadInspector,
            "Lead Inspector Title", setup.LeadInspectorTitle);
            SetSetupCellValue(worksheet,
            "Site Inspector1", setup.SiteInspector1,
            "Site Inspector1 Title", setup.SiteInspector1Title);
            SetSetupCellValue(worksheet,
            "Site Inspector2", setup.SiteInspector2,
            "Site Inspector2 Title", setup.SiteInspector2Title);
            SetSetupCellValue(worksheet,
            "Other Site Inspectors", setup.OtherSiteInspectors,
            "Other Site Inspectors Title", setup.OtherSiteInspectorsTitle);
            NewRow(worksheet, mergedCells);

            SetSetupCellValue(worksheet, "Notes", setup.Notes);
            worksheet.Row(worksheet.Dimension.Rows).Height = 100;

            MergeCellsToMatchMaxColumn(mergedCells, worksheet);

            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 40;
            worksheet.Column(5).Width = 40;

            SetTopTitle(worksheet, setup.Title + " - Audit");
            AddTopImages(worksheet, 1, 20);
        }

        private static void AddTopImages(ExcelWorksheet worksheet, int lastCol, int lastColOffset)
        {
            InsertRowToTop(worksheet);
            worksheet.Row(1).Height = 50;

            if (File.Exists(@"Xcelerator.png"))
            {
                var xcelerator = Image.FromFile(@"Xcelerator.png");
                var xceleratorPic = worksheet.Drawings.AddPicture("Xcelerator", xcelerator);
                xceleratorPic.SetPosition(0, 5, 0, 0);
            }

            if (File.Exists(@"STP.png"))
            {
                var stp = Image.FromFile(@"STP.png");
                var stpPic = worksheet.Drawings.AddPicture("STP", stp);
                stpPic.SetPosition(0, 5, worksheet.Dimension.End.Column - lastCol, lastColOffset);
            }
        }

        private static void SetSetupCellValue(ExcelWorksheet worksheet, params string[] values)
        {
            int row = GetNewRow(worksheet);

            SetSetupLabel(worksheet.Cells[row, 1], values[0]);
            SetSetupValue(worksheet.Cells[row, 2], values[1]);

            if (values.Length == 2)
            {
                worksheet.Cells[row, 2, row, 5].Merge = true;
                worksheet.Cells[row, 2, row, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
                return;
            }

            SetSetupLabel(worksheet.Cells[row, 4], values[2]);
            SetSetupValue(worksheet.Cells[row, 5], values[3]);
        }

        private static void SetSetupValue(ExcelRange cell, string value)
        {
            cell.Value = value;
            cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
            cell.Style.WrapText = true;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

            if (string.IsNullOrEmpty(value))
            {
                SetBackgroundColor(cell, "#DCE6F0");
            }
        }

        private static void SetBackgroundColor(ExcelRange cell, string hex)
        {
            Color colFromHex = ColorTranslator.FromHtml(hex);
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(colFromHex);
        }

        private static void SetSetupLabel(ExcelRange cell, string name)
        {
            cell.Value = name + ":";
            cell.Style.Font.Bold = true;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
            cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
        }

        private static void NewRow(ExcelWorksheet worksheet, List<ExcelRange> mergedCells)
        {
            int row = GetNewRow(worksheet);
            var cell = worksheet.Cells[row, 1];
            cell.Value = "";
            mergedCells.Add(cell);
        }

        private static void InsertRowToTop(ExcelWorksheet worksheet)
        {
            worksheet.InsertRow(1, 1);
            var cell = worksheet.Cells[1, 1];
            cell.Value = "";
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column].Merge = true;
        }

        private static void SetTopTitle(ExcelWorksheet worksheet, string value)
        {
            InsertRowToTop(worksheet);
            var cell = worksheet.Cells[1, 1];
            cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 16;
        }

        private static void SetTitle(ExcelWorksheet worksheet, string value, List<ExcelRange> mergedCells)
        {
            int row = GetNewRow(worksheet);
            var cell = worksheet.Cells[row, 1];
            cell.Value = value;
            cell.Style.Font.Bold = true;
            cell.Style.Font.Size = 11;
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            mergedCells.Add(cell);
        }

        private static void MergeCellsToMatchMaxColumn(List<ExcelRange> mergedCells, ExcelWorksheet worksheet)
        {
            mergedCells.ForEach(cell =>
            {
                worksheet.Cells[cell.Start.Row, 1, cell.Start.Row, worksheet.Dimension.End.Column].Merge = true;
            });
        }

        private static int GetNewRow(ExcelWorksheet worksheet)
        {
            int row = worksheet.Dimension?.Rows + 1 ?? 1;
            return row;
        }
}
