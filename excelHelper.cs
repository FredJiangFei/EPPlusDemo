

using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

class QuestionnaireExcelTool
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

    public static void CreateCorrectiveActionSheet(ExcelPackage package, List<QuestionDto> questions, string setupTitle)
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

        int row = worksheet.GetNewRowIdx();
        var table = worksheet.Cells[row, 1, row + questions.Count, headers.Count];
        worksheet.Tables.Add(table, "CorrectiveActionTable");

        int col = 1;
        foreach (var header in headers)
        {
            worksheet.SetTableHeader(header, col);
            col++;
        }

        foreach (var question in questions)
        {
            worksheet.NewRow();
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

            int valCol = 1;
            foreach (var item in values)
            {
                worksheet.SetTableRow(item, valCol, 1);
                valCol++;
            }

        }

        worksheet.Cells.AutoFitColumns();
        worksheet.Column(2).Width = 20;
        worksheet.Column(6).Width = 50;
        worksheet.Column(7).Width = 50;

        var title = setupTitle + " - Applicable Regulations - Corrective Action Report";
        worksheet.SetTopTitle(title).AddTopImages(2, -50);
    }

    public static void CreateResultSheet(ExcelPackage package, IEnumerable<AuditingQuestionnaireResultDto> results, string setupTitle)
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

            int col = 1;
            foreach (var header in headers)
            {
                worksheet.SetTableHeader(header, col);
                col++;
            }

            int tableRow = 1;
            foreach (var result in results)
            {
                worksheet.NewRow()
                .SetTableRow(result.Complete.ToString("0.00%"), 1, tableRow)
                .SetResultDashboardCellBackgroundColor(result.Complete);

                worksheet.SetTableRow(result.ScoreSheet, 2, tableRow);

                worksheet
                .SetTableRow(result.Compliance.ToString("0.00%"), 3, tableRow)
                .SetResultDashboardCellBackgroundColor(result.Compliance);

                worksheet
                .SetTableRow(result.Score.ToString(), 4, tableRow)
                .SetResultDashboardCellBackgroundColor(result.Score);


                worksheet.SetTableRow(result.MaxScore.ToString(), 5, tableRow);

                var dynamicValues = result.AuditingRatingCounts.Select(x => x.Value.ToString());
                int valCol = 6;
                foreach (var item in dynamicValues)
                {
                    worksheet.SetTableRow(item, valCol, tableRow);
                    valCol++;
                }

                tableRow++;
            }
        }

        worksheet.Cells.AutoFitColumns();

        var title = setupTitle + " - Applicable Regulations - Results Dashboard";
        worksheet.SetTopTitle(title).AddTopImages(2, -30);
    }

    public static void CreateSetupSheet(ExcelPackage package, AuditingQuestionnaireSetupDto setup)
    {
        var worksheet = package.Workbook.Worksheets.Add("Setup");
        var mergedCells = new List<ExcelRange>();
        var startDate = setup.InspectionStartDate.ToShortDateString();
        var endDate = setup.InspectionEndDate.ToShortDateString();

        worksheet
        .SetFromValue("Facility", setup.Facility).NewRow(mergedCells)
        .SetFromValue("Site", setup.Site).NewRow(mergedCells)
        .SetFromValue("Department", setup.Department).NewRow(mergedCells)
        .SetFromValue("CustomField", setup.CustomField).NewRow(mergedCells)
        .SetFromValue("Site Manager", setup.SiteManager, "Site Manager Title", setup.SiteManagerTitle)
        .SetFromValue("Audit Manager", setup.AuditManager, "Audit Manager Title", setup.AuditManagerTitle)
        .SetFromValue("Managers", setup.Managers, "Managers Title", setup.ManagersTitle)
        .NewRow(mergedCells)
        .SetTitle("Inspection", mergedCells).NewRow(mergedCells)
        .SetFromValue("Start Date", startDate, "End Date", endDate).NewRow(mergedCells)
        .SetFromValue("Lead Inspector", setup.LeadInspector, "Lead Inspector Title", setup.LeadInspectorTitle)
        .SetFromValue("Site Inspector1", setup.SiteInspector1, "Site Inspector1 Title", setup.SiteInspector1Title)
        .SetFromValue("Site Inspector2", setup.SiteInspector2, "Site Inspector2 Title", setup.SiteInspector2Title)
        .SetFromValue(
        "Other Site Inspectors", setup.OtherSiteInspectors,
        "Other Site Inspectors Title", setup.OtherSiteInspectorsTitle)
        .NewRow(mergedCells)
        .SetFromValue("Notes", setup.Notes);

        worksheet.Row(worksheet.Dimension.Rows).Height = 100;

        worksheet.MergeCellsToMatchMaxColumn(mergedCells);

        worksheet.Cells.AutoFitColumns();
        worksheet.Column(2).Width = 40;
        worksheet.Column(5).Width = 40;

        worksheet.SetTopTitle(setup.Title + " - Audit").AddTopImages(1, 20);
    }
}

static class ExcelHelper
{
    public static ExcelWorksheet SetTopTitle(this ExcelWorksheet worksheet, string value)
    {
        worksheet.InsertRowToTop();
        var cell = worksheet.Cells[1, 1];
        cell.Value = value;
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 16;
        return worksheet;
    }


    public static void InsertRowToTop(this ExcelWorksheet worksheet)
    {
        worksheet.InsertRow(1, 1);
        var cell = worksheet.Cells[1, 1];
        cell.Value = "";
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        worksheet.Cells[1, 1, 1, worksheet.Dimension.End.Column].Merge = true;
    }


    public static ExcelWorksheet SetTitle(this ExcelWorksheet worksheet, string value, List<ExcelRange> mergedCells)
    {
        int row = worksheet.GetNewRowIdx();
        var cell = worksheet.Cells[row, 1];
        cell.Value = value;
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = 11;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        mergedCells.Add(cell);

        return worksheet;
    }

    public static int GetNewRowIdx(this ExcelWorksheet worksheet)
    {
        int row = worksheet.Dimension?.Rows + 1 ?? 1;
        return row;
    }

    public static int GetCurrentRowIdx(this ExcelWorksheet worksheet)
    {
        int row = worksheet.Dimension?.Rows ?? 1;
        return row;
    }

    public static void AddTopImages(this ExcelWorksheet worksheet, int lastCol, int lastColOffset)
    {
        worksheet.InsertRowToTop();
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

    public static void SetTableHeader(this ExcelWorksheet worksheet, string value, int col)
    {
        int row = worksheet.GetCurrentRowIdx();
        var cell = worksheet.Cells[row, col];
        cell.Value = value;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Font.Color.SetColor(Color.White);
        cell.Style.Font.Bold = true;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        cell.SetTableHeaderBackgroundColor();
    }

    public static ExcelRange SetTableRow(this ExcelWorksheet worksheet, string value, int col, int tableRow)
    {
        int row = worksheet.GetCurrentRowIdx();
        var cell = worksheet.Cells[row, col];
        cell.SetWhiteBackgroundColor();
        cell.Value = value;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Thin);
        cell.Style.WrapText = true;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

        if (tableRow % 2 == 0)
        {
            cell.SetTableRowSeparateBackgroundColor();
        }
        return cell;
    }

    public static ExcelWorksheet SetFromValue(this ExcelWorksheet worksheet, params string[] values)
    {
        int row = worksheet.GetNewRowIdx();

        SetSetupLabel(worksheet.Cells[row, 1], values[0]);
        SetSetupValue(worksheet.Cells[row, 2], values[1]);

        if (values.Length == 2)
        {
            worksheet.Cells[row, 2, row, 5].Merge = true;
            worksheet.Cells[row, 2, row, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
        }
        else
        {
            SetSetupLabel(worksheet.Cells[row, 4], values[2]);
            SetSetupValue(worksheet.Cells[row, 5], values[3]);
        }

        return worksheet;
    }

    public static void SetSetupValue(ExcelRange cell, string value)
    {
        cell.Value = value;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
        cell.Style.WrapText = true;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

        if (string.IsNullOrEmpty(value))
        {
            cell.SetSetupNoValueBackgroundColor();
        }
    }

    public static void SetSetupLabel(ExcelRange cell, string name)
    {
        cell.Value = name + ":";
        cell.Style.Font.Bold = true;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
    }

    public static ExcelWorksheet NewRow(this ExcelWorksheet worksheet)
    {
        int row = worksheet.GetNewRowIdx();
        var cell = worksheet.Cells[row, 1];
        cell.Value = "";
        return worksheet;
    }

    public static ExcelWorksheet NewRow(this ExcelWorksheet worksheet, List<ExcelRange> mergedCells)
    {
        int row = worksheet.GetNewRowIdx();
        var cell = worksheet.Cells[row, 1];
        cell.Value = "";
        mergedCells.Add(cell);

        return worksheet;
    }

    public static void MergeCellsToMatchMaxColumn(this ExcelWorksheet worksheet, List<ExcelRange> mergedCells)
    {
        mergedCells.ForEach(cell =>
        {
            worksheet.Cells[cell.Start.Row, 1, cell.Start.Row, worksheet.Dimension.End.Column].Merge = true;
        });
    }

    public static void SetBackgroundColor(this ExcelRange cell, string hex)
    {
        Color colFromHex = ColorTranslator.FromHtml(hex);
        cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
        cell.Style.Fill.BackgroundColor.SetColor(colFromHex);
    }

    public static void SetSetupNoValueBackgroundColor(this ExcelRange cell)
    {
        cell.SetBackgroundColor("#DCE6F0");
    }

    public static void SetWhiteBackgroundColor(this ExcelRange cell)
    {
        cell.SetBackgroundColor("#fff");
    }

    public static void SetTableHeaderBackgroundColor(this ExcelRange cell)
    {
        cell.SetBackgroundColor("#343896");
    }

    public static void SetTableRowSeparateBackgroundColor(this ExcelRange cell)
    {
        cell.SetBackgroundColor("#D3D3D3");
    }

    public static void SetResultDashboardCellBackgroundColor(this ExcelRange cell, decimal value)
    {
        if (value == 0.00m)
        {
            cell.SetBackgroundColor("#FAC000");
            return;
        }
        
        cell.SetBackgroundColor("#92D050");
    }
}