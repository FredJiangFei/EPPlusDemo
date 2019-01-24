

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
        using (var excelPackage = new ExcelPackage())
        {
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            var setup = MockData.GetSetup();
            CreateSetupSheet(excelPackage, setup);

            var mergedCells = new List<ExcelRange>();
            var results = MockData.GetResult();
            var worksheet = excelPackage.Workbook.Worksheets.Add("Results_Dashboard");

            NewRow(worksheet, mergedCells, 50);
            AddSheetHeaderImages(worksheet);

            var title = setup.Title + " - Applicable Regulations - Results Dashboard";
            SetTitle(worksheet, title, mergedCells, 16);

            NewRow(worksheet, mergedCells);

            SetDashboardTableHeader(worksheet, results);

            


            MergeCellsToMatchMaxColumn(mergedCells, worksheet);
            worksheet.Cells.AutoFitColumns();

            var fi = new FileInfo(@"F:\File.xlsx");
            excelPackage.SaveAs(fi);
        }
    }

    private static void MergeCellsToMatchMaxColumn(List<ExcelRange> mergedCells, ExcelWorksheet worksheet)
    {
        mergedCells.ForEach(cell =>
        {
            worksheet.Cells[cell.Start.Row, 1, cell.Start.Row, worksheet.Dimension.End.Column].Merge = true;
        });
    }

    private static void SetDashboardTableHeader(ExcelWorksheet worksheet, List<AuditingQuestionnaireResultDto> results)
    {
        var result = results.FirstOrDefault();
        if (result == null)
        {
            return;
        }

        var headers = new List<string>() {
            "% Completed", "Scoresheet", "% Compliance", "Score", "Max Score"
        };
        var dynamicHeaders = result.AuditingRatingCounts.Select(x => x.Key);
        headers.AddRange(dynamicHeaders);

        int row = GetNewRow(worksheet);
        int col = 1;
        foreach (var item in headers)
        {
            var cell = worksheet.Cells[row, col];
            cell.Value = headers[col - 1];
            cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            cell.Style.Font.Color.SetColor(Color.White);
            cell.Style.Font.Bold = true;
            SetBackgroundColor(cell, "#343896");
            col++;
        }
    }

    private static void CreateSetupSheet(ExcelPackage package, AuditingQuestionnaireSetupDto setup)
    {
        var worksheet = package.Workbook.Worksheets.Add("Setup");
        var mergedCells = new List<ExcelRange>();

        NewRow(worksheet, mergedCells, 50);
        AddSheetHeaderImages(worksheet);

        SetTitle(worksheet, setup.Title + " - Audit", mergedCells, 16);

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
    }

    private static void AddSheetHeaderImages(ExcelWorksheet worksheet)
    {
        Image img = Image.FromFile(@"Xcelerator.png");
        ExcelPicture pic = worksheet.Drawings.AddPicture("Xcelerator", img);
        pic.SetPosition(0, 5, 0, 0);

        Image img2 = Image.FromFile(@"STP.png");
        ExcelPicture pic2 = worksheet.Drawings.AddPicture("STP", img2);
        pic2.SetPosition(0, 5, 4, 15);
    }

    private static void SetSetupCellValue(ExcelWorksheet worksheet, params string[] values)
    {
        int row = GetNewRow(worksheet);

        var cell1 = worksheet.Cells[row, 1];
        SetSetupLabel(cell1, values[0]);

        var cell2 = worksheet.Cells[row, 2];
        SetSetupValue(cell2, values[1]);

        if (values.Length == 2)
        {
            worksheet.Cells[row, 2, row, 5].Merge = true;
            worksheet.Cells[row, 2, row, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            return;
        }

        var cell3 = worksheet.Cells[row, 4];
        SetSetupLabel(cell3, values[2]);

        var cell4 = worksheet.Cells[row, 5];
        SetSetupValue(cell4, values[3]);
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
        Color colFromHex = System.Drawing.ColorTranslator.FromHtml(hex);
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

    private static void NewRow(ExcelWorksheet worksheet, List<ExcelRange> mergedCells, int? height = null)
    {
        int row = GetNewRow(worksheet);
        var cell = worksheet.Cells[row, 1];
        cell.Value = "";

        if (height != null)
        {
            worksheet.Row(row).Height = 50;
        }
        mergedCells.Add(cell);
    }

    private static void SetTitle(ExcelWorksheet worksheet, string value, List<ExcelRange> mergedCells, int fontSize = 11)
    {
        int row = GetNewRow(worksheet);
        var cell = worksheet.Cells[row, 1];
        cell.Value = value;
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = fontSize;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        mergedCells.Add(cell);
    }

    private static int GetNewRow(ExcelWorksheet worksheet)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        return row;
    }
}
