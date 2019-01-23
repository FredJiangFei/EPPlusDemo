

using System;
using System.Drawing;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Drawing;
using OfficeOpenXml.Style;

class ExcelHelper
{

    public static void Create(AuditingQuestionnaireSetupDto setup)
    {
        using (var excelPackage = new ExcelPackage())
        {
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            var worksheet = excelPackage.Workbook.Worksheets.Add("Setup");

            NewRow(worksheet);
            // Image img = Image.FromFile(@"Sample.png");  
            // ExcelPicture pic = worksheet.Drawings.AddPicture("Sample", img);  

            var titleCell = worksheet.Cells[worksheet.Dimension.Rows, 1, worksheet.Dimension.Rows, 5];
            titleCell.Value = setup.Title;
            titleCell.Style.Font.Bold = true;
            titleCell.Style.Font.Size = 16;
            titleCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            titleCell.Merge = true;

            SetCellValue(worksheet, "Facility", setup.Facility);
            Merge(worksheet);
            SetBorder(worksheet);
            NewRow(worksheet);

            SetCellValue(worksheet, "Site", setup.Site);
            Merge(worksheet);
            SetBorder(worksheet);
            NewRow(worksheet);

            SetCellValue(worksheet, "Department", setup.Department);
            Merge(worksheet);
            SetBorder(worksheet);
            NewRow(worksheet);

            SetCellValue(worksheet, "CustomField", setup.CustomField);
            Merge(worksheet);
            SetBorder(worksheet);
            NewRow(worksheet);

            SetCellValue(worksheet, "Site Manager", setup.SiteManager);
            SetCellValueSameRow(worksheet, "Site Manager Title", setup.SiteManagerTitle);

            SetCellValue(worksheet, "Audit Manager", setup.AuditManager);
            SetCellValueSameRow(worksheet, "Audit Manager Title", setup.AuditManagerTitle);

            SetCellValue(worksheet, "Start Date", setup.InspectionStartDate.ToShortDateString());
            SetCellValueSameRow(worksheet, "End Date", setup.InspectionEndDate.ToShortDateString());

            NewRow(worksheet);

            var inspectionCell = worksheet.Cells[worksheet.Dimension.Rows + 1, 1, worksheet.Dimension.Rows + 1, 5];
            inspectionCell.Value = "Inspection";
            inspectionCell.Style.Font.Bold = true;
            inspectionCell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            inspectionCell.Merge = true;

            NewRow(worksheet);

            SetCellValue(worksheet, "Lead Inspector", setup.LeadInspector);
            SetCellValueSameRow(worksheet, "Lead Inspector Title", setup.LeadInspectorTitle);

            SetCellValue(worksheet, "Site Inspector1", setup.SiteInspector1);
            SetCellValueSameRow(worksheet, "Site Inspector1 Title", setup.SiteInspector1Title);

            SetCellValue(worksheet, "Site Inspector2", setup.SiteInspector2);
            SetCellValueSameRow(worksheet, "Site Inspector2 Title", setup.SiteInspector2Title);

            SetCellValue(worksheet, "Other Site Inspectors", setup.OtherSiteInspectors);
            SetCellValueSameRow(worksheet, "Other Site Inspectors Title", setup.OtherSiteInspectorsTitle);
            NewRow(worksheet);

            SetCellValue(worksheet, "Notes", setup.Notes);
            Merge(worksheet);
            SetBorder(worksheet);

            worksheet.Cells.AutoFitColumns();
            worksheet.Column(2).Width = 40;
            worksheet.Column(5).Width = 40;

            var fi = new FileInfo(@"F:\File.xlsx");
            excelPackage.SaveAs(fi);
        }
    }

    private static void SetBorder(ExcelWorksheet worksheet)
    {
        worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
    }

    private static void Merge(ExcelWorksheet worksheet)
    {
        worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 5].Merge = true;
    }

    private static void SetCellValue(ExcelWorksheet worksheet, string name, string value)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        SetLabel(worksheet.Cells[row, 1], name);
        SetValue(worksheet.Cells[row, 2], value);
    }

    private static void SetCellValueSameRow(ExcelWorksheet worksheet, string name, string value)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows;
        SetLabel(worksheet.Cells[row, 4], name);
        SetValue(worksheet.Cells[row, 5], value);
    }

    private static void SetValue(ExcelRange cell, string value)
    {
        cell.Value = value;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
        cell.Style.WrapText = true;
        if (string.IsNullOrEmpty(value))
        {
            Color colFromHex = System.Drawing.ColorTranslator.FromHtml("#DCE6F0");
            cell.Style.Fill.PatternType = ExcelFillStyle.Solid;
            cell.Style.Fill.BackgroundColor.SetColor(colFromHex);
        }
    }

    private static void SetLabel(ExcelRange cell, string name)
    {
        cell.Value = name + ":";
        cell.Style.Font.Bold = true;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;
    }

    private static void NewRow(ExcelWorksheet worksheet)
    {
        int newRow = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        worksheet.Cells[newRow, 1].Value = "";
    }

    public static void Open()
    {
        var fi = new FileInfo(@"F:\File.xlsx");
        using (var excelPackage = new ExcelPackage(fi))
        {
            var firstWorksheet = excelPackage.Workbook.Worksheets[0];

            var namedWorksheet = excelPackage.Workbook.Worksheets["SomeWorksheet"];

            var anotherWorksheet =
                excelPackage.Workbook.Worksheets.FirstOrDefault(x => x.Name == "SomeWorksheet");

            string valA1 = firstWorksheet.Cells["A1"].Value.ToString();
            string valB1 = firstWorksheet.Cells[1, 2].Value.ToString();

            excelPackage.Save();
        }
    }
}
