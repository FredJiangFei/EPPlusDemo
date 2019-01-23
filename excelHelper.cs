

using System;
using System.IO;
using System.Linq;
using OfficeOpenXml;
using OfficeOpenXml.Style;

class ExcelHelper
{
    private static void SetCellValue(ExcelWorksheet worksheet, string name, string value)
    {
        int newRow = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        worksheet.Cells[newRow, 1].Value = name + ":";
        worksheet.Cells[newRow, 1].Style.Font.Bold = true;

        worksheet.Cells[newRow, 2].Value = value;
        worksheet.Cells[newRow, 2].Style.Border.BorderAround(ExcelBorderStyle.Medium);
    }

    private static void NewRow(ExcelWorksheet worksheet)
    {
        int newRow = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        worksheet.Cells[newRow, 1].Value = "";
    }

    private static void SetCellValueSameRow(ExcelWorksheet worksheet, string name, string value)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows;
        worksheet.Cells[row, 3].Value = name + ":";
        worksheet.Cells[row, 3].Style.Font.Bold = true;

        worksheet.Cells[row, 4].Value = value;
        worksheet.Cells[row, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
    }

    public static void Create(AuditingQuestionnaireSetupDto setup)
    {
        using (var excelPackage = new ExcelPackage())
        {
            excelPackage.Workbook.Properties.Created = DateTime.Now;

            var worksheet = excelPackage.Workbook.Worksheets.Add("Setup");

            SetCellValue(worksheet, "Facility", setup.Facility);
            worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 4].Merge = true;
            worksheet.Cells[worksheet.Dimension.Rows, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells[worksheet.Dimension.Rows, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);

            NewRow(worksheet);

            SetCellValue(worksheet, "Site", setup.Site);
            worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 4].Merge = true;
            worksheet.Cells[worksheet.Dimension.Rows, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells[worksheet.Dimension.Rows, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            NewRow(worksheet);

            SetCellValue(worksheet, "Department", setup.Department);
            worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 4].Merge = true;
            worksheet.Cells[worksheet.Dimension.Rows, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells[worksheet.Dimension.Rows, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            NewRow(worksheet);

            SetCellValue(worksheet, "CustomField", setup.CustomField);
            worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 4].Merge = true;
            worksheet.Cells[worksheet.Dimension.Rows, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells[worksheet.Dimension.Rows, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            NewRow(worksheet);

            SetCellValue(worksheet, "Site Manager", setup.SiteManager);
            SetCellValueSameRow(worksheet, "Site Manager Title", setup.SiteManagerTitle);

            SetCellValue(worksheet, "Audit Manager", setup.AuditManager);
            SetCellValueSameRow(worksheet, "Audit Manager Title", setup.AuditManagerTitle);

            SetCellValue(worksheet, "Start Date", setup.InspectionStartDate.ToShortDateString());
            SetCellValueSameRow(worksheet, "End Date", setup.InspectionEndDate.ToShortDateString());
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
            worksheet.Cells[worksheet.Dimension.Rows, 2, worksheet.Dimension.Rows, 4].Merge = true;
            worksheet.Cells[worksheet.Dimension.Rows, 3].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            worksheet.Cells[worksheet.Dimension.Rows, 4].Style.Border.BorderAround(ExcelBorderStyle.Medium);

            worksheet.Cells.AutoFitColumns();
            var fi = new FileInfo(@"F:\File.xlsx");
            excelPackage.SaveAs(fi);
        }
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
