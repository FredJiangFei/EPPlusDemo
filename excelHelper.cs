

using System;
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


            var results = MockData.GetResult();
            var worksheet = excelPackage.Workbook.Worksheets.Add("Results_Dashboard");
            AddSheetHeaderImages(worksheet);
            SetTitle(worksheet, setup.Title + " - Applicable Regulations - Results Dashboard", 16);
            NewRow(worksheet);



            var fi = new FileInfo(@"F:\File.xlsx");
            excelPackage.SaveAs(fi);
        }
    }

    private static void CreateSetupSheet(ExcelPackage package, AuditingQuestionnaireSetupDto setup)
    {
        var worksheet = package.Workbook.Worksheets.Add("Setup");
        AddSheetHeaderImages(worksheet);

        SetTitle(worksheet, setup.Title + " - Audit", 16);

        SetCellValue(worksheet, "Facility", setup.Facility);
        NewRow(worksheet);

        SetCellValue(worksheet, "Site", setup.Site);
        NewRow(worksheet);

        SetCellValue(worksheet, "Department", setup.Department);
        NewRow(worksheet);

        SetCellValue(worksheet, "CustomField", setup.CustomField);
        NewRow(worksheet);

        SetCellValue(worksheet,
        "Site Manager", setup.SiteManager,
        "Site Manager Title", setup.SiteManagerTitle);
        SetCellValue(worksheet,
        "Audit Manager", setup.AuditManager,
        "Audit Manager Title", setup.AuditManagerTitle);
        SetCellValue(worksheet,
        "Managers", setup.Managers,
        "Managers Title", setup.ManagersTitle);
        NewRow(worksheet);

        SetTitle(worksheet, "Inspection");
        NewRow(worksheet);

        var startDate = setup.InspectionStartDate.ToShortDateString();
        var endDate = setup.InspectionEndDate.ToShortDateString();
        SetCellValue(worksheet, "Start Date", startDate, "End Date", endDate);
        NewRow(worksheet);

        SetCellValue(worksheet,
        "Lead Inspector", setup.LeadInspector,
        "Lead Inspector Title", setup.LeadInspectorTitle);
        SetCellValue(worksheet,
        "Site Inspector1", setup.SiteInspector1,
        "Site Inspector1 Title", setup.SiteInspector1Title);
        SetCellValue(worksheet,
        "Site Inspector2", setup.SiteInspector2,
        "Site Inspector2 Title", setup.SiteInspector2Title);
        SetCellValue(worksheet,
        "Other Site Inspectors", setup.OtherSiteInspectors,
        "Other Site Inspectors Title", setup.OtherSiteInspectorsTitle);
        NewRow(worksheet);

        SetCellValue(worksheet, "Notes", setup.Notes);
        worksheet.Row(worksheet.Dimension.Rows).Height = 100;

        worksheet.Cells.AutoFitColumns();
        worksheet.Column(2).Width = 40;
        worksheet.Column(5).Width = 40;
    }
    
    private static void AddSheetHeaderImages(ExcelWorksheet worksheet)
    {
        NewRow(worksheet);
        worksheet.Row(worksheet.Dimension.Rows).Height = 50;
        Image img = Image.FromFile(@"Xcelerator.png");
        ExcelPicture pic = worksheet.Drawings.AddPicture("Xcelerator", img);
        pic.SetPosition(0, 5, 0, 0);

        Image img2 = Image.FromFile(@"STP.png");
        ExcelPicture pic2 = worksheet.Drawings.AddPicture("STP", img2);
        pic2.SetPosition(0, 5, 4, 15);
    }

    private static void SetCellValue(ExcelWorksheet worksheet, params string[] values)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;

        var cell1 = worksheet.Cells[row, 1];
        SetLabel(cell1, values[0]);

        var cell2 = worksheet.Cells[row, 2];
        SetValue(cell2, values[1]);

        if (values.Length == 2)
        {
            worksheet.Cells[row, 2, row, 5].Merge = true;
            worksheet.Cells[row, 2, row, 5].Style.Border.BorderAround(ExcelBorderStyle.Medium);
            return;
        }

        var cell3 = worksheet.Cells[row, 4];
        SetLabel(cell3, values[2]);

        var cell4 = worksheet.Cells[row, 5];
        SetValue(cell4, values[3]);
    }

    private static void SetValue(ExcelRange cell, string value)
    {
        cell.Value = value;
        cell.Style.Border.BorderAround(ExcelBorderStyle.Medium);
        cell.Style.WrapText = true;
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;

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
        cell.Style.VerticalAlignment = ExcelVerticalAlignment.Top;
    }

    private static void NewRow(ExcelWorksheet worksheet)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        var cell = worksheet.Cells[row, 1, row, 5];
        cell.Value = "";
        cell.Merge = true;
    }

    private static void SetTitle(ExcelWorksheet worksheet, string value, int fontSize = 11)
    {
        int row = worksheet.Dimension == null ? 1 : worksheet.Dimension.Rows + 1;
        var cell = worksheet.Cells[row, 1, row, 5];
        cell.Value = value;
        cell.Style.Font.Bold = true;
        cell.Style.Font.Size = fontSize;
        cell.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
        cell.Merge = true;
    }
}
