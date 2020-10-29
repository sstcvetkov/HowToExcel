using OfficeOpenXml;
using OfficeOpenXml.Drawing.Chart;
using OfficeOpenXml.Style;

namespace HowToExcel
{
    public class MarketExcelGenerator
    {
        public byte[] Generate(MarketReport report)
        {
            var package = new ExcelPackage();

            var sheet = package.Workbook.Worksheets
                .Add("Market Report");
            
            sheet.Cells["B2"].Value = "Company:";
            sheet.Cells[2, 3].Value = report.Company.Name;
            sheet.Cells["B3"].Value = "Location:";
            sheet.Cells["C3"].Value = $"{report.Company.Address}, " +
                                      $"{report.Company.City}, " +
                                      $"{report.Company.Country}";
            sheet.Cells["B4"].Value = "Sector:";
            sheet.Cells["C4"].Value = report.Company.Sector;
            sheet.Cells["B5"].Value = report.Company.Description;

            sheet.Cells[8, 2, 8, 4].LoadFromArrays(new object[][]{ new []{"Capitalization", "SharePrice", "Date"} });
            var row = 9;
            var column = 2;
            foreach (var item in report.History)
            {
                sheet.Cells[row, column].Value = item.Capitalization;
                sheet.Cells[row, column + 1].Value = item.SharePrice;
                sheet.Cells[row, column + 2].Value = item.Date;
                row++;
            }

            sheet.Cells[1, 1, row, column + 2].AutoFitColumns();
            sheet.Column(2).Width = 14;
            sheet.Column(3).Width = 12;
            
            sheet.Cells[9, 4, 9+ report.History.Length, 4].Style.Numberformat.Format = "yyyy";
            sheet.Cells[9, 2, 9+ report.History.Length, 2].Style.Numberformat.Format =  "### ### ### ##0";

            sheet.Column(2).Style.HorizontalAlignment = ExcelHorizontalAlignment.Left;
            sheet.Cells[8, 3, 8 + report.History.Length, 3].Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            sheet.Column(4).Style.HorizontalAlignment = ExcelHorizontalAlignment.Right;

            sheet.Cells[8, 2, 8, 4].Style.Font.Bold = true;
            sheet.Cells["B2:C4"].Style.Font.Bold = true;
            
            sheet.Cells[8, 2, 8 + report.History.Length, 4].Style.Border.BorderAround(ExcelBorderStyle.Double);
            sheet.Cells[8, 2, 8, 4].Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            
            var capitalizationChart = sheet.Drawings.AddChart("FindingsChart", OfficeOpenXml.Drawing.Chart.eChartType.Line);
            capitalizationChart.Title.Text = "Capitalization";
            capitalizationChart.SetPosition(7, 0, 5, 0);
            capitalizationChart.SetSize(800, 400);
            var capitalizationData = (ExcelChartSerie)(capitalizationChart.Series.Add(sheet.Cells["B9:B28"], sheet.Cells["D9:D28"]));
            capitalizationData.Header = report.Company.Currency;
            
            sheet.Protection.IsProtected = true;
            return package.GetAsByteArray();
        }
    }
}