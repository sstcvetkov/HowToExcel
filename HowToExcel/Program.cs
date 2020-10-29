using System.IO;

namespace HowToExcel
{
    internal static class Program
    {
        static void Main(string[] args)
        {
            var reportData = new MarketReporter()
                .GetReport();
            var reportExcel = new MarketExcelGenerator()
                .Generate(reportData);
            
            File.WriteAllBytes("Report.xlsx", reportExcel);
        }
    }
}
