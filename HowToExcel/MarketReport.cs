using System;
namespace HowToExcel
{
    public class MarketReport
    {
        public Company Company { set; get; }
        public HistoryItem[] History { set; get; }
    }
    
    public class Company
    {
        public string Name { set; get; }
        public string Address { set; get; }
        public string City { set; get; }
        public string Country { set; get; }
        public string Currency { set; get; }
        public string Description { set; get; }
        public string Sector { set; get; }
    }
    
    public class HistoryItem
    {
        public long Capitalization { set; get; }
        public decimal SharePrice { set; get; }
        public DateTime Date { set; get; }
    }
}