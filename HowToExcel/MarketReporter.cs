using System;

namespace HowToExcel
{
    public class MarketReporter
    {
        public MarketReport GetReport()
        {
            return new MarketReport
            {
                Company = new Company
                {
                    Address = "7th Park Way", Name = "Dream Inc.", City = "Dreamton", Country = "US", Currency = "USD",
                    Description = @"Dream Inc. is an American multinational technology company " +
                                  "headquartered in Dreamton, California. Designs, develops, and sells " +
                                  "consumer electronics, computer software, and online services.",
                    Sector = "Information Technology"
                },
                History = new []
                {
                    new HistoryItem{ Capitalization = 234823, SharePrice = 654m, Date = new DateTime(2000, 1, 1)}, 
                    new HistoryItem{ Capitalization = 434823, SharePrice = 923m, Date = new DateTime(2001, 1, 1)}, 
                    new HistoryItem{ Capitalization = 874823, SharePrice = 1289m, Date = new DateTime(2002, 1, 1)}, 
                    new HistoryItem{ Capitalization = 2134823, SharePrice = 1598m, Date = new DateTime(2003, 1, 1)}, 
                    new HistoryItem{ Capitalization = 2034823, SharePrice = 1543m, Date = new DateTime(2004, 1, 1)}, 
                    new HistoryItem{ Capitalization = 3124823, SharePrice = 1890m, Date = new DateTime(2005, 1, 1)}, 
                    new HistoryItem{ Capitalization = 4148523, SharePrice = 2538m, Date = new DateTime(2006, 1, 1)}, 
                    new HistoryItem{ Capitalization = 5843823, SharePrice = 5432m, Date = new DateTime(2007, 1, 1)}, 
                    new HistoryItem{ Capitalization = 7934823, SharePrice = 6934m, Date = new DateTime(2008, 1, 1)}, 
                    new HistoryItem{ Capitalization = 12434823, SharePrice = 8534m, Date = new DateTime(2009, 1, 1)}, 
                    new HistoryItem{ Capitalization = 23484823, SharePrice = 9523m, Date = new DateTime(2010, 1, 1)}, 
                    new HistoryItem{ Capitalization = 72234823, SharePrice = 2543m, Date = new DateTime(2011, 1, 1)}, 
                    new HistoryItem{ Capitalization = 90234823, SharePrice = 2540m, Date = new DateTime(2012, 1, 1)}, 
                    new HistoryItem{ Capitalization = 153234823, SharePrice = 2690m, Date = new DateTime(2013, 1, 1)}, 
                    new HistoryItem{ Capitalization = 190234823, SharePrice = 2783m, Date = new DateTime(2014, 1, 1)}, 
                    new HistoryItem{ Capitalization = 234832823, SharePrice = 2504m, Date = new DateTime(2015, 1, 1)}, 
                    new HistoryItem{ Capitalization = 590234823, SharePrice = 2932m, Date = new DateTime(2016, 1, 1)}, 
                    new HistoryItem{ Capitalization = 989234823, SharePrice = 3274m, Date = new DateTime(2017, 1, 1)}, 
                    new HistoryItem{ Capitalization = 1023234823, SharePrice = 4232m, Date = new DateTime(2018, 1, 1)}, 
                    new HistoryItem{ Capitalization = 3245234823, SharePrice = 5689m, Date = new DateTime(2019, 1, 1)} 
                }
            };
        }
    }
}