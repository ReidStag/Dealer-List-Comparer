using System;
using ExcelDataReader;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;


namespace PleaseWork
{
    class Program
    {
        static void Main(string[] args)
        {
            System.Text.Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

            //Excel file paths
            string currDealerListPath = @"C:\Users\rstagemeyer\Documents\WordPress\Web Host\Workspace_Files\Docs\Dealers_Updated.csv";
            string compareDealerListPath = @"C:\Users\rstagemeyer\Documents\WordPress\Web Host\Workspace_Files\Docs\Customer_Master-Dealers_Keep_Justin_Trout.xlsx";
            string writeFilePath = @"C:\Users\rstagemeyer\Documents\WordPress\Web Host\Workspace_Files\Docs\Dealers_Add.txt";

            //parsed Data using ExcelDataReader package
            DataSet currDealerData = readExcelFile(currDealerListPath);
            DataSet compareDealerData = readExcelFile(compareDealerListPath);

            // The result of each spreadsheet is in result.Tables
            DataTable currDT = currDealerData.Tables[0];
            DataTable compareDT = compareDealerData.Tables[0];

            Dictionary<string, string> currDealerCities = new Dictionary<string, string>();
            Dictionary<string, string> currDealerStates = new Dictionary<string, string>();

            string lines = "";

            //fill Dictionaries with current dealer cities and states
            foreach (DataRow row in currDT.Rows)
            {
                string city = row[2].ToString().Trim().ToLower();
                string state = row[3].ToString().Trim();
                currDealerCities[city] = city;
                currDealerStates[state] = state;
            }

            // iterate through compare list and write entry to file if city, state not present in current list
            foreach (DataRow row in compareDT.Rows)
            {
                string city = row[7].ToString().Trim().ToLower();
                if (!currDealerCities.ContainsKey(city))
                {
                    string state = row[8].ToString().Trim();
                    string dealerName = row[3].ToString().Trim();
                    if(city.Length == 0)
                    {
                        city = row[5].ToString().Trim();
                    }
                    lines += city + ", " + state + "- " + dealerName + "\n";
                }
            }
            File.WriteAllText(writeFilePath, lines);
        }

        private static DataSet readExcelFile(string path)
        {
            DataSet result = null;
            if (path.Contains("csv"))
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    //csv file
                    using (var reader = ExcelReaderFactory.CreateCsvReader(stream))
                    {
                        // Use the AsDataSet extension method
                        result = reader.AsDataSet();
                    }
                }
            }
            else if (path.Contains(".xls") || path.Contains(".xlsx") || path.Contains(".xlsb"))
            {
                using (var stream = File.Open(path, FileMode.Open, FileAccess.Read))
                {
                    // Auto-detect format, supports:
                    //  - Binary Excel files (2.0-2003 format; *.xls)
                    //  - OpenXml Excel files (2007 format; *.xlsx, *.xlsb)
                    using (var reader = ExcelReaderFactory.CreateReader(stream))
                    {
                        // Use the AsDataSet extension method
                        result = reader.AsDataSet();
                    }
                }
            }
            return result;
        }
    }
}
