using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpenXmlTest
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Determines if a chart exists on given sheet.");
            Console.WriteLine("Provide sheet ID (property name:sheetId) and chart ID (property name:graphicFrame.Id)");

            Console.WriteLine("Insert Sheet Id:");
            var sheetId = Convert.ToInt32(Console.ReadLine());

            Console.WriteLine("Insert chart Id (guid without curly brackets):");
            var chartId = Console.ReadLine();
                        
            var filePath = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "Test.xlsx");

            var fileManager = new ExampleClass(filePath, sheetId, chartId);
            fileManager.ReportProgress += FileManager_ReportProgress;
            fileManager.SimpleChartID = true;

            var retVal = fileManager.ChartExists;
            Console.WriteLine($"The chart was {(retVal ? "FOUND" : "NOT FOUND")}");
            Console.ReadLine();
            fileManager.Dispose();            
        }

        private static void FileManager_ReportProgress(string message)
        {
            Console.WriteLine(message);
        }
    }
}
