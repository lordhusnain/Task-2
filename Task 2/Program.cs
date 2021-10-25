using ClosedXML.Excel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace Task_2
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Hello World!");
            PrepareExcelFile();
        }

        static void PrepareExcelFile()
        {
            string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Data\task 2 - hotelrates.json");

            Console.WriteLine("Reading JSON from -->" + path);
            string allText = System.IO.File.ReadAllText(path);
            var hotelRateModel = JsonConvert.DeserializeObject<HotelRateModel>(allText);

            var workbook = new XLWorkbook();
            workbook.AddWorksheet("Rates");
            var ws = workbook.Worksheet("Rates");


            ws.Cell(1, 1).Value = "ARRIVAL_DATE";
            ws.Cell(1, 2).Value = "DEPARTURE_DATE";
            ws.Cell(1, 3).Value = "PRICE";
            ws.Cell(1, 4).Value = "CURRENCY";
            ws.Cell(1, 5).Value = "RATENAME";
            ws.Cell(1, 6).Value = "ADULTS";
            ws.Cell(1, 7).Value = "BREAKFAST_INCLUDED";

            int row = 2;
            foreach (var item in hotelRateModel.HotelRates)
            {
                int column = 1;
                ws.Cell(row, column++).Value = item.TargetDay.ToString("dd.MM.yy");
                ws.Cell(row, column++).Value = item.TargetDay.AddDays(item.Los).ToString("dd.MM.yy");
                ws.Cell(row, column++).Value = item.Price.NumericInteger.ToString("N");
                ws.Cell(row, column++).Value = item.Price.Currency.ToString();
                ws.Cell(row, column++).Value = item.RateName.ToString();
                ws.Cell(row, column++).Value = item.Adults.ToString();
                ws.Cell(row, column++).Value = item.RateTags.Any(c => c.Name == "breakfast" && c.Shape) ? 1 : 0;
                row++;
            }

            ws.Columns().AdjustToContents();

            Console.WriteLine("Saving excel file at -->" + Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location));
            workbook.SaveAs("HotelRates.xlsx");
        }
    }
}
