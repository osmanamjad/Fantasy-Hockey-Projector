using System;
using System.Data;
using System.IO;
using System.Reflection;

namespace FantasyHockeyProjector
{
    class Parser
    {
        static void Main(string[] args)
        {
            Console.WriteLine("hello");
            ProcessWorkbook();
        }

        public static void ProcessWorkbook()
        {
            string file = @"C:\GitHub\Fantasy-Hockey-Projector\Summary.xlsx";
            Console.WriteLine(file);

            foreach (var worksheet in Excel.Workbook.Worksheets(file))
            {
                foreach (var row in worksheet.Rows) 
                {

                    foreach (var cell in row.Cells)
                    {
                        Console.WriteLine(cell.Text);
                    }
                }
            }
        }
    }
}
        
