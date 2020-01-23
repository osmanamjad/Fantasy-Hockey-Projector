using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections;

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
            int rowTracker = 0;
            int cellTracker = 0;
            ArrayList stats = new ArrayList();
            foreach (var worksheet in Excel.Workbook.Worksheets(file))
            {
                foreach (var row in worksheet.Rows) 
                {
                    foreach (var cell in row.Cells)
                    {
                        if (rowTracker == 0 && cellTracker >= 5)
                        {
                            stats.Add(cell.Text);
                        }
                        cellTracker++;
                    }
                    rowTracker++;
                }
            }
        }
    }
}
        
