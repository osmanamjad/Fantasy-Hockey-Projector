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
            ArrayList statCategories = new ArrayList();
            ArrayList players = new ArrayList();
            ArrayList statTotals = new ArrayList();
            foreach (var worksheet in Excel.Workbook.Worksheets(file))
            {
                rowTracker = 0;
                foreach (var row in worksheet.Rows) 
                {
                    cellTracker = 0;
                    float individualStatTotal = 0;
                    foreach (var cell in row.Cells)
                    {
                        if (cellTracker == 0 && rowTracker >= 5)
                        {
                            statCategories.Add(cell.Text);
                        }
                        else if(cellTracker > 0 && rowTracker >=5)
                        {
                            individualStatTotal += float.Parse(cell.Text);
                        }
                        if (rowTracker == 0 && cellTracker >= 1)
                        {
                            players.Add(cell.Text);
                        }
                        cellTracker++;
                    }
                    statTotals.Add(individualStatTotal);
                    rowTracker++;
                }
            }
        }
    }
}
        
