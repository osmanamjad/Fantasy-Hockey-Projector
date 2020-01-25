using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;

namespace FantasyHockeyProjector
{
    class Parser
    {
        static void Main(string[] args)
        {
            ProcessWorkbook();
        }

        public static void ProcessWorkbook()
        {
            string file = @"C:\GitHub\Fantasy-Hockey-Projector\Summary.xlsx";
            Console.WriteLine(file);
            int rowTracker = 0;
            int cellTracker = 0;
            List<string> statCategories = new List<string>();
            List<string> players = new List<string>();
            //float[] statTotals = new float[];
            List<float> statTotals = new List<float>();
            foreach (var worksheet in Excel.Workbook.Worksheets(file))
            {
                rowTracker = 0;
                foreach (var row in worksheet.Rows) 
                {
                    cellTracker = 0;
                    float individualStatTotal = 0;
                    foreach (var cell in row.Cells)
                    {
                        if (cellTracker >= 6 && rowTracker == 0)
                        {
                            statCategories.Add(cell.Text.ToString());
                        }
                        else if(cellTracker >= 6 && rowTracker > 0)
                        {
                            individualStatTotal += float.Parse(cell.Text);
                        }
                        if (rowTracker >=1  && cellTracker == 0)
                        {
                            players.Add(cell.Text.ToString());
                        }
                        else if(rowTracker >= 5 && cellTracker >= )
                        cellTracker++;
                    }

                    if (rowTracker >= 5)
                        statTotals.Add(individualStatTotal);

                    rowTracker++;
                }
            }
            float[] statAverages = new float[statTotals.Count];
            int numberOfPlayers = players.Count;
            for(int i = 0; i < statTotals.Count; i++)
            {
                statAverages[i] = statTotals[i] / numberOfPlayers;
            }
        }
    }
}
        
