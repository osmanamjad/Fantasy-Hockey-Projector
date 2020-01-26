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
            int rowNumber = 0;
            int cellNumber = 0;
            List<string> statCategories = new List<string>();
            Dictionary<string, List<float>> players = new Dictionary<string, List<float>>();
            List<float> statTotals = new List<float>();
            string playerName = "";
            foreach (var worksheet in Excel.Workbook.Worksheets(file))
            {
                rowNumber = 0;
                foreach (var row in worksheet.Rows) 
                {
                    cellNumber = 0;
                    foreach (var cell in row.Cells)
                    {
                        if (rowNumber == 0 && cellNumber >= 6)
                        {
                            statCategories.Add(cell.Text.ToString());
                        }

                        if (rowNumber == 1 && cellNumber >= 6)
                        {
                            statTotals.Add(0);
                        }

                        if (rowNumber >= 1 && cellNumber == 0)
                        {
                            playerName = cell.Text.ToString();
                            players.Add(playerName, new List<float>());
                        }

                        if (rowNumber >= 1 && cellNumber >= 6)
                        {
                            float statValue = float.Parse(cell.Text);
                            players[playerName].Add(statValue);
                            statTotals[cellNumber - 6] += statValue;
                        }
                        //else if(rowTracker >= 6 && cellTracker >= )
                        cellNumber++;
                    }
                    rowNumber++;
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
        
