using System;
using System.Data;
using System.IO;
using System.Reflection;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace FantasyHockeyProjector
{
    class Parser
    {
        static void Main(string[] args)
        {
            ProcessWorkbook();
        }

        public static Dictionary<string, float> CalculateValuesAdded(Dictionary<string, List<float>> playersToStats, float[] statAverages)
        {
            Dictionary<string, float> playersToValueAdded = new Dictionary<string, float>();
            foreach (KeyValuePair<string, List<float>> entry in playersToStats)
            {
                float valueAdded = 0;
                for (int i = 0; i < statAverages.Length; i++)
                {
                    valueAdded += entry.Value[i] / statAverages[i];
                }
                playersToValueAdded.Add(entry.Key, valueAdded);
            }
            return playersToValueAdded;
        }

        public static void SortAndPrintValuesAdded(Dictionary<string, float> playersToValueAdded)
        {
            var valueAddedList = playersToValueAdded.ToList();
            valueAddedList.Sort((pair1, pair2) => pair2.Value.CompareTo(pair1.Value));
            int i = 1;
            foreach (KeyValuePair<string, float> entry in valueAddedList)
            {
                Console.WriteLine(i + ". " + entry);
                i++;
            }
        }

        public static float[] CalculateStatAverages(int numberOfPlayers, List<float> statTotals)
        {
            float[] statAverages = new float[statTotals.Count];
            for (int i = 0; i < statTotals.Count; i++)
            {
                statAverages[i] = statTotals[i] / numberOfPlayers;
            }
            return statAverages;
        }

        public static void ProcessWorkbook()
        {
            string file = @"C:\GitHub\Fantasy-Hockey-Projector\Summary.xlsx";
            int rowNumber = 0;
            int cellNumber = 0;
            int firstStatColumnIndex = 6;
            int firstStatRowIndex = 1;
            int columnNamesRowIndex = 0;
            List<string> statCategories = new List<string>();
            Dictionary<string, List<float>> playersToStats = new Dictionary<string, List<float>>();
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
                        if (rowNumber == columnNamesRowIndex && cellNumber >= firstStatColumnIndex)
                        {
                            statCategories.Add(cell.Text.ToString());
                        }

                        if (rowNumber == firstStatRowIndex && cellNumber >= firstStatColumnIndex)
                        {
                            statTotals.Add(0);
                        }

                        if (rowNumber >= firstStatRowIndex && cellNumber == columnNamesRowIndex)
                        {
                            playerName = cell.Text.ToString();
                            playersToStats.Add(playerName, new List<float>());
                        }

                        if (rowNumber >= firstStatRowIndex && cellNumber >= firstStatColumnIndex)
                        {
                            float statValue = float.Parse(cell.Text);
                            playersToStats[playerName].Add(statValue);
                            statTotals[cellNumber - 6] += statValue;
                        }
                        cellNumber++;
                    }
                    rowNumber++;
                }
            }

            int numberOfPlayers = playersToStats.Count;
            float[] statAverages = CalculateStatAverages(numberOfPlayers, statTotals);

            Dictionary<string, float> playersToValueAdded = CalculateValuesAdded(playersToStats, statAverages);

            SortAndPrintValuesAdded(playersToValueAdded);
        }



    }
}
        
