using System;
using System.Data;
using System.IO;
using System.Data.OleDb;
using System.Reflection;
using NUnit.Framework;
using ExcelTools = Ms.Office;
using Excel = Microsoft.Office.Interop.Excel;

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
            string file = @"C:\GitHub\Fantasy-Hockey-Projector\PlayerStats.xlsx";
            Console.WriteLine(file);

            Excel.Application excel = null;
            Excel.Workbook wkb = null;

            try
            {
                excel = new Excel.Application();

                wkb = ExcelTools.OfficeUtil.OpenBook(excel, file);

                Excel.Worksheet sheet = wkb.Sheets["Data"] as Excel.Worksheet;

                Excel.Range range = null;

                if (sheet != null)
                    range = sheet.get_Range("A1", Missing.Value);

                string A1 = String.Empty;

                if (range != null)
                    A1 = range.Text.ToString();

                Console.WriteLine("A1 value: {0}", A1);

            }
            catch (Exception ex)
            {
                //if you need to handle stuff
                Console.WriteLine(ex.Message);
            }
            finally
            {
                if (wkb != null)
                    ExcelTools.OfficeUtil.ReleaseRCM(wkb);

                if (excel != null)
                    ExcelTools.OfficeUtil.ReleaseRCM(excel);
            }
        }
    }
}
        
