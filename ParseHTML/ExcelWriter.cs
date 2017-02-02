using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;

namespace ParseHTML
{
    class ExcelWriter
    {
        Application xlApp;
        Workbook xlWorkBook;
        Worksheet xlWorkSheet;
        Range range;

        progress_bar PB;
        int i;

        public ExcelWriter() { }

        public void StartExcelApp()
        {
            xlApp = new Application();
        }

        public void redraw()
        {
            ClearLine();
            PB.print_progressBar(i-1);
            //Console.Write("\n");
            //Console.WriteLine("Writing data to Excel file:");
            //PB_Excel.print_progressBar(i);
        }

        public static void ClearLine()
        {
            int currentLineCursor = Console.CursorTop;
            Console.SetCursorPosition(0, Console.CursorTop);
            Console.Write(new string(' ', Console.WindowWidth));
            Console.SetCursorPosition(0, currentLineCursor);
        }

        public void WriteData_to_Excell(string path, Dictionary<string, string> Data)
        {
            PB = new progress_bar(0, Data.Count);
            Console.WriteLine("Writing data to Excel file:");
            if (System.IO.File.Exists(path))
            {
                xlWorkBook = xlApp.Workbooks.Open(path, 0, false, 5, "", "", true, XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Worksheet)xlWorkBook.Worksheets.get_Item(1);

                i = 1;
                foreach (KeyValuePair<string,string> data_row in Data)
                {
                    range = (Range)xlWorkSheet.Cells[i, 1];
                    range.Value = data_row.Key.ToString();

                    range = (Range)xlWorkSheet.Cells[i, 2];
                    range.Value = data_row.Value.ToString();

                    i++;
                    redraw();
                    //Console.WriteLine(String.Format("info wrote: \t{0}:\t{1}", data_row.Key, data_row.Value));
                    //System.Threading.Thread.Sleep(500);
                }

                xlWorkBook.Save();
                Console.WriteLine("\n\nExcel file saved");

                GC.Collect();
                GC.WaitForPendingFinalizers();
                xlWorkBook.Close(true, false, false);
                Marshal.FinalReleaseComObject(xlWorkBook);
            }
        }


        public void CloseExcelApp()
        {
            Marshal.FinalReleaseComObject(xlApp);
        }

        
    }
}
