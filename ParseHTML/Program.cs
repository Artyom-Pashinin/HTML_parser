using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using HtmlAgilityPack;

namespace ParseHTML
{
    class Program
    {
        public static progress_bar PB_HTML;
        //public progress_bar PB_Excel;
        static int i;
        public static void redraw()
        {
            Console.Clear();
            Console.WriteLine("Collecting data from HTML files:");
            PB_HTML.print_progressBar(i);
            //Console.WriteLine("Writing data to Excel file:");
            //PB_Excel.print_progressBar(i);
        }
        public static void create_pb(int max)
        {
            PB_HTML = new progress_bar(0, max);
        }
        public static void Main(string[] args)
        {

            string HTMLfolder = @"D:\Отчеты для директора\ОТХОДЫ\HTML\Julia_";
            string XLSfilePath = @"D:\Отчеты для директора\ОТХОДЫ\HTML\L_Julia.xlsx";
            string temp = string.Empty;
            int row_with_waste = 0;

            List<string> HTMLs = Directory.GetFiles(HTMLfolder).ToList<string>();
            Dictionary<string, string> FinalData = new Dictionary<string, string>();
            HtmlDocument doc = new HtmlDocument();

            ExcelWriter EW = new ExcelWriter();

            create_pb(HTMLs.Count);
            //PB_Excel = new progress_bar(0, HTMLs.Count);

            i = 0;
            foreach (string item in HTMLs)
            {
                doc.Load(item);

                //Определяем какой TOPS создал HTML файл
                
                string doc_title = doc.DocumentNode.SelectSingleNode("//head/title").InnerText.ToString();

                if (doc_title.Contains("TruTops")) row_with_waste = 19;
                else row_with_waste = 17;

                List<List<string>> table = doc.DocumentNode.SelectNodes("//table")
                    .Descendants("tr")
                    .Skip(1)
                    .Select(tr => tr.Elements("td")
                    .Select(td => td.InnerText.Trim()).ToList())
                    .ToList();

                string waste = table[row_with_waste][1].ToString();
                waste = waste.Remove(waste.Length - 7);
                string program = item.Remove(0, 43);
                program = program.Remove(program.Length - 5);
                program = "L" + program;

                FinalData.Add(program, waste);
                i++;
                redraw();
                //Console.WriteLine(String.Format("info collected: \t{0}:\t{1}", program, waste));
                //System.Threading.Thread.Sleep(100);
            }

            Console.WriteLine("\n");

            EW.StartExcelApp();
            EW.WriteData_to_Excell(XLSfilePath, FinalData);
            EW.CloseExcelApp();
            

            Console.WriteLine("\nInformation has been collected successfully");
        }


    }
}