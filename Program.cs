using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace HoursAccrued
{
    class Program
    {
        public static Excel.Application xP = new Excel.Application();

        static void Main(string[] args)
        {
            Console.WriteLine("Type \"start\" to begin recording your time . . .");
            string start = Console.ReadLine().ToLower();
            if (start == "start")
            {
                string path = "C:\\Users\\Ry\\Desktop\\WorkHours.txt";
                DateTime now = DateTime.Now;
                InPunch(now, path);
                Console.WriteLine("Do not close this window!\nType \"stop\" when you are done . . .");
                Console.ReadLine();
                DateTime newNow = DateTime.Now;
                TimeSpan diff = newNow.Subtract(now);
                double hours = Math.Round(((int)diff.Minutes / 60.00), 2);
                string punch = $"[OUT-PUNCH: {newNow.Hour}:{newNow.Minute}]\nHOURS WORKED: {hours}\n";
                using StreamWriter sw = File.AppendText(path);
                sw.WriteLine(punch);
                /*try
                    *{
                    *    ExcelWrite(diff.Minutes);
                    *}
                    *catch (FileNotFoundException e)
                    *{
                    *    File.Create(e.FileName);
                    *}
                    */
            }
            else
            {
                Console.WriteLine("Seriously?");
            }
        }

        static void InPunch(DateTime now, string path)
        {
            string punch = $"[IN-PUNCH: {now.Hour}:{now.Minute}]";
            if (!File.Exists(path))
            {
                using StreamWriter sw = File.CreateText(path);
                sw.WriteLine(punch);
            }
            else
            {
                using StreamWriter sw = File.AppendText(path);
                sw.WriteLine(punch);
            }
            Console.WriteLine(punch);
        }

        static void ExcelWrite(int min)
        {
            string xPath = "C:\\Users\\Ry\\Desktop\\HoursAccrued3.xlsx";
            Excel.Workbook xB = xP.Workbooks.Open(xPath);
            Excel.Worksheet? xS = xP.ActiveSheet as Excel.Worksheet;
            // Excel.Range cell = xS.Range["A1"];
            // int total = (int)cell.Value;
            xS.Cells["A4"] = min;
            xB.Close();
            xP.Quit();
        }
    }
}
