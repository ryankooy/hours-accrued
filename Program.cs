using System;
using System.IO;
using Excel = Microsoft.Office.Interop.Excel;

namespace HoursAccrued
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Type \"start\" to begin recording your time . . .");
            string start = Console.ReadLine().ToLower();
            if (start == "start")
            {
                DateTime now = DateTime.Now;
                LogTime(now);
                Console.WriteLine("Do not close this window!\nType \"stop\" when you are done . . .");
                string stop = Console.ReadLine().ToLower();
                if (stop == "stop")
                {
                    DateTime newNow = DateTime.Now;
                    LogTime(newNow);
                    TimeSpan diff = newNow.Subtract(now);
                    Console.WriteLine($"You've worked {diff.Hours} hours, {diff.Minutes} minutes.");
					string path = "C:\\Users\\Ry\\Desktop\\WorkHours.txt";
					if (!File.Exists(path))
					{
                        using StreamWriter sw = File.CreateText(path);
						sw.WriteLine($"{diff.Hours}:{diff.Minutes}");
					}
					else
					{
                        using StreamWriter sw = File.AppendText(path);
                        string punch = $"PUNCH: {newNow.Hour}:{newNow.Minute}\n>{diff.Minutes} worked";
                        sw.WriteLine(punch);
                        Console.WriteLine(punch);
                    }
                    try
                    {
                        ExcelWrite(diff.Minutes);
                    }
                    catch (FileNotFoundException e)
                    {
                        File.Create(e.FileName);
                    }
                }
                else
                {
                    Console.WriteLine("Bye.");
                }
            }
            else
            {
                Console.WriteLine("Seriously?");
            }
        }

        static void LogTime(DateTime now)
        {
            Console.WriteLine($"Current time: {now.Hour}:{now.Minute}:{now.Second}");
        }

        static void ExcelWrite(int min)
        {
            string xPath = "C:\\Users\\Ry\\Desktop\\HoursAccrued2.xlsx";
            Excel.Application xP = new Excel.Application();
            Excel.Workbook xB = xP.Workbooks.Open(xPath);
            Excel.Worksheet? xS = xP.ActiveSheet as Excel.Worksheet;
            Excel.Range cell = xS.Range["A1"];
            int total = (int)cell.Value;
            xS.Cells["A1"] = total + min;
            xB.Close();
            xP.Quit();
        }
    }
}
