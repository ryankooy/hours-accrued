using System;
using System.IO;
using System.Linq;
// using Excel = Microsoft.Office.Interop.Excel;

namespace HoursAccrued
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Enter \"start\" to begin recording your time.");
            string start = Console.ReadLine().ToLower();
            if (start == "start")
            {
                string path = @"C:\Users\Ry\Desktop\WorkHours.txt";
                string lastTotal = File.ReadLines(path).Last();
                DateTime now = DateTime.Now;
                InPunch(now, path);
                Console.WriteLine("Do not close this window!\nEnter \"stop\" when you're done . . .");
                Console.ReadLine();
                Console.WriteLine('.');
                Console.WriteLine(". .");
                Console.WriteLine(". . .");
                Console.WriteLine(". . . .");
                Console.WriteLine(". . . . .");
                Console.WriteLine(". . . .");
                Console.WriteLine(". . .");
                Console.WriteLine(". .");
                Console.WriteLine('.');
                Console.WriteLine("Great work!");
                DateTime newNow = DateTime.Now;
                TimeSpan diff = newNow.Subtract(now);
                double hours = Math.Round(((int)diff.Hours + ((int)diff.Minutes / 60.00)), 2);
                string punch = $"[OUT-PUNCH: {newNow.Hour}:{newNow.Minute} {newNow.Month}/{newNow.Day} {newNow.DayOfWeek}]\n+ {hours} hrs";
                double total = Convert.ToDouble(lastTotal);
                double newTotal = Math.Round(total + hours, 2);
                using StreamWriter sw = File.AppendText(path);
                sw.WriteLine(punch);
                sw.WriteLine("TOTAL HOURS FOR WEEK:");
                sw.WriteLine(newTotal);
            }
            else
            {
                Console.WriteLine("Seriously?");
            }
        }

        static void InPunch(DateTime now, string path)
        {
            string punch = $"\n[IN-PUNCH: {now.Hour}:{now.Minute} {now.Month}/{now.Day} {now.DayOfWeek}]";
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
    }
}
