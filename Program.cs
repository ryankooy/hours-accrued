using System;

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
    }
}
