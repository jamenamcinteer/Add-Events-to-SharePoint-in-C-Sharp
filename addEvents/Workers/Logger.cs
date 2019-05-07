using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace addEvents.Workers
{
    class Logger
    {
        public void Log(string msg, ConsoleColor color = ConsoleColor.DarkGray)
        {
            Console.ForegroundColor = color;
            Console.WriteLine($"[{DateTime.Now}] {msg}");
            Console.ResetColor();
        }

        public void LogConsoleAndFile(string msg, StreamWriter sw)
        {
            Console.WriteLine(msg);
            sw.WriteLine(msg);
        }
    }
}
