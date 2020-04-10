using System;
using HMarkupClassifier;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            Tools.ParseMarkDst("D:\\Temp\\Test\\Marked", "D:\\Temp\\Test\\CSV\\CSV-New-1");
            Console.WriteLine("Finish.");
        }
    }
}
