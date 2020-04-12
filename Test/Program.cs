using System;
using HMarkupClassifier;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            Tools.ParseMarkDst("D:\\Temp\\Marked-04-12-M1", "D:\\Temp\\Test\\CSV\\CSV-04-12-M1-C1");
            Console.WriteLine("Finish.");
        }
    }
}
