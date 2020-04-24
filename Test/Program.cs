using System;
using HMarkupClassifier.Predict;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            //Console.WriteLine("Start...");
            //Tools.ParseMarkDst("D:\\Temp\\TestMarked", "D:\\Temp\\TestCSV");
            //Console.WriteLine("Finish.");
            var result = Predict.PredictHeader("D:\\Workspace\\Python\\HeaderClf\\main.py", 0, "D:\\Temp\\test.xlsx", "Sheet");
            foreach (var (row, col) in result)
            {
                Console.WriteLine($"R{row}C{col}");
            }
        }
    }
}
