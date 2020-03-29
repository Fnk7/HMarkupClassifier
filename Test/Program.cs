using System;
using HMarkupClassifier;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            Utils.ParseDataset(Utils.annotatedDataset, Utils.csvDataset);
            Utils.RunPython(Utils.pythonFile, $"train {Utils.csvDataset} {Utils.pythonModel}");
            Utils.RunPython(Utils.pythonFile, $"predict {Utils.pythonModel}");
            Console.WriteLine("Finish.");
        }
    }
}
