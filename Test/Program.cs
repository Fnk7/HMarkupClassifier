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
            Console.WriteLine("Finish.");
        }
    }
}
