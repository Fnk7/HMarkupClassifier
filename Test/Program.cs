using System;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            HMarkupClassifier.Utils.ParseDataset("D:\\Workspace\\HMarkupDataset\\Annotated", "D:\\Workspace\\HMarkupDataset\\CSV-10");
            //HMarkupClassifier.Utils.Test();
            Console.WriteLine("Finish.");
        }
    }
}
