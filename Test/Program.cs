using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Test
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("Start...");
            //HMarkupClassifier.Utils.ParseDataset("D:\\Workspace\\HMarkupDataset\\Annotated", "D:\\Workspace\\HMarkupDataset\\CSV-1");
            HMarkupClassifier.Utils.Test();
            Console.WriteLine("Finish.");
        }
    }
}
