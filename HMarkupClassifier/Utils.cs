using System;

namespace HMarkupClassifier
{
    public static class Utils
    {
        public static int ParseColumn(string col)
        {
            int temp = 0;
            foreach (var c in col)
                temp = temp * 26 + c - 'A' + 1;
            return temp;
        }

        public static int RunPython(string pythonFile, string argument)
        {
            using (System.Diagnostics.Process process = new System.Diagnostics.Process())
            {
                process.StartInfo.FileName = "python";
                process.StartInfo.Arguments = $"{pythonFile} {argument}";
                process.StartInfo.UseShellExecute = false;
                process.StartInfo.RedirectStandardOutput = true;
                process.Start();
                process.WaitForExit();
                var message = process.StandardOutput.ReadToEnd();
                Console.WriteLine(message);
                return process.ExitCode;
            }
        }
    }
}
