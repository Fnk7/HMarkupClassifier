using ClosedXML.Excel;
using System.Text;
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

        public static int GetColor(XLColor xlColor)
        {
            switch (xlColor.ColorType)
            {
                case XLColorType.Theme:
                    int theme = (int)xlColor.ThemeColor;
                    int tint = (int)xlColor.ThemeTint * 10;
                    return (theme << 24) + tint;
                case XLColorType.Indexed:
                    return xlColor.Indexed;
                default:
                    return xlColor.Color.ToArgb();
            }
        }

        public static int RunPython(string pythonFile, string[] argument)
        {
            using (System.Diagnostics.Process process = new System.Diagnostics.Process())
            {
                process.StartInfo.FileName = "python";
                if(argument == null || argument.Length == 0)
                    process.StartInfo.Arguments = $"{pythonFile}";
                else
                    process.StartInfo.Arguments = $"{pythonFile} {string.Join(" ", argument)}";
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
