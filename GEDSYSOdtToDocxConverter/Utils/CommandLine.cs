using System;
using System.Diagnostics;

namespace GEDSYSOdtToDocxConverter.Utils
{
    class CommandLine
    {
        public static int RunLibreOfficeConverter(string target, string outdir, string convertTo = "pdf")
        {
            Process p = new Process();
            String processName = "soffice.exe";
            String commandArgs = String.Format("--headless --convert-to {2} --outdir \"{1}\" \"{0}\"",
                target, outdir, convertTo);
            String libreOfficeDir = Environment.GetEnvironmentVariable("libreoffice");

            ProcessStartInfo s = new ProcessStartInfo(processName, commandArgs);

            s.WindowStyle = ProcessWindowStyle.Hidden;
            s.CreateNoWindow = true;
            s.UseShellExecute = true;
            p.StartInfo = s;

            s.WorkingDirectory = libreOfficeDir;
            p.Start();
            p.WaitForExit();
            //Console.WriteLine("Exit code: " + p.ExitCode);
            GC.Collect();
            return p.ExitCode;
        }
    }
}
