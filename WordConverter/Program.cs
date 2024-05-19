using System;
using System.Diagnostics;
using System.IO;

namespace LibreOfficeConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            string inputDirectory = @"C:\Documents";
            string outputDirectory = @"C:\Documents\html";

            // Ensure the output directory exists
            if (!Directory.Exists(outputDirectory))
            {
                Directory.CreateDirectory(outputDirectory);
            }

            // Process each .docx file in the input directory
            foreach (string wordFilePath in Directory.GetFiles(inputDirectory, "*.docx"))
            {
                string htmlFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(wordFilePath) + ".html");
                ConvertWordToHtml(wordFilePath, htmlFilePath);
                Console.WriteLine($"Converted {wordFilePath} to {htmlFilePath}");
            }

            Console.WriteLine("All documents converted!");
        }

        public static void ConvertWordToHtml(string wordFilePath, string htmlFilePath)
        {
            try
            {
                string libreOfficePath = "C:\\Program Files\\LibreOffice\\program\\soffice.exe"; // Ensure this is in your PATH
                string arguments = $"--headless --convert-to html \"{wordFilePath}\" --outdir \"{Path.GetDirectoryName(htmlFilePath)}\"";

                ProcessStartInfo processStartInfo = new ProcessStartInfo
                {
                    FileName = libreOfficePath,
                    Arguments = arguments,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true,
                    UseShellExecute = false,
                    CreateNoWindow = true
                };

                using (Process process = new Process { StartInfo = processStartInfo })
                {
                    process.Start();
                    process.WaitForExit();

                    string output = process.StandardOutput.ReadToEnd();
                    string error = process.StandardError.ReadToEnd();

                    if (process.ExitCode != 0)
                    {
                        Console.WriteLine($"Error during conversion: {error}");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error during conversion: " + ex.Message);
            }
        }
    }
}
