using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.IO;
using System.Text;


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
        foreach (string docxFilePath in Directory.GetFiles(inputDirectory, "*.docx"))
        {
            string htmlFilePath = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(docxFilePath) + ".html");

            string htmlContent = ConvertDocxToHtml(docxFilePath);
            File.WriteAllText(htmlFilePath, htmlContent);

            Console.WriteLine($"Converted {docxFilePath} to {htmlFilePath}");
        }

        Console.WriteLine("All documents converted!");
    }

    static string ConvertDocxToHtml(string filePath)
    {
        StringBuilder htmlBuilder = new StringBuilder();
        using (WordprocessingDocument wordDoc = WordprocessingDocument.Open(filePath, false))
        {
            Body body = wordDoc.MainDocumentPart.Document.Body;
            htmlBuilder.AppendLine("<html><body>");
            foreach (var element in body.Elements())
            {
                ConvertElementToHtml(element, htmlBuilder);
            }
            htmlBuilder.AppendLine("</body></html>");
        }
        return htmlBuilder.ToString();
    }

    static void ConvertElementToHtml(OpenXmlElement element, StringBuilder htmlBuilder)
    {
        if (element is Paragraph)
        {
            htmlBuilder.AppendLine("<p>");
            foreach (var childElement in element.Elements<Run>())
            {
                ConvertRunToHtml(childElement, htmlBuilder);
            }
            htmlBuilder.AppendLine("</p>");
        }
        else if (element is Table)
        {
            htmlBuilder.AppendLine("<table border='1'>");
            foreach (var row in element.Elements<TableRow>())
            {
                htmlBuilder.AppendLine("<tr>");
                foreach (var cell in row.Elements<TableCell>())
                {
                    htmlBuilder.AppendLine("<td>");
                    foreach (var childElement in cell.Elements<Paragraph>())
                    {
                        ConvertElementToHtml(childElement, htmlBuilder);
                    }
                    htmlBuilder.AppendLine("</td>");
                }
                htmlBuilder.AppendLine("</tr>");
            }
            htmlBuilder.AppendLine("</table>");
        }
        // Add more cases as needed for other elements like headers, lists, etc.
    }

    static void ConvertRunToHtml(Run run, StringBuilder htmlBuilder)
    {
        foreach (var text in run.Elements<Text>())
        {
            htmlBuilder.Append(text.Text);
        }
    }
}
