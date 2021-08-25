using System;
using ClosedXML.Excel;
using System.Data;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace Frends.Office
{
    /// <summary>
    /// Input for file writers.
    /// </summary>
    public class WriteWordFileInput : IWriteFileInput
    {
        /// <summary>
        /// Input string data.
        /// </summary>
        [DefaultValue("\"Test input\\r\\nNew line\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string StringInput { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\"\\r\\n\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string LineDelimiter { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting pages. Default is "\br".
        /// </summary>
        [DefaultValue("\"\\br\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string PageDelimiter { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .docx, e.g. FileName.docx
        /// </summary>
        [DefaultValue(@"c:\file.docx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string TargetPath { get; set; }
    }

    /// <summary>
    /// Used for writing Excel files.
    /// </summary>
    public class WriteWordFileTask
    {
        /// <summary>
        /// Allows you to write word files in .docx format. Repository: https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Returns JToken.</returns>
        public static JToken WriteWordFile([PropertyTab] WriteWordFileInput input)
        {
            JToken taskResponse = JToken.Parse("{}");
            try
            {
                using (WordprocessingDocument wordDocument =
                    WordprocessingDocument.Create(input.TargetPath, WordprocessingDocumentType.Document))
                {
                    // Add a main document part. 
                    MainDocumentPart mainPart = wordDocument.AddMainDocumentPart();

                    // Create the document structure and add some text.
                    mainPart.Document = new Document();
                    Body body = mainPart.Document.AppendChild(new Body());
                    Paragraph para = body.AppendChild(new Paragraph());
                    Run run = para.AppendChild(new Run());

                    string[] pages = input.StringInput.Split(new string[] { input.PageDelimiter }, StringSplitOptions.None);

                    foreach (string page in pages)
                    {

                        string[] records = page.Split(new string[] { input.LineDelimiter }, StringSplitOptions.None);

                        foreach (string record in records)
                        {
                            run.AppendChild(new Text(record));
                            run.Append(new Break());
                        }

                        run.Append(new Break() { Type = BreakValues.Page });
                    }
                }
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to build and save word file.", ex);
            }

            taskResponse["message"] = "The file has been written correctly.";
            taskResponse["savedTo"] = input.TargetPath;

            return taskResponse;
        }
    }
}
