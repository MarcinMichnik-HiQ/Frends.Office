using System;
using ClosedXML.Excel;
using System.Data;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Frends.Office
{
    /// <summary>
    /// Input for excel file writers.
    /// </summary>
    public class WriteExcelFileInput : IWriteFileInput
    {
        /// <summary>
        /// Input csv string.
        /// </summary>
        [DefaultValue("\"1;2;3\\r\\na;b;c\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string StringInput { get; set; }

        /// <summary>
        /// Determines what character will be used for splitting based on cell in csv. Deafult is ';'.
        /// </summary>
        [DefaultValue("';'")]
        [DisplayFormat(DataFormatString = "Expression")]
        public char CellDelimiter { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\"\\r\\n\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string LineDelimiter { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .xlsx, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string TargetPath { get; set; }
    }

    /// <summary>
    /// Used for writing Excel files.
    /// </summary>
    public class WriteExcelFileTask
    {
        /// <summary>
        /// Allows you to write excel files in .xlsx format. Repository: https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Returns JToken.</returns>
        public static JToken WriteExcelFile([PropertyTab] WriteExcelFileInput input)
        {
            JToken taskResponse = JToken.Parse("{}");

            XLWorkbook workbook = CreateWorkbookObject(input);
            try
            {
                workbook.SaveAs(input.TargetPath);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to save the file.", ex);
            }

            taskResponse["message"] = "The file has been written correctly.";
            taskResponse["filePath"] = input.TargetPath;

            return taskResponse;
        }

        /// <summary>
        /// This method creates an excel object from string input.
        /// </summary>
        public static XLWorkbook CreateWorkbookObject(WriteExcelFileInput input)
        {
            XLWorkbook workbook = new XLWorkbook();
            DataTable dataTable;

            try
            {
                dataTable = CsvToDataTable(input.StringInput, input.LineDelimiter, input.CellDelimiter);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to build DataTable from csv.", ex);
            }

            IXLWorksheet mainWorksheet = workbook.Worksheets.Add(dataTable, "Default");

            // Adjust rows and columns to text length
            mainWorksheet.Rows().AdjustToContents();
            mainWorksheet.Columns().AdjustToContents();

            return workbook;
        }

        /// <summary>
        /// This method parses the input csv string and returns DataTable object.
        /// </summary>
        public static DataTable CsvToDataTable(string input, string lineDelimiter, char cellDelimiter)
        {
            List<Dictionary<string, string>> parsed = new List<Dictionary<string, string>>();
            string[] rows = input.Split(new string[] { lineDelimiter }, StringSplitOptions.None);

            DataTable tableResult = new DataTable();

            // Populate list of dictionaries
            foreach (string row in rows)
            {
                string[] cells = row.Split(cellDelimiter);
                Dictionary<string, string> recordItem = new Dictionary<string, string>();

                int i = 0;
                foreach (string cell in cells)
                {
                    recordItem.Add(i.ToString(), cell);
                    i++;
                }
                parsed.Add(recordItem);
            }

            if (rows.Length > 0) {
                // Create columns. Their values will be first row values. This first row must not be included later to avoid duplicates.
                Dictionary<string, string> firstRow = parsed[0];
                foreach (KeyValuePair<string, string> pair in firstRow)
                {
                    tableResult.Columns.Add(pair.Value);
                }
                parsed.RemoveAt(0);

                // Populate rows except for first one (column names)
                foreach (Dictionary<string, string> dic in parsed)
                {
                    DataRow workRow = tableResult.NewRow();

                    int counter = 0;
                    foreach (KeyValuePair<string, string> y in dic)
                    {
                        workRow[counter] = y.Value;
                        counter++;
                    }

                    tableResult.Rows.Add(workRow);
                }
            }

            return tableResult;
        }
    }
}
