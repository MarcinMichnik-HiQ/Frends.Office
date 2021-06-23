﻿using System;
using ClosedXML.Excel;
using System.Data;
using Newtonsoft.Json.Linq;
using System.ComponentModel;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace Frends.Office
{
    /// <summary>
    /// Input for file writers.
    /// </summary>
    public class WriteExcelFileInput : IWriteExcelFileInput
    {
        /// <summary>
        /// Input csv string.
        /// </summary>
        [DefaultValue("\"1;2;3\\r\\na;b;c\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string csv { get; set; }

        /// <summary>
        /// Determines what character will be used for splitting based on cell in csv. Deafult is ';'.
        /// </summary>
        [DefaultValue("';'")]
        [DisplayFormat(DataFormatString = "Expression")]
        public char cellDelimiter { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\"\\r\\n\"")]
        [DisplayFormat(DataFormatString = "Expression")]
        public string lineDelimiter { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .xlsx, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        [DisplayFormat(DataFormatString = "Text")]
        public string path { get; set; }

        /// <summary>
        /// This method parses the input csv string and returns DataTable object.
        /// </summary>
        public DataTable CsvToDataTable()
        {
            List<Dictionary<string, string>> parsedResult = new List<Dictionary<string, string>>();
            string[] records = csv.Split(new string[] { lineDelimiter }, StringSplitOptions.None);

            DataTable table = new DataTable();
            int recordsPerLine = 0;

            foreach (string record in records)
            {
                recordsPerLine = 0;
                string[] fields = record.Split(cellDelimiter);
                Dictionary<string, string> recordItem = new Dictionary<string, string>();
                int i = 0;
                foreach (string field in fields)
                {
                    recordItem.Add(i.ToString(), field);
                    i++;
                    recordsPerLine++;
                }
                parsedResult.Add(recordItem);
            }

            Dictionary<string, string> row = parsedResult[0];
            foreach (KeyValuePair<string, string> pair in row)
            {
                table.Columns.Add(pair.Value);
            }
            parsedResult.RemoveAt(0);

            foreach (Dictionary<string, string> dic in parsedResult)
            {
                DataRow workRow = table.NewRow();

                int counter = 0;
                foreach (KeyValuePair<string, string> y in dic)
                {
                    workRow[counter] = y.Value;
                    counter++;
                }

                table.Rows.Add(workRow);
            }
            return table;
        }
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
        public static JToken WriteExcelFile(WriteExcelFileInput input)
        {
            JToken taskResponse = JToken.Parse("{}");
            IXLWorkbook workbook = new XLWorkbook();
            DataTable dataTable;

            try
            {
                dataTable = input.CsvToDataTable();
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to build DataTable from csv.", ex);
            }

            IXLWorksheet mainWorksheet = workbook.Worksheets.Add(dataTable, "Default");

            // Adjust rows and columns to text length
            mainWorksheet.Rows().AdjustToContents();
            mainWorksheet.Columns().AdjustToContents();

            try
            {
                workbook.SaveAs(input.path);
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to save the file.", ex);
            }

            taskResponse["message"] = "The file has been written correctly.";
            taskResponse["filePath"] = input.path;

            return taskResponse;
        }
    }
}
