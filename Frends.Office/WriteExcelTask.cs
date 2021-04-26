using System;
using System.Collections.Generic;
using System.ComponentModel;
using ClosedXML.Excel;

namespace Frends.Office
{
    /// <summary>
    /// Input for the excel task
    /// </summary>
    public class InputWriteExcel
    {
        /// <summary>
        /// Input csv string.
        /// </summary>
        [DefaultValue("col1;col2\none;two")]
        public string csvString { get; set; }

        /// <summary>
        /// Determines what character will be used for splitting based on cell in csv. Deafult is ';'.
        /// </summary>
        [DefaultValue(';')]
        public char delimiter { get; set; }

        /// <summary>
        /// Determines what string will be used for splitting lines. Default is "\r\n".
        /// </summary>
        [DefaultValue("\r\n")]
        public string lineDelimiter { get; set; }

        /// <summary>
        /// If input csv includes column names (headers). Type boolean.
        /// </summary>
        [DefaultValue(true)]
        public bool includeHeaders { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .xlsx, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        public string path { get; set; }

        /// <summary>
        /// This method parses the input csv string and returns DataTable object.
        /// </summary>
        public System.Data.DataTable ExportToExcel()
        {
            var parsedResult = new List<Dictionary<string, string>>();
            var records = csvString.Split(new string[] { lineDelimiter }, StringSplitOptions.None);

            System.Data.DataTable table = new System.Data.DataTable();
            int recordsPerLine = 0;

            foreach (var record in records)
            {
                recordsPerLine = 0;
                var fields = record.Split(delimiter);
                var recordItem = new Dictionary<string, string>();
                var i = 0;
                foreach (var field in fields)
                {
                    recordItem.Add(i.ToString(), field);
                    i++;
                    recordsPerLine++;
                }
                parsedResult.Add(recordItem);
            }

            if (includeHeaders == true)
            {
                var row = parsedResult[0];
                foreach (var name in row)
                {
                    table.Columns.Add(name.Value, typeof(string));
                }
                parsedResult.RemoveAt(0);
            }
            else
            {
                for (int i = 0; i < recordsPerLine; i++)
                {
                    table.Columns.Add(i.ToString(), typeof(string));
                }
            }

            foreach (Dictionary<string, string> dic in parsedResult)
            {
                System.Data.DataRow workRow = table.NewRow();

                int counter = 0;
                foreach (var y in dic)
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
    /// Office task package for handling files, e.g. Excel.
    /// </summary>
    /// 
    public class WriteExcelTask
    {
        /// <summary>
        /// Allows you to write excel files in .xlsx format. https://github.com/MarcinMichnik-HiQ/Frends.Office
        /// </summary>
        /// <param name="inputWriteExcel"></param>
        /// <returns>Returns true if the file was written to correctly Otherwise throws an exception</returns>
        public static bool WriteExcel(InputWriteExcel inputWriteExcel)
        {
            try
            {
                using (System.Data.DataTable dt = inputWriteExcel.ExportToExcel()) {
                    var workbook = new XLWorkbook();

                    if (inputWriteExcel.includeHeaders == false)
                    {
                        var ws = workbook.Worksheets.Add("Default");
                        ws.FirstRow().FirstCell().InsertData(dt.Rows);
                        ws.Rows().AdjustToContents();
                        ws.Columns().AdjustToContents();
                    }

                    else
                    {
                        var ws = workbook.Worksheets.Add(dt, "Default");
                        ws.Rows().AdjustToContents();
                        ws.Columns().AdjustToContents();
                    }

                    workbook.SaveAs(inputWriteExcel.path);
                    
                }
                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to write file.", ex);
            }
        }
    }
}
