using System;
using System.Collections.Generic;
using System.ComponentModel;
using ClosedXML.Excel;

namespace Frends.Office
{
    /// <summary>
    /// Input for the excel task
    /// </summary>
    public class Input
    {
        /// <summary>
        /// Input csv string.
        /// </summary>
        public string CsvString { get; set; }

        /// <summary>
        /// Determines what character will be used for splitting based on cell in csv. Deafult is ';'.
        /// </summary>
        public char Delimiter { get; set; }

        /// <summary>
        /// If input csv includes column names (headers). Type boolean.
        /// </summary>
        public bool IncludeHeaders { get; set; }

        /// <summary>
        /// Full path of the target file to be written. File format should be .xlsx, e.g. FileName.xlsx
        /// </summary>
        [DefaultValue(@"c:\temp\file.xlsx")]
        public string Path { get; set; }

        /// <summary>
        /// This method parses the input csv string and returns DataTable object.
        /// </summary>
        public System.Data.DataTable ExportToExcel()
        {
            var parsedResult = new List<Dictionary<string, string>>();
            var records = CsvString.Split('\n');

            foreach (var record in records)
            {
                var fields = record.Split(Delimiter);
                var recordItem = new Dictionary<string, string>();
                var i = 0;
                foreach (var field in fields)
                {
                    recordItem.Add(i.ToString(), field);
                    i++;
                }
                parsedResult.Add(recordItem);
            }

            System.Data.DataTable table = new System.Data.DataTable();

            if (IncludeHeaders == true)
            {
                var row = parsedResult[0];
                foreach (var name in row)
                {
                    table.Columns.Add(name.Value, typeof(string));
                }
                parsedResult.RemoveAt(0);
            }

            foreach (Dictionary<string, string> dic in parsedResult)
            {
                List<string> f = new List<string>();
                foreach (var y in dic)
                {
                    f.Add(y.Value);
                }
                table.Rows.Add(f.ToArray());
            }
            return table;
        }
    }


    /// <summary>
    /// Office task package for handling files, e.g. Excel.
    /// </summary>
    /// 

    public class Office
    {
        /// <summary>
        /// Allows you to write excel files in .xlsx format.
        /// </summary>
        /// <param name="input"></param>
        /// <returns>Returns true if the file was written to correctly Otherwise throws an exception</returns>
        public static bool WriteExcel(Input input)
        {
            try
            {
                using (System.Data.DataTable dt = input.ExportToExcel()) {
                    using (var workbook = new XLWorkbook())
                    {
                        workbook.Worksheets.Add(dt, "Default");
                        workbook.SaveAs(input.Path);
                    }
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
