using System;
using System.Collections.Generic;
using System.ComponentModel;
using Microsoft.Office.Interop.Excel;

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
        /// Determines what character will be used for splitting based on line in csv. Deafult is '\n'.
        /// </summary>
        [DefaultValue('\n')]
        public char LineDelimiter { get; set; }

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
            var records = CsvString.Split(LineDelimiter);

            foreach (var record in records)
            {
                var fields = record.Split(this.Delimiter);
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
    /// Options for the Write Task
    /// </summary>
    public class WriteOptions
    {
        /// <summary>
        /// What should happen if the file already exist
        /// </summary>
        [DefaultValue(true)]
        public bool OverWrite { get; set; }
    }

    /// <summary>
    /// Example task package for handling files
    /// </summary>
    /// 

    public class Office
    {
        /// <summary>
        /// Allows you to write excel files in .xlsx format.
        /// </summary>
        /// <param name="input"></param>
        /// <param name="options"></param>
        /// <returns>Returns true if the file was written to correctly Otherwise throws an exception</returns>
        public static bool WriteExcel(Input input, WriteOptions options)
        {
            Application excel = new Application();
            Workbook workbook = excel.Workbooks.Add(Type.Missing);
            Worksheet worksheet = (Worksheet)workbook.ActiveSheet;
            Range cellrange;
            try
            {
                worksheet.Name = "Default";

                System.Data.DataTable d = input.ExportToExcel();

                int rowcount = 0;

                foreach (System.Data.DataRow datarow in d.Rows)
                {
                    for (int i = 1; i <= d.Columns.Count; i++)
                    {
                        worksheet.Cells[rowcount, i] = datarow[i - 1].ToString();

                        if (rowcount > 3)
                        {
                            if (i == d.Columns.Count)
                            {
                                if (rowcount % 2 == 0)
                                {
                                    cellrange = worksheet.Range[worksheet.Cells[rowcount, 1], worksheet.Cells[rowcount, d.Columns.Count]];
                                }
                            }
                        }
                    }
                }
                cellrange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, d.Columns.Count]];
                cellrange.EntireColumn.AutoFit();

                Borders border = cellrange.Borders;
                border.LineStyle = XlLineStyle.xlContinuous;
                border.Weight = 2d;

                cellrange = worksheet.Range[worksheet.Cells[1, 1], worksheet.Cells[2, d.Columns.Count]];

                if (options.OverWrite == true)
                {
                    excel.DisplayAlerts = false;
                    workbook.SaveAs(input.Path, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing,
                    Type.Missing, Type.Missing);
                }
                else
                {
                    workbook.SaveAs(input.Path);
                }

                workbook.Close();
                excel.Quit();

                return true;
            }
            catch (Exception ex)
            {
                throw new Exception("Unable to write file.", ex);
            }
            finally
            {
                worksheet = null;
                cellrange = null;
                workbook = null;
            }
        }
    }
}
