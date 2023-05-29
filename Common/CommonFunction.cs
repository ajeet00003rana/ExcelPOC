using ClosedXML.Excel;
using System.Data;

namespace ExcelPOC.Common
{
    public static class CommonFunction
    {
        public static string TrimNullCheck(this string str)
        {
            return !string.IsNullOrEmpty(str) ? str.Trim() : str;
        }

        public static void ExportDataSetWithInBuiltFunction(DataSet ds, string destination)
        {
            DataTable table = ds.Tables[0];


            using (var workbook = new XLWorkbook())
            {
                IXLWorksheet sheet = workbook.Worksheets.Add("Sheet1");
                sheet.Cell(1, 1).InsertTable(table);
                sheet.Column(1).Hide();
                workbook.SaveAs(destination);
            }
        }

        public static void ExportDataSetWithCustomLogic(DataSet ds, string destination)
        {
            DataTable table = ds.Tables[0];

            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add(table.TableName);
                var currentRow = 1;

                // Add columns dynamically 
                for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                {
                    var columnName = table.Columns[columnIndex].ColumnName;
                    worksheet.Cell(currentRow, columnIndex + 1).Value = columnName;
                }

                // Add rows dynamically
                currentRow++; // Move to the row below column headers
                for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++)
                {
                    for (int columnIndex = 0; columnIndex < table.Columns.Count; columnIndex++)
                    {
                        worksheet.Cell(currentRow, columnIndex + 1).Value = table.Rows[rowIndex][columnIndex].ToString();
                    }
                    currentRow++;
                }

                workbook.SaveAs(destination);
            }
        }

    }
}
