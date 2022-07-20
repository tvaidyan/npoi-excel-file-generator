using NPOI.XSSF.UserModel;
using System.Data;

namespace ExcelFileGenerationDemo;
public class ExcelFileGenerator
{
    public byte[] Create(DataTable table, string sheetName = "Sheet1")
    {
        byte[] bytes = default!;

        using (var ms = new MemoryStream())
        {
            var workbook = new XSSFWorkbook();
            var excelSheet = workbook.CreateSheet(sheetName);

            var columns = new List<string>();
            var row = excelSheet.CreateRow(0);
            var columnIndex = 0;

            foreach (DataColumn column in table.Columns)
            {
                columns.Add(column.ColumnName);
                row.CreateCell(columnIndex).SetCellValue(column.ColumnName);
                columnIndex++;
            }

            int rowIndex = 1;
            foreach (DataRow dsrow in table.Rows)
            {
                row = excelSheet.CreateRow(rowIndex);
                int cellIndex = 0;
                foreach (string col in columns)
                {
                    row.CreateCell(cellIndex).SetCellValue(dsrow[col].ToString());
                    cellIndex++;
                }

                rowIndex++;
            }

            workbook.Write(ms, true);
            ms.Position = 0;
            bytes = ms.ToArray();
        }

        return bytes;
    }
}
