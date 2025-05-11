using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
namespace testexcle.Services
{
    public interface IExcelImportService1
{
    Task<List<Dictionary<string, string>>> ImportExcelAsync(Stream fileStream);
}

public class ExcelImportService1 : IExcelImportService1
{
    public async Task<List<Dictionary<string, string>>> ImportExcelAsync(Stream fileStream)
    {
        var data = new List<Dictionary<string, string>>();

        using (var package = new ExcelPackage(fileStream))
        {
            var worksheet = package.Workbook.Worksheets[0];
            var rowCount = worksheet.Dimension.Rows;
            var colCount = worksheet.Dimension.Columns;

            // Read headers
            var headers = new List<string>();
            for (int col = 1; col <= colCount; col++)
            {
                headers.Add(worksheet.Cells[1, col].Text);
            }

            // Read rows
            for (int row = 2; row <= rowCount; row++)
            {
                var rowData = new Dictionary<string, string>();
                for (int col = 1; col <= colCount; col++)
                {
                    rowData[headers[col - 1]] = worksheet.Cells[row, col].Text;
                }
                data.Add(rowData);
            }
        }

        return data;
    }
}
}