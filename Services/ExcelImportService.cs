using OfficeOpenXml;
using testexcle.Models;

namespace testexcle.Services
{
    public class ExcelImportService
    {
        public async Task<List<Product>> ImportAsync(Stream stream)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            using var package = new ExcelPackage(stream);
            var worksheet = package.Workbook.Worksheets.FirstOrDefault();

            var products = new List<Product>();
            if (worksheet == null) return products;

            int rowCount = worksheet.Dimension.Rows;

            for (int row = 2; row <= rowCount; row++) // Skip header
            {
                products.Add(new Product
                {
                    ReceiptDate = worksheet.Cells[row, 1].Text,
                    No = worksheet.Cells[row, 2].Text,
                    MasterCode = worksheet.Cells[row, 3].Text,
                    Secode = worksheet.Cells[row, 4].Text,
                    TrackingNo = worksheet.Cells[row, 5].Text,
                    WeightNo = worksheet.Cells[row, 6].Text,
                    CodNo = worksheet.Cells[row, 7].Text,
                    Note = worksheet.Cells[row, 8].Text
                });
            }

            return products;
        }
    }
}
