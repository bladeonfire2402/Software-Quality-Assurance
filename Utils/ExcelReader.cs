using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;

public static class ExcelReader
{
    //private static int sheetCount;
    public static IEnumerable<object[]> GetExcelData(string filePath,int sheetCount)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        var data = new List<object[]>();

        using (var package = new ExcelPackage(new FileInfo(filePath)))
        {
            var worksheet = package.Workbook.Worksheets[sheetCount]; // Sheet1
            int rowCount = 9;

            for (int row = 2; row <= rowCount; row++)
            {
                string soA = worksheet.Cells[row, 2].Text;
                string soB = worksheet.Cells[row, 3].Text;
                string expectedResult = worksheet.Cells[row, 4].Text;
                string actualResult = worksheet.Cells[row, 5].Text;
                int rowIndex = row; // Lưu chỉ số dòng để ghi lại sau

                data.Add(new object[] { soA, soB, expectedResult, actualResult, rowIndex });
            }
        }

        return data;
    }

    public static void WriteTestResult(string filePath, int rowIndex, string result, int sheetCount)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Mở file Excel để ghi
        FileInfo file = new FileInfo(filePath);
        using (var package = new ExcelPackage(file))
        {
            var worksheet = package.Workbook.Worksheets[sheetCount];
            worksheet.Cells[rowIndex, 6].Value = result;
            package.Save();
        }
    }

    public static void WriteActualResult(string filePath, int rowIndex, string actualResult, int sheetCount)
    {
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Mở file Excel để ghi
        FileInfo file = new FileInfo(filePath);
        using (var package = new ExcelPackage(file))
        {
            var worksheet = package.Workbook.Worksheets[sheetCount];
            worksheet.Cells[rowIndex, 5].Value = actualResult;
            package.Save();
        }
    }
}