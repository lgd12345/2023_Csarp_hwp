using System.Runtime.InteropServices;
using Excel = Microsoft.Office.Interop.Excel;
using System.Data;

public static class ExcelReader
{
    public static Excel.Range ReadExcelFile(string Ex_filePath, string rangeAdd)
    {
        Excel.Application excelApp = new Excel.Application();
        Excel.Workbook workbook = excelApp.Workbooks.Open(Ex_filePath);
        Excel.Worksheet worksheet = (Excel.Worksheet)workbook.Sheets[1];
        Excel.Range range = worksheet.Range[rangeAdd];

        // 결과 사용
        PrintExcelRange(range);

        static void PrintExcelRange(Excel.Range range)
        {
            object[,] values = range.Value;

            for (int row = 1; row <= values.GetLength(0); row++)
            {
                for (int col = 1; col <= values.GetLength(1); col++)
                {
                    Console.Write($"{values[row, col]}\t");
                }
                Console.WriteLine();
            }
        }

        // Excel 객체 해제
        Marshal.ReleaseComObject(worksheet);
        workbook.Close();
        Marshal.ReleaseComObject(workbook);
        excelApp.Quit();
        Marshal.ReleaseComObject(excelApp);

        return range;
    }
}