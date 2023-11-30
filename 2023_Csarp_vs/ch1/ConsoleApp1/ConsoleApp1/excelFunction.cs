using System.Runtime.InteropServices;
using System.Data;
using ExcelDataReader;
using System.Text;

public static class ExcelReader
{
    public static DataTable ReadExcel(string filePath, int sheetIndex)
    {
        // .NET Core에서 추가 인코딩을 지원하기 위해 사용하는 코드
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
        {
            // 한글 깨짐 현상 해결 가능
            ExcelReaderConfiguration config = new ExcelReaderConfiguration();
            config.FallbackEncoding = Encoding.GetEncoding("ks_c_5601-1987");

            using (var reader = ExcelReaderFactory.CreateReader(stream, config))
            {
                var result = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        //Column 자동생성을 무시하고 첫번째 행을 열로 자동 지정.
                        UseHeaderRow = true
                        //UseHeaderRow = false
                    }
                });
                return result.Tables[sheetIndex];
            }
        }
    }
}