
using Microsoft.Office.Interop.Excel;

namespace selenium.Utils
{
    public  class ExcelHelpers
    {
        Workbook workbook;
        Worksheet worksheet;

        public void ExcelPage(int page)
        {
            worksheet = workbook.Sheets[page];
        }


    }
}
