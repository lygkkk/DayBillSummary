using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DayBillSummary.Abstract.IService
{
    public interface INpoiService
    {
        IWorkbook Workbook { get; set; }
        ISheet Sheet { get; set; }
        IRow Row { get; set; }
        ICell Cell { get; set; }
        string FilePath { get; set; }
        string FileSavePath { get; set; }
        XSSFWorkbook WriteData<T>(XSSFWorkbook workbook, string sheetName, T[,] source, int[] location = null);
    }
}