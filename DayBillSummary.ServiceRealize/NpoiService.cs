using System.IO;
using DayBillSummary.Abstract.IService;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace DayBillSummary.ServiceRealize
{
    public class NpoiService : INpoiService
    {
        public IWorkbook Workbook { get; set; }
        public ISheet Sheet { get; set; }
        public IRow Row { get; set; }
        public ICell Cell { get; set; }
        public string FilePath { get; set; }
        public string FileSavePath { get; set; }
        public NpoiService()
        {
        }

        public NpoiService(string filePath, string fileSavePath)
        {
            FilePath = filePath;
            FileSavePath = fileSavePath;
        }

        public XSSFWorkbook WriteData<T>(XSSFWorkbook workbook, string sheetName, T[,] source, int[] location = null)
        {
            Sheet = workbook.GetSheet(sheetName);

            int startRow = 0;
            int endRow = 0;
            int startCol = 0;
            int endCol = 0;

            if (location != null)
            {
                startRow = location[0];
                endRow = location[1];
                startCol = location[2];
                endCol = location[3];
            }

            for (int i = 0; i < source.GetLength(0); i++)
            {
                Row = Sheet.GetRow(i + startRow);
                for (int j = 0; j < source.GetLength(1); j++)
                {
                    Row.Cells[j + startCol].SetCellValue(source[i, j].ToString());
                }
            }
            return workbook;
        }
    }
}