using System.IO;
using System.Text;

namespace DayBillSummary.FileOperational
{
    public static class LoadExcelLocation
    {
        public static string ReadJsconfig(string filePath)
        {
            string str = "";
            StreamReader streamReader = new StreamReader(filePath, Encoding.Default);
            string line;
            while ((line = streamReader.ReadLine()) != null)
            {
                str += line;
            }
            return str;
        }
    }
}