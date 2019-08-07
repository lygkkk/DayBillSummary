using System.Globalization;
using System.Security.Policy;

namespace DayBillSummary.Abstract.Models
{
    public class ReturnFormInfo
    {
        public string Warehouse { get; set; }
        public string CompanyName { get; set; }
        public string OrderNumber { get; set; }
        public int NumberOfSheets { get; set; }
    }
}