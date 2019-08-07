namespace DayBillSummary.Abstract.Models
{
    public class OrderInfo
    {
        public string Category { get; set; }
        public string OrderNumber { get; set; }
        public string CompanyName { get; set; }
        public string WareHouse { get; set; }
        public int NumberOfSheets { get; set; }
    }
}