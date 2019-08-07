namespace DayBillSummary.Abstract.Models
{
    public class Location<T>
    {
        public string Category { get; set; }
        public T[][] Data { get; set; }
    }
}