using System.Reflection;
using System.Runtime.CompilerServices;
using DayBillSummary.Abstract.IService;
using DayBillSummary.Abstract.Models;

namespace DayBillSummary.EF
{
    public class OledbContext
    {
        public string DataBase { get; set; }
        public string SheetName { get; set; }

        public IOledbService<OrderInfo> OrderInfos => DbSet<OrderInfo>();
        public IOledbService<CompanyInfo> CompanyInfos => DbSet<CompanyInfo>();
        public OledbContext(string dataBase, string sheetName)
        {
            DataBase = dataBase;
            SheetName = sheetName;
        }

        private IOledbService<TEntity> DbSet<TEntity>() where TEntity : class, new()
        {
            return new OledbService<TEntity>(DataBase, SheetName);
        }
    }
}