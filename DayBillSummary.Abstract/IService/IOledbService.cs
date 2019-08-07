using System.Collections.Generic;

namespace DayBillSummary.Abstract.IService
{
    public interface IOledbService<TEntity> where TEntity : class, new()
    {
        string DataBase { get; set; }
        string Provider { get; set; }
        string SheetName { get; set; }
        IEnumerable<TEntity> Where(TEntity tEntity);
    }
}