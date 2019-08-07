using System.Collections.Generic;

namespace DayBillSummary.Abstract.IService
{
    public interface IBaseService<TEntity> where TEntity: class, new()
    {
        IEnumerable<TEntity> Where(TEntity tEntity);
    }
}