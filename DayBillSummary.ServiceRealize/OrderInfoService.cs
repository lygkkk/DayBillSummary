using System.Collections.Generic;
using DayBillSummary.Abstract.IService;
using DayBillSummary.Abstract.Models;
using DayBillSummary.EF;

namespace DayBillSummary.ServiceRealize
{
    public class OrderInfoService : IOrderInfoService
    {
        private readonly OledbContext _oledbContext;
        public OrderInfoService(OledbContext oledbContext)
        {
            _oledbContext = oledbContext;
        }
        public IEnumerable<OrderInfo> Where(OrderInfo orderInfo)
        {
            return _oledbContext.OrderInfos.Where(orderInfo);
        }
    }
}