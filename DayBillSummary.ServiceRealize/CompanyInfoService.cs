using System.Collections.Generic;
using DayBillSummary.Abstract.IService;
using DayBillSummary.Abstract.Models;
using DayBillSummary.EF;

namespace DayBillSummary.ServiceRealize
{
    public class CompanyInfoService : ICompanyInfoService
    {
        private readonly OledbContext _oledbContext;

        public CompanyInfoService(OledbContext oledbContext)
        {
            _oledbContext = oledbContext;
        }

        public IEnumerable<CompanyInfo> Where(CompanyInfo tEntity)
        {
            string currentDirectory = System.Environment.CurrentDirectory;
            _oledbContext.DataBase = $@"{currentDirectory}\BaseInfo.xlsx";
            _oledbContext.SheetName = "CompanyInfo";
            return _oledbContext.CompanyInfos.Where(tEntity);
        }
    }
}