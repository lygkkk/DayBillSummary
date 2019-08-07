using DayBillSummary.Abstract.IService;
using DayBillSummary.EF;
using DayBillSummary.ServiceRealize;
using Microsoft.Extensions.DependencyInjection;

namespace DayBillSummary.DI
{
    public static class Provider
    {
        public static IServiceCollection AddProvider(this IServiceCollection serviceCollection, string dataBase, string sheetName)
        {
            serviceCollection.AddScoped<InjectionPoint>();
            serviceCollection.AddScoped(oledbContext => new OledbContext(dataBase, sheetName));
            serviceCollection.AddScoped<ICompanyInfoService, CompanyInfoService>();
            serviceCollection.AddScoped<INpoiService, NpoiService>();
            return serviceCollection.AddScoped<IOrderInfoService, OrderInfoService>();
        }
    }
}