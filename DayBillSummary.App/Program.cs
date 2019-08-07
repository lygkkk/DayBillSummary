using DayBillSummary.DI;
using Microsoft.Extensions.DependencyInjection;
using System;

namespace DayBillSummary.App
{
    class Program
    {
        private static ServiceProvider _serviceProvider;
        static void Main(string[] args)
        {
            string dataBase = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\单据审核中心.xls";
            IServiceCollection serviceCollection = new ServiceCollection();
            serviceCollection.AddProvider(dataBase, "单据审核中心");

            _serviceProvider = serviceCollection.BuildServiceProvider();

            var context = _serviceProvider.GetService<InjectionPoint>();
            context.MessageEvent += SendMessage;
            context.T1();

            Console.WriteLine("单据汇总完毕，按任意键退出。");
            Console.ReadKey();
        }

        static void SendMessage(string msg)
        {
            Console.WriteLine(msg);
        }
    }
}
