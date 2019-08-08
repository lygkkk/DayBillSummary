using DayBillSummary.DI;
using Microsoft.Extensions.DependencyInjection;
using System;
using System.IO;

namespace DayBillSummary.App
{
    class Program
    {
        private static ServiceProvider _serviceProvider;
        static void Main(string[] args)
        {
            string dataBase = $@"{Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory)}\单据审核中心.xls";

            while (!File.Exists(dataBase))
            {
                Console.WriteLine(@"单据审核中心文件不存在。\r\n请导出后按任意键继续。");
                Console.ReadKey();
            }

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
