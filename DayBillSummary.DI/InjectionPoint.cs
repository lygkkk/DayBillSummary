using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using DayBillSummary.Abstract.IService;
using DayBillSummary.Abstract.Models;
using DayBillSummary.Extensions;
using Newtonsoft.Json;
using NPOI.XSSF.UserModel;

namespace DayBillSummary.DI
{
    public class InjectionPoint
    {
        private readonly IOrderInfoService _orderInfoService;
        private readonly ICompanyInfoService _companyInfoService;
        private readonly INpoiService _npoiService;

        private XSSFWorkbook workbook;
        private List<Location<int>> _exceLocations = null;
        private string _deskTopPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        private string _currentDirectory = System.Environment.CurrentDirectory;

        public delegate void SenderMessage(string message);
        public event SenderMessage MessageEvent;
        public InjectionPoint(IOrderInfoService orderInfoService, ICompanyInfoService companyInfoService, INpoiService npoiService)
        {
            _orderInfoService = orderInfoService;
            _companyInfoService = companyInfoService;
            _npoiService = npoiService;
        }

        public void T1()
        {
            MessageEvent?.Invoke("获取单据信息。");
            var orderInfos = _orderInfoService.Where(new OrderInfo());
            var companyInfos = _companyInfoService.Where(new CompanyInfo());
 
            #region 二次分类汇总 根据单据分类和公司名称

            //二级分类汇总 根据单据分类和公司名称
            //var groupingOrderInfos = ReplaceCompanyName(orderInfos, companyInfos).GroupBy(o =>
            //{
            //    if (o.Category == "销售出库单") return new { a = "销售出库单", o.CompanyName };
            //    if (o.Category == "销售退货单") return new { a = "销售退货单", o.CompanyName };
            //    return new { a = "其他单据", o.CompanyName };
            //});

            #endregion

            #region 自定义键对值分类汇总 

            //var groupingOrderInfos = ReplaceCompanyName(orderInfos, companyInfos).GroupBy(o =>
            //{
            //    if (o.Category == "销售出库单")return "销售出库单";
            //    if (o.Category == "销售退货单") return "销售退货单";
            //    return "其他单据";
            //}, elementSelector => elementSelector, (key, entity) => new {key, entity});

            #endregion

            #region 一级分类汇总

            //一级分类汇总
            var groupingOrderInfos = ReplaceCompanyName(orderInfos, companyInfos).GroupBy(o =>
            {
                if (o.Category == "销售出库单") return "销售出库单";
                if (o.Category == "销售退货单") return "销售退货单";
                return "其他单据";
            });

            #endregion

            MessageEvent?.Invoke("获取汇总单文件流。");
            using (FileStream fileStream = File.OpenRead("每日汇总单.xlsx"))
            {
                workbook = new XSSFWorkbook(fileStream);
            }

            foreach (var key in groupingOrderInfos)
            {
                if (key.Key == "销售出库单")
                {
                    var salesOrderInfos = key.GroupBy(o => o.CompanyName).Select(g => 
                        new SalesOrderInfo{
                            CompanyName = g.First(s => true).CompanyName,
                            NumberOfOrder =g.Select(s => int.Parse(s.OrderNumber)).OrderNumberCombine(),
                            NumberOfSheets = g.Sum(s => s.NumberOfSheets)
                        });
                    WriteExce(salesOrderInfos, key.Key);
                }
                else if(key.Key == "销售退货单")
                {
                    var returnFormInfos = key.GroupBy(o => new {o.WareHouse, o.CompanyName}).Select(g =>
                        new ReturnFormInfo
                        {
                            CompanyName = g.First(r => true).CompanyName,
                            NumberOfSheets = g.Sum(r => r.NumberOfSheets),
                            OrderNumber = g.Select(r => int.Parse(r.OrderNumber)).OrderNumberCombine(),
                            Warehouse = g.First(r => true).WareHouse
                        });
                    WriteExce(returnFormInfos, key.Key);
                }
                else
                {
                    var otherOrderInfos = key.GroupBy(o => new {o.Category, o.CompanyName}).Select(g =>
                        new OtherOrderInfo
                        {
                            CompanyName = g.First(o => true).CompanyName,
                            Category = g.First(o => true).Category,
                            OrderNumber = g.Select(o => int.Parse(o.OrderNumber)).OrderNumberCombine(),
                            OrderOfSheets = g.Sum(o => o.NumberOfSheets)
                        });
                    WriteExce(otherOrderInfos, key.Key);
                }
            }
            MessageEvent?.Invoke("生成汇总单。");
            using (FileStream fileStream = File.OpenWrite(_deskTopPath + @"\每日汇总单.xlsx"))
            {
                workbook.Write(fileStream);
            }
        }

        #region 写入Excel

        private void WriteExce<T>(IEnumerable<T> source, string sheetName) where T: class, new()
        {
            MessageEvent?.Invoke("写入Excel。");
            if (_exceLocations == null)
            {
                _exceLocations = new List<Location<int>>();
                var str = FileOperational.LoadExcelLocation.ReadJsconfig("jsconfig1.json");
                _exceLocations = JsonConvert.DeserializeObject<List<Location<int>>>(str);
            }
           
            var location = _exceLocations.Where(e => e.Category == sheetName).Select(e => e.Data).ToList();

            int count = source.Count() < 17 ? 1 : (int)Math.Ceiling(source.Count() / 16.00);
            int skip = 0;

            for (int i = 0; i < count; i++)
            {
                var data = ListToArray(source.Skip(skip).ToArray());

                _npoiService.WriteData(workbook, sheetName, data, location[0][i]);
                skip += 16;
                
            }
        }

        #endregion

        #region ListToArray

        private string[,] ListToArray<T>(IEnumerable<T> source)
        {
            var properties = typeof(T).GetProperties();
            string[,] data = new string[source.Count(), properties.Length];

            int lbound = 0;
            int ubound;

            foreach (var entity in source)
            {
                ubound = 0;
                foreach (var property in properties)
                {
                    var tmp = entity.GetType().GetProperty(property.Name).GetValue(entity).ToString();
                    data[lbound, ubound] = tmp;
                    ubound++;
                }

                lbound++;
            }

            return data;
        }

        #endregion

        #region 替换单位名称-计算张数-截取单号
        /// <summary>
        /// 替换单位名称-计算张数-截取单号
        /// </summary>
        /// <param name="orderInfos"></param>
        /// <param name="companyInfos"></param>
        private IEnumerable<OrderInfo> ReplaceCompanyName(IEnumerable<OrderInfo> orderInfos, IEnumerable<CompanyInfo> companyInfos)
        {

            var tmp = orderInfos.Select(o =>
            {
                //截取单号后5位
                o.OrderNumber = o.OrderNumber.Substring(10);
                //计算张数 ceiling 小数点后加1
                o.NumberOfSheets = (int)Math.Ceiling((o.NumberOfSheets + 5) / 19.00);
                ////修改单据类型非销售出库单和销售退货单为其他单据
                //if (o.Category != "销售出库单" || o.Category != "销售退货单") o.Category = "其他单据";
                //判断公司名称是否为空
                if (string.IsNullOrEmpty(o.CompanyName)) return o;
                
                var companyShortName = companyInfos.FirstOrDefault(c => c.CompanyFullName == o.CompanyName)?.CompanyShortName;

                if (companyShortName != null)
                {
                    o.CompanyName = companyShortName;
                }

                return o;

            });

            return tmp;
        }

        #endregion
    }
}