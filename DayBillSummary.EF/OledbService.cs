using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Linq;
using System.Text;
using DayBillSummary.Abstract.IService;

namespace DayBillSummary.EF
{
    public class OledbService<TEntity> : IOledbService<TEntity> where TEntity: class, new()
    {
        public string DataBase { get; set; }
        public string Provider { get; set; } = @"Provider=Microsoft.Ace.OleDb.12.0;Extended Properties=Excel 12.0;Data Source=";
        public string SheetName { get; set; }
        public OledbService(string dataBase, string sheetName)
        {
            DataBase = dataBase;
            SheetName = sheetName;
        }


        #region 获取数据

        /// <summary>
        /// 获取数据
        /// </summary>
        /// <param name="tEntity">实体</param>
        /// <returns></returns>
        public IEnumerable<TEntity> Where(TEntity tEntity)
        {
            var sql = $"SELECT {GenerateCommand(tEntity)} FROM [{SheetName}$]";
            DataTable dt = new DataTable();
            using (OleDbConnection connection = new OleDbConnection(Provider + DataBase))
            {
                using (OleDbDataAdapter dataAdapter = new OleDbDataAdapter(sql, connection))
                {
                    dataAdapter.Fill(dt);
                }
            }

            return DataTableToEntity(dt);
        }

        #endregion

        #region 转换sql命令

        private string GenerateCommand(TEntity tEntiy)
        {
            StringBuilder sb = new StringBuilder();

            tEntiy.GetType().GetProperties().ToList().ForEach(
                e => sb.Append($"{e.Name},")
                );
            return sb.ToString().TrimEnd(',');
        }

        #endregion


        #region DataTable转list实体
        /// <summary>
        /// DataTable转DataTable转list实体
        /// </summary>
        /// <param name="dt">DataTable</param>
        /// <returns></returns>
        private IEnumerable<TEntity> DataTableToEntity(DataTable dt)
        {
            List<TEntity> enties = new List<TEntity>();

            foreach (DataRow row in dt.Rows)
            {
                var entity = new TEntity();
                var properties = entity.GetType().GetProperties();

                foreach (var property in properties)
                {
                    var type = property.GetType();
                    //property.SetValue(orderInfo, row[property.Name].GetType());

                    property.GetSetMethod().Invoke(entity,
                        new[] { Convert.ChangeType(row[property.Name], property.PropertyType) });
                    //property.SetValue(property, row.GetType().GetProperty(property.Name).GetValue(property));

                }
                enties.Add(entity);
            }

            return enties;
        }

        #endregion
    }
}