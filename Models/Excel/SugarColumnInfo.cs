using KalevaAalto.Models.Excel;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using static KalevaAalto.Static.Main;

namespace KalevaAalto.Models.Excel
{
    /// <summary>
    /// SqlSugar类的信息暂存类
    /// </summary>
    public class ClassColumnInfo
    {
        public string ColumnName { get; set; } = string.Empty;
        public string PropertyName { get; set; } = string.Empty;
        public System.Type Type { get; set; } = typeof(string);

        public DataColumnStyle excelDataColumn => new DataColumnStyle(this.ColumnName, this.Type);

    }
}


namespace KalevaAalto
{
    public static partial class Main
    {
        public static Models.Excel.ClassColumnInfo[] GetClassColumnInfos(this Type entityType)
        {
            List<Models.Excel.ClassColumnInfo> result = new List<Models.Excel.ClassColumnInfo>();

            PropertyInfo[] properties = entityType.GetProperties(BindingFlags.Instance | BindingFlags.Public);

            foreach (PropertyInfo property in properties.Where(it => it.PropertyType.IsStandardDataTableType()))
            {
                var sugarColumnAttribute = property.GetCustomAttribute<SugarColumn>(inherit: true);
                if (sugarColumnAttribute is not null && !sugarColumnAttribute.ColumnName.IsNullOrEmpty())
                {
                    result.Add(new Models.Excel.ClassColumnInfo { ColumnName = sugarColumnAttribute.ColumnName, PropertyName = property.Name, Type = property.PropertyType });
                }
                else
                {
                    result.Add(new Models.Excel.ClassColumnInfo { ColumnName = property.Name, PropertyName = property.Name, Type = property.PropertyType });
                }
            }

            return result.ToArray();
        }



        /// <summary>
        /// 将数组转化为DataTable数据表
        /// </summary>
        /// <typeparam name="T">数组的数据类型</typeparam>
        /// <param name="values">要转化的数组</param>
        /// <param name="tableName">表名</param>
        /// <returns>返回DataTable数据表</returns>
        public static DataTable ToDataTable<T>(T[] values, string tableName = emptyString)
        {
            DataTable result = new DataTable(tableName);

            Type valueType = typeof(T);
            ClassColumnInfo[] sugarColumnInfos = valueType.GetClassColumnInfos();

            //添加字段名称
            foreach (var column in sugarColumnInfos)
            {
                result.Columns.Add(column.ColumnName, column.Type);
            }

            //添加数据
            foreach (var value in values)
            {
                DataRow row = result.Rows.Add();
                foreach (var column in sugarColumnInfos)
                {
                    PropertyInfo propertyInfo = valueType.GetProperty(column.PropertyName)!;
                    object? obj = propertyInfo.GetValue(value);
                    if (obj is null)
                    {
                        row[column.ColumnName] = DBNull.Value;
                    }
                    else
                    {
                        row[column.ColumnName] = obj;
                    }
                }
            }

            return result;
        }


        /// <summary>
        /// 将DataTable转化为数组
        /// </summary>
        /// <typeparam name="T">数组类型</typeparam>
        /// <param name="dataTable">要转化的DataTable</param>
        public static T[] ToArray<T>(this DataTable dataTable) where T : class, new()
        {

            Type objType = typeof(T);
            ClassColumnInfo[] sugarColumnInfos = objType.GetClassColumnInfos();
            HashSet<string> dataTableColumns = dataTable.Columns.Cast<DataColumn>().Select(iterator => iterator.ColumnName).ToHashSet();


            List<T> result = new List<T>();
            foreach (DataRow row in dataTable.Rows)
            {
                T obj = new T();
                foreach (var column in sugarColumnInfos)
                {
                    if (dataTableColumns.Contains(column.ColumnName))
                    {
                        DataColumn dataColumn = dataTable.Columns[column.ColumnName]!;
                        PropertyInfo propertyInfo = objType.GetProperty(column.PropertyName)!;
                        object? valueSource = row[column.ColumnName];
                        object? value = column.Type.GetValue(valueSource);
                        propertyInfo.SetValue(obj, value);
                    }
                }
                result.Add(obj);
            }
            return result.ToArray();
        }
    }
}
