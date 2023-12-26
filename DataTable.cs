using MySqlX.XDevAPI.Relational;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using OfficeOpenXml;
using Org.BouncyCastle.Math.EC;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using NetTaste;
using System.Text.RegularExpressions;
using System.Drawing;

namespace KalevaAalto
{
    public static partial class Main
    {

        
        public static Dictionary<TKey, int> ToSortDictionary<TKey>(this TKey[] values) where TKey : notnull
        {
            return values.Select((item, index) => new { item, index }).ToDictionary(it => it.item, it => it.index);
        }

        /// <summary>
        /// SqlSugar类的信息暂存类
        /// </summary>
        public class SugarColumnInfo
        {
            public string columnName { get; set; } = string.Empty;
            public string propertyName { get; set; } = string.Empty;
            public System.Type type { get; set; } = typeof(string);

            public ExcelDataColumn excelDataColumn
            {
                get
                {
                    return new ExcelDataColumn(this.columnName,this.type);
                }
            }
        }



        public static SugarColumnInfo[] GetSugarColumns(this Type entityType)
        {
            List<SugarColumnInfo> result = new List<SugarColumnInfo>();

            PropertyInfo[] properties = entityType.GetProperties(BindingFlags.Instance | BindingFlags.Public);

            foreach (PropertyInfo property in properties.Where(it => it.PropertyType.IsStandardDataTableType()))
            {
                var sugarColumnAttribute = property.GetCustomAttribute<SugarColumn>(inherit: true);
                if (sugarColumnAttribute is not null && !sugarColumnAttribute.ColumnName.IsNullOrEmpty())
                {
                    result.Add(new SugarColumnInfo { columnName = sugarColumnAttribute.ColumnName, propertyName = property.Name, type = property.PropertyType });
                }
                else
                {
                    result.Add(new SugarColumnInfo { columnName = property.Name,propertyName = property.Name,type=property.PropertyType });
                }
            }

            return result.ToArray();
        }



        public static string GetSugarTableName(this Type entityType)
        {
            var sugarTableAttribute = entityType.GetCustomAttributes(typeof(SugarTable), true)
                                           .FirstOrDefault() as SugarTable;

            if (sugarTableAttribute is not null)
            {
                return sugarTableAttribute.TableName;
            }
            else
            {
                return entityType.Name;
            }
        }



        /// <summary>
        /// 获取DataTable对象的csv格式字符串
        /// </summary>
        /// <param name="dataTable">要生成csv格式字符串</param>
        /// <param name="maxCount">要限制生成的csv格式字符串的数据行数</param>
        /// <returns>返回dataTable的csv格式字符串</returns>
        public static string CSV(this DataTable dataTable, int maxCount = int.MaxValue)
        {
            //检查DataTable是否为空表，是的话就返回空字符串
            if (dataTable.Columns.Count <= 0)
            {
                return string.Empty;
            }


            //创建要生成csv字符串的可变字符串对象
            StringBuilder result = new StringBuilder();

            //录入字段
            foreach (DataColumn column in dataTable.Columns)
            {
                result.Append(column.ColumnName);
                result.Append(',');
            }
            result.Remove(result.Length - 1, 1);
            result.Append('\n');


            //录入数据
            int counter = 0; //定义计数器
            foreach (DataRow dataRow in dataTable.Rows)
            {
                if (counter > maxCount)
                {
                    break;//如果计数器大于所限制的行数，则退出
                }
                foreach (object? cell in dataRow.ItemArray)
                {
                    result.Append(cell);
                    result.Append(',');
                }
                result.Remove(result.Length - 1, 1);
                result.Append('\n');
                counter++;
            }
            result.Remove(result.Length - 1, 1);



            return result.ToString();
        }




        /// <summary>
        /// 检查DataTable对象是否含有名为fieldName的字段，是的话就返回true，否则返回false
        /// </summary>
        /// <param name="dataTable">要检查的DataTable对象</param>
        /// <param name="columnName">要检查的字段名称</param>
        public static bool IsExistsColumn(this DataTable dataTable, string columnName)
        {
            foreach (DataColumn dataColumn in dataTable.Columns)
            {
                if (dataColumn.ColumnName == columnName)
                {
                    return true;
                }
            }
            return false;
        }


        /// <summary>
        /// 检查DataTable对象是否含有名为fieldNames中字符名称的所有字段，是的话就返回true，否则返回false
        /// </summary>
        /// <param name="dataTable">要检查的DataTable对象</param>
        /// <param name="columnNames">要检查的字段名称数组</param>
        public static bool IsExistsColumn(this DataTable dataTable, string[] columnNames)
        {

            if (dataTable.Columns.Count < columnNames.Length)
            {
                return false;
            }

            HashSet<string> dataTableColumnNamesHashSet = new HashSet<string>();
            foreach (DataColumn dataColumn in dataTable.Columns)
            {
                dataTableColumnNamesHashSet.Add(dataColumn.ColumnName);
            }

            foreach (string columnName in columnNames)
            {
                if (!dataTableColumnNamesHashSet.Contains(columnName))
                {
                    return false;
                }
            }


            return true;
        }




        /// <summary>
        /// 以某种描述，筛选DataTable中的行
        /// </summary>
        /// <param name="dataTable">要筛选的DataTable对象</param>
        /// <param name="conditions">筛选描述字符串</param>
        /// <returns>返回DataTable对象筛选后的DataTable对象</returns>
        public static DataTable TableSelect(this DataTable dataTable, string conditions)
        {
            DataRow[] dataRows = dataTable.Select(conditions);
            DataTable result = dataTable.Clone();
            foreach (DataRow row in dataRows)
            {
                result.Rows.Add(row.ItemArray);
            }
            return result;
        }



        /// <summary>
        /// 以某种描述，筛选DataTable中的行，并选择列
        /// </summary>
        /// <param name="dataTable">要筛选的DataTable对象</param>
        /// <param name="column_names">要选择的列</param>
        /// <param name="conditions">筛选描述字符串</param>
        /// <returns>返回DataTable对象筛选后的DataTable对象</returns>
        public static DataTable TableSelect(this DataTable dataTable, string[] column_names, string? conditions = null)
        {
            DataTable result = new DataTable();

            //筛选
            if (!string.IsNullOrEmpty(conditions))
            {
                dataTable = dataTable.TableSelect(conditions);
            }


            //获取字段名称列表
            HashSet<string> strings = dataTable.Columns.Cast<DataColumn>().Select(it => it.ColumnName).ToHashSet();
            foreach (string column_name in column_names)
            {
                if (strings.Contains(column_name))
                {
                    result.Columns.Add(column_name, dataTable.Columns[column_name]!.DataType);
                }
            }

            //录入数据
            foreach (DataRow row in dataTable.Rows)
            {
                DataRow newRow = result.Rows.Add();
                foreach (DataColumn column in result.Columns)
                {
                    newRow[column.ColumnName] = row[column.ColumnName];
                }
            }

            return result;
        }


        /// <summary>
        /// 将某个字符串或者说数字汇总
        /// </summary>
        /// <param name="dataTable">要汇总的DataTable对象</param>
        /// <param name="column_name">要汇总的列</param>
        /// <param name="conditions">筛选描述字符串</param>
        /// <returns>返回汇总后的结果</returns>
        public static object? Sum(this DataTable dataTable, string column_name, string? conditions = null)
        {
            if (!dataTable.IsExistsColumn(column_name))
            {
                return null;
            }


            DataRow[] dataRows;
            if (conditions is null)
            {
                dataRows = dataTable.Rows.Cast<DataRow>().ToArray();
            }
            else
            {
                dataRows = dataTable.Select(conditions);
            }


            Type type = dataTable.Columns[column_name]!.DataType;
            if (type == typeof(short) || type == typeof(short?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0;
                    }
                    else
                    {
                        return (short)cell;
                    }
                });
            }
            else if (type == typeof(int) || type== typeof(int?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0;
                    }
                    else
                    {
                        return (int)cell;
                    }
                });
            }
            else if (type == typeof(long) || type == typeof(long?) || type == typeof(ulong) || type == typeof(ulong?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0L;
                    }
                    else
                    {
                        return (long)cell;
                    }
                });
            }
            else if (type == typeof(byte) || type == typeof(byte?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0;
                    }
                    else
                    {
                        return (byte)cell;
                    }
                });
            }
            else if (type == typeof(ushort) || type == typeof(ushort?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0;
                    }
                    else
                    {
                        return (ushort)cell;
                    }
                });
            }
            else if (type == typeof(uint) || type == typeof(uint?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0;
                    }
                    else
                    {
                        return (uint)cell;
                    }
                });
            }
            else if (type == typeof(float) || type == typeof(float?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0F;
                    }
                    else
                    {
                        return (float)cell;
                    }
                });
            }
            else if (type == typeof(double) || type == typeof(double?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0D;
                    }
                    else
                    {
                        return (double)cell;
                    }
                });
            }
            else if (type == typeof(decimal) || type == typeof(decimal?))
            {
                return dataRows.Sum(it =>
                {
                    var cell = it[column_name];
                    if (cell is null || cell is DBNull)
                    {
                        return 0M;
                    }
                    else
                    {
                        return (decimal)cell;
                    }
                });
            }
            else if (type == typeof(string))
            {
                StringBuilder rs = new StringBuilder();

                foreach (DataRow row in dataRows)
                {
                    rs.Append((string)row[column_name]);
                    rs.Append(',');
                }
                if (rs.Length > 0)
                {
                    rs.Remove(rs.Length - 1, 1);
                }

                return rs.ToString();
            }


            return null;
        }



        /// <summary>
        /// 以某种描述，对DataTable中的行进行排序
        /// </summary>
        /// <param name="dataTable">要排序的DataTable对象</param>
        /// <param name="conditions">排序描述字符串</param>
        /// <returns>返回DataTable对象排序后的DataTable对象</returns>
        public static DataTable Sort(this DataTable dataTable, string conditions)
        {
            dataTable.DefaultView.Sort = conditions;
            return dataTable.DefaultView.ToTable();
        }

        /// <summary>
        /// 以某种描述，对DataTable中的行进行排序
        /// </summary>
        /// <param name="dataTable">要排序的DataTable对象</param>
        /// <param name="comparison">排序描述</param>
        /// <returns>返回DataTable对象排序后的DataTable对象</returns>
        public static DataTable Sort(this DataTable dataTable, Comparison<DataRow> comparison)
        {
            DataTable newDt = dataTable.Clone();
            List<DataRow> dataRows = new List<DataRow>();
            foreach (DataRow row in dataTable.Rows)
            {
                dataRows.Add(row);
            }
            dataRows.Sort(comparison);
            foreach (DataRow row in dataRows)
            {
                newDt.Rows.Add(row.ItemArray);
            }
            return newDt;
        }


        /// <summary>
        /// 找到数据表中某个字段的序号
        /// </summary>
        /// <param name="dataTable">数据表</param>
        /// <param name="dataColumnName">字段名称</param>
        /// <returns>返回dataTable中名为dataColumnName字段的序号，找不到就返回notfound</returns>
        public static int FindColumn(this DataTable dataTable, string dataColumnName)
        {
            int dataColumnCount = dataTable.Columns.Count;
            for (int i = 0; i < dataColumnCount; i++)
            {
                if (dataTable.Columns[i].ColumnName == dataColumnName)
                {
                    return i;
                }
            }
            return notfound;
        }



        /// <summary>
        /// 将数组转化为DataTable数据表
        /// </summary>
        /// <typeparam name="T">数组的数据类型</typeparam>
        /// <param name="values">要转化的数组</param>
        /// <param name="tableName">表名</param>
        /// <returns>返回DataTable数据表</returns>
        public static DataTable ToDataTable<T>(T[] values,string tableName = emptyString)
        {
            DataTable result = new DataTable(tableName);

            Type valueType = typeof(T);
            SugarColumnInfo[] sugarColumnInfos = valueType.GetSugarColumns();

            //添加字段名称
            foreach(var column in sugarColumnInfos)
            {
                result.Columns.Add(column.columnName, column.type);
            }

            //添加数据
            foreach(var value in values)
            {
                DataRow row = result.Rows.Add();
                foreach (var column in sugarColumnInfos)
                {
                    PropertyInfo propertyInfo = valueType.GetProperty(column.propertyName)!;
                    object? obj = propertyInfo.GetValue(value);
                    if(obj is null)
                    {
                        row[column.columnName] = DBNull.Value;
                    }
                    else
                    {
                        row[column.columnName] = obj;
                    }
                }
            }

            return result;
        }



        #region ExcelDataColumn

        /// <summary>
        /// 包含成员列名、水平对齐、垂直对齐、宽度、数字格式、是否汇总
        /// </summary>
        public class ExcelDataColumn
        {
            public string columnName { get;private set; }

            private ExcelHorizontalAlignment? _horizontalAlignment = null;
            public ExcelHorizontalAlignment horizontalAlignment
            {
                get
                {
                    if(this._horizontalAlignment is null)
                    {
                        if (this.type.IsNumber())
                        {
                            return ExcelHorizontalAlignment.Right;
                        }
                        else if (this.type.IsByteArray())
                        {
                            return ExcelHorizontalAlignment.Left;
                        }
                        else
                        {
                            return ExcelHorizontalAlignment.Center;
                        }
                    }
                    else
                    {
                        return this._horizontalAlignment.Value;
                    }
                }
                set
                {
                    this._horizontalAlignment = value;
                }
            }
            private ExcelVerticalAlignment? _verticalAlignment = null;
            public ExcelVerticalAlignment verticalAlignment
            {
                get
                {
                    if(this._verticalAlignment is null)
                    {
                        return ExcelVerticalAlignment.Center;
                    }
                    else
                    {
                        return this._verticalAlignment.Value;
                    }
                }
                set
                {
                    this._verticalAlignment = value;
                }
            }
            private int? _width = null;
            public int width
            {
                get
                {
                    if(this._width is null)
                    {

                        if (type.IsOrNullableChar() || type == typeof(byte) || type == typeof(byte?))
                        {
                            return 6;
                        }
                        else if (type.IsOrNullableBool())
                        {
                            return 8;
                        }
                        else if (type.IsOrNullableNumber() || type.IsOrNullableDateTime())
                        {
                            return 12;
                        }
                        else if (type.IsByteArray())
                        {
                            return 30;
                        }
                        else
                        {
                            return 10;
                        }
                    }
                    else
                    {
                        return this._width.Value.Around(6,100);
                    }
                }
                set
                {
                    this._width = value;
                }
            }

            private string? _numberFormat = null;
            public string numberFormat 
            {
                get
                {
                    if(this._numberFormat is null)
                    {
                        if (this.type.IsOrNullableInteger() || this.type.IsOrNullableUInteger())
                        {
                            return NumberFormat.Int;
                        }
                        else if(this.type.IsOrNullableFloat() || this.type.IsOrNullableDecimal())
                        {
                            return NumberFormat.Decimal;
                        }
                        else if (this.type.IsOrNullableDateTime())
                        {
                            return NumberFormat.Date;
                        }
                        else
                        {
                            return string.Empty;
                        }
                    }
                    else
                    {
                        return this._numberFormat;
                    }
                }
                set
                {
                    this._numberFormat = value;
                }
            
            }


            public Color fontColor { get; set; } = Color.Black;

            public bool isAddFoot { get; set; } = false;
            public Type type { get;private set; }
            public ExcelDataColumn(string columnName,Type type)
            {
                this.columnName = columnName; 
                this.type = type;
            }
            public static class NumberFormat
            {

                public const string Normal = @"@";
                public const string Int = @"#,##0";
                public const string Decimal = @"#,##0.00";

                public static string GetDecimal(int count)
                {
                    if (count <= 0)
                    {
                        return Int;
                    }

                    return @"#,##0." + new string('0', count);

                }

                public const string DateTime = @"yyyy-MM-dd HH:mm:ss";
                public const string ChineseDateTime = @"yyyy年MM月dd日HH时mm分ss秒";
                public const string Date = @"yyyy-MM-dd";
                public const string ChineseDate = @"yyyy年MM月dd日";
                public const string Time = @"HH:mm:ss";
                public const string ChineseTime = @"HH时mm分ss秒";

                
                public const string Percent = @"0.00%";

                public static string GetPercent(int count)
                {
                    if (count <= 0)
                    {
                        return @"0%";
                    }
                    return @"0." + new string('0', count) + '%';

                }


            }

            
        }
        public static string ToString(this ExcelHorizontalAlignment excelHorizontalAlignment)
        {
            switch (excelHorizontalAlignment)
            {
                case ExcelHorizontalAlignment.Left:
                    return @"left";
                case ExcelHorizontalAlignment.Right:
                    return @"right";
                default:
                    return @"center";
            }
        }
        public static Dictionary<string, ExcelDataColumn> ToDictionary(this ExcelDataColumn[]? excelDataColumns)
        {
            Dictionary<string, ExcelDataColumn> result = new Dictionary<string, ExcelDataColumn>();
            if (excelDataColumns is not null)
            {
                foreach (var item in excelDataColumns)
                {
                    result[item.columnName] = item;
                }
            }


            return result;
        }
        public static ExcelDataColumn[] GetExcelDataColumns(this Type type, ExcelDataColumn[]? outExcelDataColumns = null)
        {
            List<ExcelDataColumn> result = new List<ExcelDataColumn>();
            if (outExcelDataColumns is not null)
            {
                result.AddRange(outExcelDataColumns);
            }

            PropertyInfo? staticPropertyInfo = type.GetProperty(@"excelDataColumns", BindingFlags.Static | BindingFlags.Public);
            if (staticPropertyInfo is not null)
            {
                object? value = staticPropertyInfo.GetValue(null);
                if (value is ExcelDataColumn[] excelDataColumns)
                {
                    result.AddRange(excelDataColumns.Where(iter => !result.Any(it => it.columnName == iter.columnName)).ToArray());
                }
            }


            SugarColumnInfo[] sugarColumnInfos = type.GetSugarColumns();
            result.AddRange(sugarColumnInfos.Where(iter => !result.Any(it => it.columnName == iter.columnName)).Select(it => it.excelDataColumn).ToArray());

            return result.ToArray();
        }
        public static ExcelDataColumn[] GetExcelDataColumns(this DataTable dataTable, ExcelDataColumn[]? outExcelDataColumns = null)
        {
            List<ExcelDataColumn> result = new List<ExcelDataColumn>();
            if (outExcelDataColumns is not null)
            {
                result.AddRange(outExcelDataColumns);
            }
            DataColumn[] dataColumns = dataTable.Columns.Cast<DataColumn>().ToArray();
            result.AddRange(dataColumns.Where(iter => !result.Any(it => it.columnName == iter.ColumnName)).Select(it => new ExcelDataColumn(it.ColumnName,it.DataType)).ToArray());

            return result.ToArray();
        }
        #endregion


        #region ToHtml
        private static XmlElement AddRootElement(this XmlDocument xml, string? styleString)
        {
            XmlElement rootElement = xml.CreateElement(@"table");
            xml.AppendChild(rootElement);
            rootElement.SetAttribute(@"class", @"table table-bordered");
            if (!string.IsNullOrEmpty(styleString))
            {
                rootElement.SetAttribute("style", styleString);
            }
            return rootElement;
        }
        private static XmlElement AddTheadElement(this XmlDocument xml, XmlElement rootElement, string[] columnNames, bool isFrezzon)
        {
            XmlElement theadElement = xml.CreateElement(@"thead");
            rootElement.AppendChild(theadElement);
            theadElement.SetAttribute(@"class", @"table-dark text-center");
            if (isFrezzon)
            {
                theadElement.SetAttribute(@"style", @"position:sticky;top:0;");
            }
            {
                XmlElement trElement = xml.CreateElement(@"tr");
                theadElement.AppendChild(trElement);
                foreach (var columnName in columnNames)
                {
                    XmlElement thElement = xml.CreateElement(@"th");
                    trElement.AppendChild(thElement);

                    thElement.InnerText = columnName;
                    thElement.SetAttribute(@"scope", @"col");
                }
            }
            return theadElement;
        }
        private static XmlElement? AddCaptionElement(this XmlDocument xml, XmlElement rootElement, string? tableName)
        {
            if (string.IsNullOrEmpty(tableName))
            {
                return null;
            }
            else
            {
                XmlElement captionElement = xml.CreateElement(@"caption");
                rootElement.AppendChild(captionElement);

                captionElement.InnerText = tableName;
                captionElement.SetAttribute(@"class", @"caption-top text-center fs-4");
                return captionElement;
            }
        }
        private static string GetInnerText(this Type type, object? obj, string numberFormat)
        {
            string innerString = emptyString;
            if (obj is null || obj is DBNull)
            {
                if (type.IsOrNullableInteger() || type.IsOrNullableUInteger())
                {
                    innerString = (0L).ToString(numberFormat);
                }
                else if (type.IsOrNullableFloat() || type.IsOrNullableDecimal())
                {
                    innerString = (0D).ToString(numberFormat);
                }
            }
            else
            {
                if (type.IsOrNullableInteger())
                {
                    innerString = System.Convert.ToInt64(obj).ToString(numberFormat);
                }
                else if (type.IsOrNullableUInteger())
                {
                    innerString = System.Convert.ToUInt64(obj).ToString(numberFormat);
                }

                else if (type.IsOrNullableFloat())
                {
                    innerString = System.Convert.ToDouble(obj).ToString(numberFormat);
                }
                else if (type.IsOrNullableDecimal())
                {
                    innerString = System.Convert.ToDecimal(obj).ToString(numberFormat);
                }
                else if (type.IsOrNullableBool())
                {
                    innerString = Convert.ToBoolean(obj).ToString();
                }
                else if (type.IsOrNullableChar() || type.IsString())
                {
                    innerString = System.Convert.ToString(obj) ?? string.Empty;
                }
                else if (type.IsOrNullableDateTime())
                {
                    innerString = System.Convert.ToDateTime(obj).ToString(numberFormat);
                }
                else if (type.IsByteArray())
                {
                    innerString = @"Base64:" + System.Convert.ToBase64String((byte[])obj);
                }
            }
            return innerString;
        }

        public static XmlDocument ToHtml(this DataTable dataTable, ExcelDataColumn[]? excelDataColumns = null, string? styleString = null, bool isFrezzon = false)
        {
            string subName = @"将表格转化为XML文档";
            if (!string.IsNullOrEmpty(dataTable.TableName))
            {
                subName = $"将表格“{dataTable.TableName}”转化为XML文档";
            }
            Workflow workflow = new Workflow(subName);


            workflow.WorkingContent = @"数据准备";
            #region
            XmlDocument result = new XmlDocument();
            string firstColumnName = dataTable.Columns.Count > 0 ? dataTable.Columns[0].ColumnName : string.Empty;
            Dictionary<string, Type> columns = dataTable.Columns.Cast<DataColumn>().ToDictionary(it => it.ColumnName, it => it.DataType);
            Dictionary<string, ExcelDataColumn> excelDataColumnsDic = excelDataColumns.ToDictionary();
            foreach (var column in columns)
            {
                ExcelDataColumn excelDataColumn = excelDataColumnsDic.ContainsKey(column.Key) ?
                    excelDataColumnsDic[column.Key] : new ExcelDataColumn(column.Key,column.Value);
                excelDataColumnsDic[column.Key] = excelDataColumn;
            }
            HashSet<string> addFootColumnsHash = columns.Where(it =>
            {
                if (!it.Value.IsOrNullableNumber())
                {
                    return false;
                }
                return excelDataColumnsDic[it.Key].isAddFoot;
            })
                .Select(it => it.Key)
                .ToHashSet();
            #endregion

            workflow.WorkingContent = @"添加根节点";
            #region
            XmlElement rootElement = result.AddRootElement(styleString);
            #endregion

            workflow.WorkingContent = @"添加标题";
            #region
            result.AddCaptionElement(rootElement, dataTable.TableName);
            #endregion

            workflow.WorkingContent = @"添加字段";
            #region
            XmlElement theadElement = result.AddTheadElement(rootElement, columns.Keys.ToArray(), isFrezzon);
            #endregion

            workflow.WorkingContent = @"添加数据";
            #region
            XmlElement tbodyElement = result.CreateElement(@"tbody");
            rootElement.AppendChild(tbodyElement);
            foreach (DataRow row in dataTable.Rows)
            {
                XmlElement trElement = result.CreateElement(@"tr");
                tbodyElement.AppendChild(trElement);
                foreach (var column in columns)
                {
                    XmlElement tdElement = result.CreateElement(@"td");
                    trElement.AppendChild(tdElement);

                    Type type = column.Value;
                    string columnName = column.Key;
                    ExcelDataColumn excelDataColumn = excelDataColumnsDic[columnName];

                    //设定对齐方式
                    tdElement.SetAttribute(@"align", excelDataColumn.horizontalAlignment.ToString());

                    //录入数据
                    tdElement.InnerText = type.GetInnerText(row[columnName], excelDataColumn.numberFormat);
                }
            }
            #endregion

            workflow.WorkingContent = @"添加页脚";
            #region
            if (addFootColumnsHash.Count > 0)
            {
                bool isFirstColumn = true;
                XmlElement tfootElement = result.CreateElement(@"tfoot");
                rootElement.AppendChild(tfootElement);
                {
                    XmlElement trElement = result.CreateElement(@"tr");
                    tfootElement.AppendChild(trElement);
                    foreach (var column in columns)
                    {
                        XmlElement tdElement = result.CreateElement(@"td");
                        trElement.AppendChild(tdElement);
                        if (isFirstColumn)
                        {
                            tdElement.InnerText = @"合计：";
                            tdElement.SetAttribute(@"align", @"left");
                            tdElement.SetAttribute(@"class", @"fw-bold");
                            isFirstColumn = false;
                        }
                        else if (addFootColumnsHash.Contains(column.Key))
                        {
                            Type type = column.Value;
                            string columnName = column.Key;
                            ExcelDataColumn excelDataColumn = excelDataColumnsDic[columnName];

                            tdElement.SetAttribute(@"class", @"fw-bold");

                            //设定对齐方式
                            tdElement.SetAttribute(@"align", excelDataColumn.horizontalAlignment.ToString());


                            #region 添加数据
                            if (type.IsOrNullableInteger() || type.IsOrNullableUInteger())
                            {
                                var sum = dataTable.Rows.Cast<DataRow>().Sum(it =>
                                {
                                    var obj = it[columnName];
                                    return System.Convert.ToInt64(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            else if (type.IsOrNullableFloat())
                            {
                                var sum = dataTable.Rows.Cast<DataRow>().Sum(it =>
                                {
                                    var obj = it[columnName];
                                    return System.Convert.ToDouble(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            else if (type.IsOrNullableDecimal())
                            {
                                var sum = dataTable.Rows.Cast<DataRow>().Sum(it =>
                                {
                                    var obj = it[columnName];
                                    return System.Convert.ToDecimal(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            #endregion


                        }
                    }
                }



            }

            #endregion

            workflow.End();
            return result;
        }
        public static XmlDocument ToHtml<T>(this T[] values, string? tableName = null, ExcelDataColumn[]? excelDataColumns = null, string? styleString = null, bool isFrezzon = false) where T : class
        {
            Type valueType = typeof(T);
            string subName;
            if (tableName.IsNullOrEmpty())
            {
                subName = $"将表格“{valueType.Name}”转化为XML文档";
            }
            else
            {
                subName = $"将表格“{tableName}”转化为XML文档";
            }
            Workflow workflow = new Workflow($"将表格“{valueType.Name}”转化为XML文档");

            workflow.WorkingContent = @"数据准备";
            #region
            XmlDocument result = new XmlDocument();
            SugarColumnInfo[] sugarColumnInfos = valueType.GetSugarColumns();
            Dictionary<string, ExcelDataColumn> excelDataColumnsDic = valueType.GetExcelDataColumns(excelDataColumns).ToDictionary();
            HashSet<string> addFootColumnsHash = sugarColumnInfos
                .Where(it =>
                {
                    if (!it.type.IsOrNullableNumber())
                    {
                        return false;
                    }
                    if (it.columnName == (sugarColumnInfos.Length > 0 ? sugarColumnInfos.First().columnName : string.Empty))
                    {
                        return false;
                    }
                    return excelDataColumnsDic[it.columnName].isAddFoot;
                })
                .Select(it => it.columnName)
                .ToHashSet();
            #endregion

            workflow.WorkingContent = @"添加根节点";
            #region
            XmlElement rootElement = result.AddRootElement(styleString);
            #endregion

            workflow.WorkingContent = @"添加标题";
            #region
            result.AddCaptionElement(rootElement, tableName);
            #endregion

            workflow.WorkingContent = @"添加字段";
            #region
            XmlElement theadElement = result.AddTheadElement(rootElement, sugarColumnInfos.Select(it => it.columnName).ToArray(), isFrezzon);
            #endregion

            workflow.WorkingContent = @"添加数据";
            #region
            XmlElement tbodyElement = result.CreateElement(@"tbody");
            rootElement.AppendChild(tbodyElement);
            foreach (T row in values)
            {
                XmlElement trElement = result.CreateElement(@"tr");
                tbodyElement.AppendChild(trElement);
                foreach (var column in sugarColumnInfos)
                {
                    XmlElement tdElement = result.CreateElement(@"td");
                    trElement.AppendChild(tdElement);

                    Type type = column.type;
                    string columnName = column.columnName;
                    ExcelDataColumn excelDataColumn = excelDataColumnsDic[columnName];

                    //设定对齐方式
                    tdElement.SetAttribute(@"align", excelDataColumn.horizontalAlignment.ToString());

                    //录入数据
                    PropertyInfo propertyInfo = valueType.GetProperty(column.propertyName)!;
                    tdElement.InnerText = type.GetInnerText(propertyInfo.GetValue(row), excelDataColumn.numberFormat);
                }
            }
            #endregion

            workflow.WorkingContent = @"添加页脚";
            #region
            if (addFootColumnsHash.Count > 0)
            {
                bool isFirstColumn = true;
                XmlElement tfootElement = result.CreateElement(@"tfoot");
                rootElement.AppendChild(tfootElement);
                {
                    XmlElement trElement = result.CreateElement(@"tr");
                    tfootElement.AppendChild(trElement);
                    foreach (var column in sugarColumnInfos)
                    {
                        XmlElement tdElement = result.CreateElement(@"td");
                        trElement.AppendChild(tdElement);
                        if (isFirstColumn)
                        {
                            tdElement.InnerText = @"合计：";
                            tdElement.SetAttribute(@"align", @"left");
                            tdElement.SetAttribute(@"class", @"fw-bold");
                            isFirstColumn = false;
                        }
                        else if (addFootColumnsHash.Contains(column.columnName))
                        {
                            Type type = column.type;
                            string columnName = column.columnName;
                            ExcelDataColumn excelDataColumn = excelDataColumnsDic[columnName];

                            tdElement.SetAttribute(@"class", @"fw-bold");

                            //设定对齐方式
                            tdElement.SetAttribute(@"align", excelDataColumn.horizontalAlignment.ToString());

                            #region 添加数据
                            PropertyInfo propertyInfo = valueType.GetProperty(column.propertyName)!;
                            if (type.IsOrNullableInteger() || type.IsOrNullableUInteger())
                            {
                                var sum = values.Sum(it =>
                                {
                                    object? obj = propertyInfo.GetValue(it);
                                    return System.Convert.ToInt64(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            else if (type.IsOrNullableFloat())
                            {
                                var sum = values.Sum(it =>
                                {
                                    object? obj = propertyInfo.GetValue(it);
                                    return System.Convert.ToDouble(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            else if (type.IsOrNullableDecimal())
                            {
                                var sum = values.Sum(it =>
                                {

                                    object? obj = propertyInfo.GetValue(it);
                                    return System.Convert.ToDecimal(obj);
                                });
                                tdElement.InnerText = sum.ToString(excelDataColumn.numberFormat);
                            }
                            #endregion
                        }
                    }
                }



            }

            #endregion

            workflow.End();
            return result;
        }
        #endregion


        #region ToExcelWorksheet
        public static ExcelWorksheet ToExcelWorksheet(this DataTable dataTable, ExcelPackage package, ExcelDataColumn[]? excelDataColumns = null)
        {

            #region 数据准备
            string tableName = string.IsNullOrEmpty(dataTable.TableName) ? @"数据源" : dataTable.TableName;
            Dictionary<string, Type> columns = dataTable.Columns.Cast<DataColumn>().ToDictionary(it => it.ColumnName, it => it.DataType);
            Dictionary<string, ExcelDataColumn> excelDataColumnsDic = excelDataColumns.ToDictionary();
            foreach (var column in columns)
            {
                ExcelDataColumn excelDataColumn = excelDataColumnsDic.ContainsKey(column.Key) ?
                    excelDataColumnsDic[column.Key] : new ExcelDataColumn(column.Key, column.Value);
                excelDataColumnsDic[column.Key] = excelDataColumn;
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(tableName);
            int titleRow = 1, startRow = 2, serCol = 1, rowsCount = dataTable.Rows.Count, columnsCount = dataTable.Columns.Count;
            ExcelRange rng, cell;
            #endregion


            #region 录入字段
            rng = worksheet.Cells[titleRow, serCol, titleRow, serCol + columnsCount - 1];
            rng.Style.Font.Bold = true;
            rng.Style.Font.Name = @"黑体";
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
            for (int j = 0; j < columnsCount; j++)
            {
                DataColumn column = dataTable.Columns[j];
                cell = worksheet.Cells[titleRow, serCol + j];
                rng = worksheet.Cells[titleRow + 1, serCol + j, titleRow + System.Math.Max(rowsCount, 1), serCol + j];
                ExcelDataColumn excelDataColumn = excelDataColumnsDic[column.ColumnName];
                System.Type type = column.DataType;

                cell.Value = excelDataColumn.columnName;
                rng.Style.Font.Name = @"宋体";
                rng.Style.Font.Bold = false;
                rng.Style.Numberformat.Format = excelDataColumn.numberFormat;
                rng.Style.HorizontalAlignment = excelDataColumn.horizontalAlignment;
                rng.Style.VerticalAlignment = excelDataColumn.verticalAlignment;
                worksheet.Columns[serCol + j].Width = excelDataColumn.width.Around(6, 30);


            }
            #endregion




            #region 添加数据
            for (int i = 0; i < rowsCount && i < rowsCount; i++)
            {
                DataRow row = dataTable.Rows[i];
                for (int j = 0; j < columnsCount; j++)
                {
                    cell = worksheet.Cells[startRow + i, serCol + j];
                    if (row[j] is null || row[j] is DBNull)
                    {
                        System.Type type = dataTable.Columns[j].DataType;
                        if (type.IsOrNullableNumber())
                        {
                            cell.Value = 0M;
                        }
                    }
                    else
                    {
                        cell.Value = row[j];
                    }
                }
            }
            #endregion


            #region 添加边框
            rng = worksheet.Cells[titleRow, serCol, titleRow + System.Math.Max(rowsCount, 1), serCol + columnsCount - 1];
            ExcelTable table = worksheet.Tables.Add(rng, tableName);// 将数据范围转换为Excel表格
            rng.Style.Font.Size = 11;//设置字体大小
            rng.Style.WrapText = true;//自动换行
            #endregion

            return worksheet;

        }


        public static ExcelWorksheet ToExcelWorksheet<T>(this T[] values, ExcelPackage package, string tableName = emptyString, ExcelDataColumn[]? excelDataColumns = null)
        {
            #region 数据准备
            Type valueType = typeof(T);
            tableName = string.IsNullOrEmpty(tableName) ? valueType.GetSugarTableName() : tableName;
            SugarColumnInfo[] sugarColumnInfos = valueType.GetSugarColumns();
            Dictionary<string, ExcelDataColumn> excelDataColumnsDic = excelDataColumns.ToDictionary();
            foreach (var column in sugarColumnInfos)
            {
                ExcelDataColumn excelDataColumn = excelDataColumnsDic.ContainsKey(column.columnName) ?
                    excelDataColumnsDic[column.columnName] : column.excelDataColumn;
                excelDataColumnsDic[column.columnName] = excelDataColumn;
            }
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add(tableName);
            ExcelRange rng, cell;
            int titleRow = 1, startRow = 2, serCol = 1, rowsCount = values.Length, columnsCount = sugarColumnInfos.Length;
            #endregion

            #region 录入字段
            rng = worksheet.Cells[titleRow, serCol, titleRow, serCol + columnsCount - 1];
            rng.Style.Font.Bold = true;
            rng.Style.Font.Name = @"黑体";
            rng.Style.HorizontalAlignment = ExcelHorizontalAlignment.Center;
            rng.Style.VerticalAlignment = ExcelVerticalAlignment.Center;
           

            
            for (int j = 0; j < columnsCount; j++)
            {
                SugarColumnInfo column = sugarColumnInfos[j];
                cell = worksheet.Cells[titleRow, serCol + j];
                rng = worksheet.Cells[titleRow + 1, serCol + j, titleRow + System.Math.Max(rowsCount, 1), serCol + j];
                ExcelDataColumn excelDataColumn = excelDataColumnsDic[column.columnName];
                System.Type type = column.type;
                PropertyInfo propertyInfo = valueType.GetProperty(column.propertyName)!;

                cell.Value = excelDataColumn.columnName;
                rng.Style.Font.Bold = false;
                rng.Style.Font.Name = @"宋体";
                rng.Style.Numberformat.Format = excelDataColumn.numberFormat;
                rng.Style.HorizontalAlignment = excelDataColumn.horizontalAlignment;
                rng.Style.VerticalAlignment = excelDataColumn.verticalAlignment;
                worksheet.Columns[serCol + j].Width = excelDataColumn.width.Around(6, 30);
            }
            #endregion


            #region 添加数据
            for (int i = 0; i < rowsCount; i++)
            {
                T row = values[i];
                for (int j = 0; j < columnsCount; j++)
                {
                    SugarColumnInfo column = sugarColumnInfos[j];
                    PropertyInfo propertyInfo = valueType.GetProperty(column.propertyName)!;
                    object? obj = propertyInfo.GetValue(row);
                    cell = worksheet.Cells[startRow + i, serCol + j];
                    if (obj is null)
                    {
                        if (column.type.IsOrNullableNumber())
                        {
                            cell.Value = 0M;
                        }
                    }
                    else
                    {
                        cell.Value = obj;
                    }
                }
            }
            #endregion
            #region 添加边框
            rng = worksheet.Cells[titleRow, serCol, titleRow + System.Math.Max(rowsCount, 1), serCol + columnsCount - 1];
            _ = worksheet.Tables.Add(rng, tableName);// 将数据范围转换为Excel表格
            rng.Style.Font.Size = 11;//设置字体大小
            rng.Style.WrapText = true;//自动换行
            #endregion
            

            return worksheet;
        }

        #endregion







        


        /// <summary>
        /// 将DataTable转化为数组
        /// </summary>
        /// <typeparam name="T">数组类型</typeparam>
        /// <param name="dataTable">要转化的DataTable</param>
        public static async Task<T[]> ToArray<T>(this DataTable dataTable) where T : class, new()
        {

            Type objType = typeof(T);
            SugarColumnInfo[] sugarColumnInfos = objType.GetSugarColumns();
            HashSet<string> dataTableColumns = dataTable.Columns.Cast<DataColumn>().Select(iterator => iterator.ColumnName).ToHashSet();
            List<Task> tasks = new List<Task>();

            List<T> result = new List<T>();
            foreach (DataRow row in dataTable.Rows)
            {
                DataRow _row = row;
                tasks.Add(Task.Run(() =>
                {
                    T obj = new T();
                    foreach (var column in sugarColumnInfos)
                    {
                        if (dataTableColumns.Contains(column.columnName))
                        {
                            DataColumn dataColumn = dataTable.Columns[column.columnName]!;
                            PropertyInfo propertyInfo = objType.GetProperty(column.propertyName)!;
                            object? valueSource = _row[column.columnName];
                            object? value = column.type.GetValue(valueSource);
                            propertyInfo.SetValue(obj, value);
                        }
                    }
                    result.Add(obj);
                }));
            }

            await Task.WhenAll(tasks);

            return result.ToArray();
        }

    }
}
