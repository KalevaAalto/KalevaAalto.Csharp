using KalevaAalto.Models.Excel;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using OfficeOpenXml;
using System.Reflection;
using KalevaAalto.Models.Excel.Enums;
using KalevaAalto;

namespace KalevaAalto.Models.Excel
{
    public abstract class IWorkbook : IDisposable
    {

        public string FileName { get; set; }
        public IWorkbook(string fileName)
        {
            FileName = fileName;
            Init();
        }
        protected abstract void Init();
        public bool FileExist { get => File.Exists(FileName); }
        public abstract IWorksheet[] Worksheets { get; }
        public IWorksheet this[string Name] { get => Worksheets.First(it => it.Name == Name); }


        public abstract IWorksheet AddWorksheet(string name);
        public void Save()
        {
            if (Worksheets.Length == 0)
            {
                AddWorksheet(@"Sheet1");
            }
            _Save();
        }
        public string Description { get; set; } = string.Empty;
        protected abstract void _Save();

        public IWorksheet AddWorksheet(System.Data.DataTable dataTable, DataColumnStyle[]? dataColumnStyles = null)
        {
            if (string.IsNullOrEmpty(dataTable.TableName))
            {
                throw new Exception(@"表格没有命名！");
            }
            int rowCount = dataTable.Rows.Count, columnCount = dataTable.Columns.Count;
            if (columnCount <= 0)
            {
                throw new Exception(@"表格中没有字段！");
            }
            IWorksheet worksheet = this.AddWorksheet(dataTable.TableName);
            int titleRow = 1, startRow = titleRow + 1, serCol = 1;
            ICell cell;
            IRange rng;


            //设置字段格式
            rng = worksheet[titleRow, serCol, titleRow, serCol + columnCount - 1];
            rng.Style.FontWeight = FontWeight.Thick;
            rng.Style.FontFamily = @"黑体";
            rng.Style.HorizontalAlignment = HorizontalAlignment.Center;
            rng.Style.VerticalAlignment = VerticalAlignment.Center;


            Dictionary<string, DataColumnStyle> dataColumnStyleDic = dataTable.GetDataColumnStyles(dataColumnStyles).ToDictionary();
            //录入字段名称
            for (int j = 0; j < columnCount; j++)
            {
                cell = worksheet[titleRow, serCol + j];
                rng = worksheet[titleRow + 1, serCol + j, titleRow + rowCount, serCol + j];
                DataColumn dataColumn = dataTable.Columns[j];
                DataColumnStyle dataColumnStyle = dataColumnStyleDic[dataColumn.ColumnName];

                cell.Value = dataColumn.ColumnName;
                cell.Column.Width = dataColumnStyle.Width;
                rng.Style.FontFamily = @"宋体";
                rng.Style.FontWeight = FontWeight.Thin;
                rng.Style.NumberFormatString = dataColumnStyle.NumberFormat;
                rng.Style.VerticalAlignment = dataColumnStyle.VerticalAlignment;
                rng.Style.HorizontalAlignment = dataColumnStyle.HorizontalAlignment;

            }

            //录入数据
            for (int i = 0; i < rowCount; i++)
            {
                DataRow dataRow = dataTable.Rows[i];
                for (int j = 0; j < columnCount; j++)
                {
                    object? value = dataRow[j];
                    if (value != null && value != DBNull.Value)
                    {
                        cell = worksheet[startRow + i, serCol + j];
                        cell.Value = value;
                    }
                }
            }



            //全局
            rng = worksheet[titleRow, serCol, titleRow + rowCount, serCol + columnCount - 1];
            rng.Style.WarpText = true;
            rng.SetTable(dataTable.TableName);




            return worksheet;
        }

        public IWorksheet AddWorksheet<T>(T[] objs, string tableName, DataColumnStyle[]? dataColumnStyles = null)
        {
            if (string.IsNullOrEmpty(tableName))
            {
                throw new Exception(@"表格没有命名！");
            }
            ClassColumnInfo[] classColumnInfos = typeof(T).GetClassColumnInfos();
            int rowCount = objs.Length, columnCount = classColumnInfos.Length;
            if (columnCount <= 0)
            {
                throw new Exception(@"表格中没有字段！");
            }
            IWorksheet worksheet = AddWorksheet(tableName);
            int titleRow = 1, startRow = titleRow + 1, serCol = 1;
            ICell cell;
            IRange rng;


            //设置字段格式
            rng = worksheet[titleRow, serCol, titleRow, serCol + columnCount - 1];
            rng.Style.FontWeight = FontWeight.Thick;
            rng.Style.FontFamily = @"黑体";
            rng.Style.HorizontalAlignment = HorizontalAlignment.Center;
            rng.Style.VerticalAlignment = VerticalAlignment.Center;


            Dictionary<string, DataColumnStyle> dataColumnStyleDic = typeof(T).GetDataColumnStyles(dataColumnStyles).ToDictionary();
            //录入字段名称
            for (int j = 0; j < columnCount; j++)
            {
                cell = worksheet[titleRow, serCol + j];
                rng = worksheet[titleRow + 1, serCol + j, titleRow + rowCount, serCol + j];
                ClassColumnInfo classColumnInfo = classColumnInfos[j];
                DataColumnStyle dataColumnStyle = dataColumnStyleDic[classColumnInfo.ColumnName];

                cell.Value = classColumnInfo.ColumnName;
                cell.Column.Width = dataColumnStyle.Width;
                rng.Style.FontFamily = @"宋体";
                rng.Style.FontWeight = FontWeight.Thin;
                rng.Style.NumberFormatString = dataColumnStyle.NumberFormat;
                rng.Style.VerticalAlignment = dataColumnStyle.VerticalAlignment;
                rng.Style.HorizontalAlignment = dataColumnStyle.HorizontalAlignment;

            }

            //录入数据
            for (int i = 0; i < rowCount; i++)
            {
                T obj = objs[i];
                for (int j = 0; j < columnCount; j++)
                {
                    ClassColumnInfo classColumnInfo = classColumnInfos[j];
                    PropertyInfo propertyInfo = typeof(T).GetProperty(classColumnInfo.PropertyName)!;
                    object? value = propertyInfo.GetValue(obj);
                    if (value != null && value != DBNull.Value)
                    {
                        cell = worksheet[startRow + i, serCol + j];
                        cell.Value = value;
                    }
                }
            }



            //全局
            rng = worksheet[titleRow, serCol, titleRow + rowCount, serCol + columnCount - 1];
            rng.Style.WarpText = true;
            rng.SetTable(tableName);




            return worksheet;
        }

        public abstract void Dispose();

    }
}


namespace KalevaAalto
{
    public static partial class Static
    {
        public static IWorksheet ToExcelWorksheet(this DataTable dataTable, IWorkbook workbook, DataColumnStyle[]? excelDataColumns = null)
        {
            return workbook.AddWorksheet(dataTable, excelDataColumns);
        }


        public static IWorksheet ToExcelWorksheet<T>(this T[] values, IWorkbook workbook, string tableName = EmptyString, DataColumnStyle[]? excelDataColumns = null)
        {
            return workbook.AddWorksheet(values, tableName, excelDataColumns);
        }

    }
}
