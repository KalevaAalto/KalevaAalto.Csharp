using KalevaAalto.Models.Excel;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel.Enums;





namespace KalevaAalto.Models.Excel
{
    public class DataColumnStyle
    {
        


        public string ColumnName { get; private set; }

        private HorizontalAlignment? _horizontalAlignment = null;
        public HorizontalAlignment HorizontalAlignment
        {
            get
            {
                if (this._horizontalAlignment is null)
                {
                    if (this.Type.IsNumber())
                    {
                        return HorizontalAlignment.Right;
                    }
                    else if (this.Type.IsByteArray())
                    {
                        return HorizontalAlignment.Left;
                    }
                    else
                    {
                        return HorizontalAlignment.Center;
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
        private VerticalAlignment? _verticalAlignment = null;
        public VerticalAlignment VerticalAlignment
        {
            get
            {
                if (this._verticalAlignment is null)
                {
                    return VerticalAlignment.Center;
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
        private double? _width = null;
        public double Width
        {
            get
            {
                if (this._width is null)
                {
                    if (this.Type.IsOrNullableChar() || this.Type == typeof(byte) || this.Type == typeof(byte?))
                    {
                        return 6;
                    }
                    else if (this.Type.IsOrNullableBool())
                    {
                        return 8;
                    }
                    else if (this.Type.IsOrNullableNumber() || this.Type.IsOrNullableDateTime())
                    {
                        return 15;
                    }
                    else if (this.Type.IsByteArray())
                    {
                        return 30;
                    }
                    else
                    {
                        return 12;
                    }
                }
                else
                {
                    return this._width.Value.Around(6, 100);
                }
            }
            set
            {
                this._width = value;
            }
        }

        private string? _numberFormat = null;
        public string NumberFormat
        {
            get
            {
                if (this._numberFormat is null)
                {
                    if (this.Type.IsOrNullableInteger() || this.Type.IsOrNullableUInteger())
                    {
                        return NumberFormatTemplate.Int;
                    }
                    else if (this.Type.IsOrNullableFloat() || this.Type.IsOrNullableDecimal())
                    {
                        return NumberFormatTemplate.Decimal;
                    }
                    else if (this.Type.IsOrNullableDateTime())
                    {
                        return NumberFormatTemplate.Date;
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
        public Color FontColor { get; set; } = Color.Black;

        public bool IsAddFoot { get; set; } = false;
        public Type Type { get; private set; }
        public DataColumnStyle(string columnName, Type type)
        {
            this.ColumnName = columnName;
            this.Type = type;
        }
    }
}



namespace KalevaAalto
{
    public static partial class Main
    {
        public static Dictionary<string,DataColumnStyle> ToDictionary(this DataColumnStyle[]? dataColumnStyles)
        {
            if (dataColumnStyles == null) return new Dictionary<string, DataColumnStyle>();
            return dataColumnStyles.Select(it => new { it.ColumnName, Value = it }).ToDictionary(it => it.ColumnName, it => it.Value);
        }


        public static Models.Excel.DataColumnStyle[] GetDataColumnStyles(this DataTable dataTable, Models.Excel.DataColumnStyle[]? outExcelDataColumns = null)
        {
            List<Models.Excel.DataColumnStyle> result = new List<Models.Excel.DataColumnStyle>();
            if (outExcelDataColumns is not null)
            {
                result.AddRange(outExcelDataColumns);
            }
            DataColumn[] dataColumns = dataTable.Columns.Cast<DataColumn>().ToArray();
            result.AddRange(dataColumns.Where(iter => !result.Any(it => it.ColumnName == iter.ColumnName)).Select(it => new Models.Excel.DataColumnStyle(it.ColumnName, it.DataType)).ToArray());

            return result.ToArray();
        }


        public static DataColumnStyle[] GetDataColumnStyles(this Type type,DataColumnStyle[]? outExcelDataColumns = null)
        {
            List<DataColumnStyle> result = new List<DataColumnStyle>();
            if (outExcelDataColumns is not null)
            {
                result.AddRange(outExcelDataColumns);
            }
            PropertyInfo? staticPropertyInfo = type.GetProperty(@"dataColumnStyles", BindingFlags.Static | BindingFlags.NonPublic);
            if (staticPropertyInfo is not null)
            {
                object? value = staticPropertyInfo.GetValue(null);
                if (value is DataColumnStyle[] dataColumnStyles)
                {
                    result.AddRange(dataColumnStyles.Where(iter => !result.Any(it => it.ColumnName == iter.ColumnName)).ToArray());
                }
            }


            ClassColumnInfo[] sugarColumnInfos = type.GetClassColumnInfos();
            result.AddRange(sugarColumnInfos.Where(iter => !result.Any(it => it.ColumnName == iter.ColumnName)).Select(it => new DataColumnStyle(it.ColumnName, it.Type)).ToArray());

            return result.ToArray();
        }
    }
}