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
using SqlSugar;
using KalevaAalto.Models.DataTable;





namespace KalevaAalto.Models.Excel
{
    public class DataColumnStyle
    {
        private string _columnName;
        private Type _type;
        private HorizontalAlignment? _horizontalAlignment = null;
        private VerticalAlignment? _verticalAlignment = null;
        private double? _width = null;
        private string? _numberFormat = null;
        private Color _fontColor = Color.Black;
        private bool _isAddFoot = false;
        public DataColumnStyle(string columnName, Type type)
        {
            _columnName = columnName;
            _type = type;
        }

        public string ColumnName => _columnName;
        public Type Type =>_type;
        public Color FontColor { get=>_fontColor;init=> _fontColor=value; }
        public bool IsAddFoot { get=> _isAddFoot; init=> _isAddFoot=value; }
        public HorizontalAlignment HorizontalAlignment
        {
            get
            {
                if (_horizontalAlignment is null)
                {
                    if (_type.IsNumber()) return HorizontalAlignment.Right;
                    else if (_type.IsByteArray()) return HorizontalAlignment.Left;
                    else return HorizontalAlignment.Center;
                }
                else return this._horizontalAlignment.Value;
            }
            init => _horizontalAlignment = value;

        }
        public VerticalAlignment VerticalAlignment
        {
            get => _verticalAlignment is null ? VerticalAlignment.Center : _verticalAlignment.Value;
            init => _verticalAlignment = value;
        }
        public double Width
        {
            get
            {
                if (_width is null)
                {
                    if (_type.IsOrNullableChar() || _type == typeof(byte) || _type == typeof(byte?) || _type == typeof(sbyte) || _type == typeof(sbyte?)) return 6;
                    else if (_type.IsOrNullableBool()) return 8;
                    else if (_type.IsOrNullableNumber() || _type.IsOrNullableDateTime()) return 15;
                    else if (_type.IsByteArray()) return 30;
                    else return 12;
                }
                else return _width.Value.Around(6, 100);
            }
            init => _width = value;
        }

        
        public string NumberFormat
        {
            get
            {
                if (_numberFormat is null)
                {
                    if (this.Type.IsOrNullableInteger() || this.Type.IsOrNullableUInteger()) return NumberFormatTemplate.Int;
                    else if (this.Type.IsOrNullableFloat() || this.Type.IsOrNullableDecimal()) return NumberFormatTemplate.Decimal;
                    else if (this.Type.IsOrNullableDateTime()) return NumberFormatTemplate.Date;
                    else return string.Empty;
                }
                else return _numberFormat;
            }
            init => _numberFormat = value;
        }
        
        
    }
}



namespace KalevaAalto
{
    public static partial class Static
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