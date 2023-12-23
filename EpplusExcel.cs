using Google.Protobuf.WellKnownTypes;
using MySqlX.XDevAPI.Relational;
using OfficeOpenXml;
using OfficeOpenXml.Interfaces.Drawing.Text;
using OfficeOpenXml.Style;
using OfficeOpenXml.Table;
using Org.BouncyCastle.Asn1.Mozilla;
using Org.BouncyCastle.Asn1.X509.Qualified;
using SqlSugar.Extensions;
using System.Collections.Immutable;
using System.Configuration;
using System.Data;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using static KalevaAalto.Main;
using System.Xml.Linq;
using System.Xml;


namespace KalevaAalto
{
    public static partial class Main
    {
        /// <summary>
        /// 从文件中获取Excel文档
        /// </summary>
        public static ExcelPackage GetExcelPackage(this FileNameInfo fileNameInfo)
        {
            return new ExcelPackage(fileNameInfo.fileInfo);
        }


        /// <summary>
        /// 将Excel文档另存为至文件中
        /// </summary>
        /// <param name="package">要保存的Excel文档</param>
        public static void SaveAs(this ExcelPackage package,FileNameInfo fileNameInfo)
        {
            package.SaveAs(fileNameInfo.fileInfo);
        }








        /// <summary>
        /// 单元格坐标
        /// </summary>
        public class CellPos
        {
            public int row { set; get; } = 0;
            public int column { set; get; } = 0;
            public CellPos(int row, int column)
            {
                this.row = row;
                this.column = column;
            }

            public bool Status()
            {
                return this.row > 0 && this.column > 0;
            }

            public readonly static CellPos errorCellPos = new CellPos(0, 0);

        }


        /// <summary>
        /// 检查单元格是否为空
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="cellPos">单元格</param>
        /// <returns>返回单元格是否为空</returns>
        public static bool IsCellEmpty(this ExcelWorksheet worksheet,CellPos cellPos)
        {
            if (!cellPos.Status())
            {
                return true;
            }

            object? value = worksheet.GetValue(cellPos.row, cellPos.column);
            if(value is null || string.IsNullOrEmpty(value.ToString()))
            {
                return true;
            }

            return false;

        }


        /// <summary>
        /// 检查单元格是否为空
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="row">行号</param>
        /// <param name="column">列号</param>
        /// <returns>返回单元格是否为空</returns>
        public static bool IsCellEmpty(this ExcelWorksheet worksheet, int row,int column)
        {
            return worksheet.IsCellEmpty(new CellPos(row,column));
        }



        public static byte[] ToBytes(this ExcelPackage excelPackage)
        {
            using (MemoryStream memoryStream = new MemoryStream())
            {
                excelPackage.SaveAs(memoryStream);
                return memoryStream.ToArray();
            }
        }


        /// <summary>
        /// 获取工作表的起始单元格坐标
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <returns>返回工作表的起始单元格坐标</returns>
        public static CellPos GetStartCellPos(this ExcelWorksheet worksheet)
        {
            for(int row = 0; row<500; row++)
            {
                for (int column = 0; column < 500; column++)
                {
                    if (!worksheet.IsCellEmpty(row,column))
                    {
                        return new CellPos(row, column);
                    }
                }
            }


            return CellPos.errorCellPos;
        }

        /// <summary>
        /// 获取工作表的起始单元格坐标
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="startString">起始点标志字符串</param>
        /// <returns>返回工作表的起始单元格坐标</returns>
        public static CellPos GetStartCellPos(this ExcelWorksheet worksheet,string startString)
        {
            if (string.IsNullOrEmpty(startString))
            {
                return worksheet.GetStartCellPos();
            }

            for (int row = 0; row < 500; row++)
            {
                for (int column = 0; column < 500; column++)
                {
                    if (!worksheet.IsCellEmpty(row, column) && worksheet.GetValue<string>(row,column) == startString)
                    {
                        return new CellPos(row, column);
                    }
                }
            }


            return CellPos.errorCellPos;
        }



        public static DataTable GetTableFromSheet(string file_name, string sheet_name)
        {

            var package = new ExcelPackage(new FileInfo(file_name));

            var worksheet = package.Workbook.Worksheets[sheet_name];
            return worksheet.ToDataTable();
        }




        public static HashSet<string> WorksheetNames(this ExcelPackage package)
        {
            HashSet<string> worksheetNames = new HashSet<string>();

            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                worksheetNames.Add(worksheet.Name);
            }

            return worksheetNames;
        }


        public static HashSet<string> WorksheetNames(string file_name)
        {
            ExcelPackage package = new ExcelPackage(new FileInfo(file_name));
            return package.WorksheetNames();
        }


        public static bool IsExists(this ExcelPackage package, string worksheet_name)
        {
            bool rs = false;
            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets)
            {
                if (worksheet.Name == worksheet_name) return true;
            }

            return rs;
        }


        public static bool IsExists(this ExcelPackage package, string[] worksheet_names)
        {
            bool rs = true;
            HashSet<string> strings = new HashSet<string>();
            foreach (ExcelWorksheet worksheet in package.Workbook.Worksheets) strings.Add(worksheet.Name);

            foreach (string worksheet_name in worksheet_names)
            {
                if (!strings.Contains(worksheet_name))
                {
                    return false;
                }
            }

            return rs;
        }




        private readonly static Regex regexEndNumber = new Regex(@"^(.*?)(\d*)$"); 

        /// <summary>
        /// 将单元格区域的数据转化为数据表
        /// </summary>
        /// <param name="rng">单元格区域</param>
        /// <returns>返回单元格区域的数据的数据表</returns>
        public static DataTable ExcelRangeToDataTable(this ExcelRange rng)
        {
            DataTable table = new DataTable();
            int columnsCount = rng.End.Column - rng.Start.Column + 1;
            int rowsCount = rng.End.Row - rng.Start.Row + 1;

            

            //录入字段列表
            for (int columnNumber = 1; columnNumber <= columnsCount; columnNumber++)
            {
                string columnName = rng[1, columnNumber].Value.ToString()!;
                while (table.IsExistsColumn(columnName))
                {
                    Match match = regexEndNumber.Match(columnName);
                    int matchNumber = KalevaAalto.Main.StringToInt(match.Groups[2].Value);
                    columnName = match.Groups[1].Value + (matchNumber + 1);
                }

                table.Columns.Add(columnName, typeof(object));
            }

            

            //录入数据
            for (int rowNumber = 2; rowNumber <= rowsCount; rowNumber++)
            {
                DataRow row = table.Rows.Add();
                for (int columnNumber = 1; columnNumber <= columnsCount; columnNumber++)
                {
                    row[columnNumber - 1] = rng[rowNumber, columnNumber].GetCellValue<object>();
                }

            }


            return table;
        }


        /// <summary>
        /// 读取worksheet，将其录入至DataTable中
        /// </summary>
        /// <param name="worksheet">工作表</param>
        /// <param name="startString"></param>
        /// <returns></returns>
        public static DataTable ToDataTable(this ExcelWorksheet worksheet,string startString = Main.emptyString)
        {
            CellPos startCellPos = worksheet.GetStartCellPos(startString);

            //定位
            int columnsCount = int.MaxValue;
            for (int columnNumber = startCellPos.column; columnNumber <= int.MaxValue; columnNumber++)
            {
                if (worksheet.IsCellEmpty(startCellPos.row, columnNumber))
                {
                    columnsCount = columnNumber - 1;
                    break;
                }
            }

            int rowsCount = int.MaxValue;
            for (int rowsNumber = startCellPos.row; rowsNumber <= int.MaxValue; rowsNumber++)
            {
                if (worksheet.IsCellEmpty(rowsNumber, startCellPos.column))
                {
                    rowsCount = rowsNumber - 1;
                    break;
                }
            }
            return worksheet.Cells[startCellPos.row, startCellPos.column, rowsCount, columnsCount].ExcelRangeToDataTable();
        }








        public static void FieldsToDate(this DataTable table, string[] field_names)
        {
            foreach (string field_name in field_names)
            {
                if (!table.IsExistsColumn(field_name))
                {
                    continue;
                }
                foreach (DataRow row in table.Rows)
                {
                    var cell = row[field_name];
                    if (cell is null)
                    {
                        continue;
                    }
                    if (cell.GetType() == typeof(int) || cell.GetType() == typeof(double))
                    {
                        row[field_name] = DateTime.FromOADate((double)cell);
                    }
                    else if (cell.GetType() == typeof(string))
                    {
                        string date_str = (string)cell;
                        row[field_name] = DateTime.FromOADate(Convert.ToDouble(date_str));
                    }
                    else row[field_name] = null;


                }
            }
        }


        public static MemoryStream SaveAsMemoryStream(this ExcelPackage package)
        {
            MemoryStream memoryStream = new MemoryStream();
            package.SaveAs(memoryStream);
            memoryStream.Position = 0;
            return memoryStream;
        }
        public static MemoryStream SaveToExcel(this DataTable table)
        {
            // 创建一个新的Excel文件
            ExcelPackage excelPackage = new ExcelPackage();

            // 新建一个名为"MySheet"的工作表
            string work_sheet_name = table.TableName;
            if (work_sheet_name.Length == 0) work_sheet_name = "Sheet1";
            ExcelWorksheet worksheet = excelPackage.Workbook.Worksheets.Add(work_sheet_name);

            //写入表头
            for (int j = 0; j < table.Columns.Count; j++) worksheet.Cells[1, j + 1].Value = table.Columns[j].ColumnName;

            for (int i = 0; i < table.Rows.Count; i++)
            {
                for (int j = 0; j < table.Columns.Count; j++) worksheet.Cells[i + 2, j + 1].Value = table.Rows[i][j];

            }


            return excelPackage.SaveAsMemoryStream();
        }

        


        public static void SaveToExcel(this DataTable table,string FileName)
        {
            var Stream = table.SaveToExcel();
            using (FileStream fileStream = new FileStream(FileName, FileMode.Create))
            {
                Stream.WriteTo(fileStream);
            }

        }


        public static bool IsEmpty(this ExcelRange rng)
        {
            return rng.Value == null || rng.GetCellValue<string>().Length == 0;
        }



        public static void SetAllBorder(this ExcelRange rng,ExcelBorderStyle excelBorderStyle = ExcelBorderStyle.Thin)
        {
            var border = rng.Style.Border;
            border.Top.Style = excelBorderStyle;
            border.Bottom.Style = excelBorderStyle;
            border.Left.Style = excelBorderStyle;
            border.Right.Style = excelBorderStyle;
        }




        #region ExcelColumnMatch

        /// <summary>
        /// Excel表单匹配列
        /// </summary>
        public class ExcelColumnMatch
        {
            public HashSet<string> maybeColumnName { get; private set; }
            public string columnName { get; private set; }

            public System.Type type { get; private set; }

            public ExcelColumnMatch(string[] names, string columnName, System.Type? type = null)
            {
                this.maybeColumnName = names.ToHashSet();
                this.columnName = columnName;
                if (type is not null)
                {
                    this.type = type;
                }
                else
                {
                    this.type = typeof(string);
                }

            }
        }

        public static bool Status(this ExcelColumnMatch[] excelColumnMatches)
        {
            HashSet<string> columnNamesCheck = new HashSet<string>();

            foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatches)
            {
                foreach (string columnName in excelColumnMatch.maybeColumnName)
                {
                    if (columnNamesCheck.Contains(columnName))
                    {
                        return false;
                    }
                    else
                    {
                        columnNamesCheck.Add(columnName);
                    }
                }
            }

            return true;
        }



        public static DataTable GetNewDataTable(this ExcelColumnMatch[] excelColumnMatches,string tableName = Main.emptyString)
        {
            DataTable result = new DataTable(tableName);
            foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatches)
            {
                result.Columns.Add(excelColumnMatch.columnName, excelColumnMatch.type);
            }
            return result;
        }


        public static void ParseTable(this ExcelColumnMatch[] excelColumnMatchs, DataTable excelTable, DataTable table)
        {


            Dictionary<string, int> keyValuePairs = new Dictionary<string, int>();

            //进行字段匹配
            foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatchs)
            {
                for (int i = 0; i < excelTable.Columns.Count; i++)
                {
                    DataColumn column = excelTable.Columns[i];
                    if (excelColumnMatch.maybeColumnName.Contains(column.ColumnName))
                    {
                        keyValuePairs.Add(excelColumnMatch.columnName, i);
                        break;
                    }
                }
            }

            //检查字段是否匹配完全
            List<string> noMatchColumnNames = new List<string>();
            foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatchs)
            {
                if (!keyValuePairs.ContainsKey(excelColumnMatch.columnName))
                {
                    noMatchColumnNames.Add(excelColumnMatch.columnName);
                }
            }
            if (noMatchColumnNames.Count > 0)
            {
                KalevaAalto.Main.ThrowError($"工作表“{table.TableName}”找不到字段{noMatchColumnNames.Round()}");
            }



            int rowE = 1;
            string columnE = Main.emptyString;
            try
            {
                //添加数据
                foreach (DataRow excelRow in excelTable.Rows)
                {
                    rowE++;
                    columnE = Main.emptyString;
                    DataRow newRow = table.Rows.Add();
                    foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatchs)
                    {
                        columnE = excelColumnMatch.maybeColumnName.First();
                        var cell = excelRow[keyValuePairs[excelColumnMatch.columnName]];


                        if(excelColumnMatch.type == typeof(DateTime))
                        {
                            if (cell.GetType() == typeof(int) || cell.GetType() == typeof(double))
                            {
                                newRow[excelColumnMatch.columnName] = DateTime.FromOADate((double)cell);
                            }
                            else
                            {
                                newRow[excelColumnMatch.columnName] = DBNull.Value;
                            }
                        }
                        else if(excelColumnMatch.type == typeof(int))
                        {
                            if(cell is null || cell is DBNull)
                            {
                                newRow[excelColumnMatch.columnName] = 0;
                            }
                            else if (cell.GetType().IsOrNullableNumber())
                            {
                                newRow[excelColumnMatch.columnName] = Convert.ToInt32(cell);
                            }
                            else if (cell.GetType() == typeof(string))
                            {
                                newRow[excelColumnMatch.columnName] = cell?.ToString()?.StringToInt() ?? 0;
                            }
                            else
                            {
                                newRow[excelColumnMatch.columnName] = 0;
                            }
                        }
                        else if (excelColumnMatch.type == typeof(decimal))
                        {
                            if (cell is null || cell is DBNull)
                            {
                                newRow[excelColumnMatch.columnName] = 0;
                            }
                            else if (cell.GetType().IsOrNullableNumber())
                            {
                                newRow[excelColumnMatch.columnName] = Convert.ToDecimal(cell);
                            }
                            else if (cell.GetType() == typeof(string))
                            {
                                newRow[excelColumnMatch.columnName] = cell?.ToString()?.StringToDecimal() ?? 0m;
                            }
                            else
                            {
                                newRow[excelColumnMatch.columnName] = 0;
                            }
                        }
                        else if (excelColumnMatch.type == typeof(double))
                        {
                            if (cell is null || cell is DBNull)
                            {
                                newRow[excelColumnMatch.columnName] = 0;
                            }
                            else if (cell.GetType().IsOrNullableNumber())
                            {
                                newRow[excelColumnMatch.columnName] = Convert.ToDouble(cell);
                            }
                            else if (cell.GetType() == typeof(string))
                            {
                                newRow[excelColumnMatch.columnName] = cell?.ToString()?.StringToDouble() ?? 0.0;
                            }
                            else
                            {
                                newRow[excelColumnMatch.columnName] = 0.0;
                            }
                        }
                        else
                        {
                            newRow[excelColumnMatch.columnName] = cell;
                        }
                    }
                }
            }
            catch
            {
                throw new Exception($"错误数据：字段“{columnE}”，行号{rowE}；" );

            }

            


        }





        


        #endregion

    }
}
