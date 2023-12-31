﻿using Google.Protobuf.WellKnownTypes;
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
using System.Xml.Linq;
using System.Xml;
using KalevaAalto.Models;


namespace KalevaAalto;

public static partial class Static
{






    /// <summary>
    /// 从文件中获取Excel文档
    /// </summary>
    public static ExcelPackage GetExcelPackage(this Models.FileSystem.FileNameInfo fileNameInfo)
    {
        return new ExcelPackage(fileNameInfo.FileInfo);
    }


    /// <summary>
    /// 将Excel文档另存为至文件中
    /// </summary>
    /// <param name="package">要保存的Excel文档</param>
    public static void SaveAs(this ExcelPackage package,Models.FileSystem.FileNameInfo fileNameInfo)
    {
        package.SaveAs(fileNameInfo.FileInfo);
    }











    /// <summary>
    /// 检查单元格是否为空
    /// </summary>
    /// <param name="worksheet">工作表</param>
    /// <param name="cellPos">单元格</param>
    /// <returns>返回单元格是否为空</returns>
    public static bool IsCellEmpty(this ExcelWorksheet worksheet,Models.Excel.CellPos cellPos)
    {

        object? value = worksheet.GetValue(cellPos.Row, cellPos.Column);
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
        return worksheet.IsCellEmpty(new Models.Excel.CellPos(row,column));
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
    public static Models.Excel.CellPos GetStartCellPos(this ExcelWorksheet worksheet)
    {
        for(int row = 0; row<500; row++)
        {
            for (int column = 0; column < 500; column++)
            {
                if (!worksheet.IsCellEmpty(row,column))
                {
                    return new Models.Excel.CellPos(row, column);
                }
            }
        }


        throw new Exception(@"未找到起始坐标；");
    }

    /// <summary>
    /// 获取工作表的起始单元格坐标
    /// </summary>
    /// <param name="worksheet">工作表</param>
    /// <param name="startString">起始点标志字符串</param>
    /// <returns>返回工作表的起始单元格坐标</returns>
    public static Models.Excel.CellPos GetStartCellPos(this ExcelWorksheet worksheet,string startString)
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
                    return new Models.Excel.CellPos(row, column);
                }
            }
        }


        throw new Exception(@"未找到起始坐标；");
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
                int matchNumber = StringToInt(match.Groups[2].Value);
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



    public static DataTable GetNewDataTable(this ExcelColumnMatch[] excelColumnMatches,string tableName = EmptyString)
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
            throw new Exception($"工作表“{table.TableName}”找不到字段{noMatchColumnNames.Round()}");
        }



        int rowE = 1;
        string columnE = EmptyString;
        try
        {
            //添加数据
            foreach (DataRow excelRow in excelTable.Rows)
            {
                rowE++;
                columnE = EmptyString;
                DataRow newRow = table.Rows.Add();
                foreach (ExcelColumnMatch excelColumnMatch in excelColumnMatchs)
                {
                    ObjectConvert objectConvert = new ObjectConvert(excelColumnMatch.type);
                    columnE = excelColumnMatch.maybeColumnName.First();
                    object? cell = excelRow[keyValuePairs[excelColumnMatch.columnName]];
                    object? obj = objectConvert.GetValue(cell);
                    

                    if(cell == null || cell == DBNull.Value || obj == null | obj == DBNull.Value)
                    {
                        newRow[excelColumnMatch.columnName] = DBNull.Value;
                    }
                    else
                    {
                        newRow[excelColumnMatch.columnName] = obj;
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
