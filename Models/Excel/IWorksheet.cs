using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Extensions.Excel.Epplus;
using MySqlX.XDevAPI.Common;
using MySqlX.XDevAPI.Relational;

namespace KalevaAalto.Models.Excel
{
    public abstract class IWorksheet
    {
        public abstract string Name { get; set; }
        public ICell this[CellPos cellPos] => GetCell(cellPos);
        public ICell this[int row, int column] => GetCell(new CellPos(row, column));
        public IRange this[RangePos rangePos] => GetRange(rangePos);
        public IRange this[CellPos startPos, CellPos endPos] => GetRange(new RangePos(startPos, endPos));
        public IRange this[int startRow, int startColumn, int endRow, int endColumn] => this[new CellPos(startRow, startColumn), new CellPos(endRow, endColumn)];
        public IRange this[string address] => this[new RangePos(address)];

        public string Description { get; set; } = string.Empty;

        public virtual IRange? Range
        {
            get
            {
                int maxRow = 0, maxColum = 0;
                for (int row = 1; row <= CellPos.XlsxMaxRow; row++)
                {
                    for (int column = 1; column <= CellPos.XlsxMaxColumn; column++)
                    {
                        ICell? cell = this[row, column];
                        if (cell == null) continue;
                        if (cell.Value == null) continue;
                        if (string.IsNullOrEmpty(cell.Value.ToString())) continue;
                        if (row > maxRow) maxRow = row;
                        if (column > maxColum) maxColum = column;
                    }
                }

                if (maxColum <= 0 || maxColum <= 0)
                {
                    return null;
                }
                return this[1, 1, maxRow, maxColum];
            }
        }

        protected abstract ICell GetCell(CellPos cellPos);
        protected abstract IRange GetRange(RangePos rangePos);
        protected IRange GetRange(CellPos cellPos)
        {
            return GetRange(new RangePos(cellPos, cellPos));
        }

        public IRow Row(int row)
        {
            return this[row, 1].Row;
        }

        public IColumn Column(int column)
        {
            return this[1, column].Column;
        }




        public int MaxRow { get => Range?.RowCount ?? 0; }
        public int MaxColumn { get => Range?.ColumnCount ?? 0; }


        public System.Data.DataTable? GetDataTable(CellPos cellPos)
        {
            return GetDataTable(new RangePos(cellPos, new CellPos(MaxRow, MaxColumn)));
        }


        public System.Data.DataTable? GetDataTable(RangePos rangePos)
        {
            System.Data.DataTable? result = this[rangePos].DataTable;
            if (result == null) { return null; }
            result.TableName = Name;
            return result;
        }


        public CellPos? SearchString(string searchString)
        {
            for (int w = 0; w < Math.Max(CellPos.XlsxMaxColumn, CellPos.XlsxMaxRow); w++)
            {
                int row = 1, column = w;
                if (column <= CellPos.XlsxMaxColumn)
                {
                    while (row <= w)
                    {
                        ICell cell = this[row, column];
                        if (cell != null && cell.Value != null && cell.Value.ToString() == searchString)
                        {
                            return new CellPos(row, column);
                        }
                        row++;
                    }
                }
                else
                {
                    column = CellPos.XlsxMaxColumn;
                }

                row--;

                while (column > 0)
                {
                    ICell cell = this[row, column];
                    if (cell != null && cell.Value != null && cell.Value.ToString() == searchString)
                    {
                        return new CellPos(row, column);
                    }
                    column--;
                }
            }

            return null;
        }
        public System.Data.DataTable? GetDataTable(string startString, string checkColumnName = EmptyString)
        {
            CellPos? startCellPos = SearchString(startString);
            if (startCellPos == null) { throw new Exception($"找不到起始标志“{startString}”；"); }
            int checkColumn = Notfound;
            if (string.IsNullOrEmpty(checkColumnName))
            {
                for (int column = startCellPos.Column; column <= CellPos.XlsxMaxColumn; column++)
                {
                    ICell cell = this[startCellPos.Row, column];
                    if (cell != null && cell.Value != null && cell.Value.ToString() == checkColumnName)
                    {
                        checkColumn = column;
                        break;
                    }
                }
            }
            if (checkColumn == Notfound) { return GetDataTable(startCellPos); }

            for (int endRow = startCellPos.Row; endRow <= CellPos.XlsxMaxRow; endRow++)
            {
                ICell cell = this[startCellPos.Row, checkColumn];
                if (cell == null || cell.Value == null)
                {
                    return GetDataTable(new RangePos(startCellPos, new CellPos(endRow, CellPos.XlsxMaxColumn)));
                }
            }
            return null;


        }






        public abstract void Test();


        /// <summary>
        /// 清除公式
        /// </summary>
        public abstract void ClearFormulas();
        /// <summary>
        /// 清除图片对象
        /// </summary>
        public abstract void CleanDrawings();
        /// <summary>
        /// 清除数据验证
        /// </summary>
        public abstract void CleanDataValidation();




    }
}
