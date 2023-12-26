using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel;

namespace KalevaAalto.Interfaces.Excel
{
    public abstract class IWorksheet
    {
        public abstract string Name { get; set; }
        public ICell this[CellPos cellPos] =>this.GetCell(cellPos);
        public ICell this[int row,int column] =>this.GetCell(new CellPos(row,column));
        public IRange this[RangePos rangePos]=>this.GetRange(rangePos);
        public IRange this[CellPos startPos,CellPos endPos]=>this.GetRange(new RangePos(startPos,endPos));
        public IRange this[int startRow,int startColumn,int endRow,int endColumn] => this[new CellPos(startRow,startColumn),new CellPos(endRow,endColumn)];
        

        public IRange this[string address] => this[new RangePos(address)];

        public string Description { get; set; } = string.Empty;

        public virtual IRange? Range
        {
            get 
            {
                int maxRow = 0, maxColum = 0;
                for(int row=1;row <= CellPos.XlsxMaxRow; row++)
                {
                    for(int column=1;column<= CellPos.XlsxMaxColumn; column++)
                    {
                        ICell? cell = this[row,column];
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
                return this[1,1,maxRow, maxColum];
            }
        }



        protected abstract ICell GetCell(CellPos cellPos);
        protected abstract IRange GetRange(RangePos rangePos);
        protected IRange GetRange(CellPos cellPos)
        {
            return this.GetRange(new RangePos(cellPos, cellPos));
        }

        public IRow Row(int row)
        {
            return this[row, 1].Row;
        }

        public IColumn Column(int column)
        {
            return this[1, column].Column;
        }


        

        public int MaxRow { get => this.Range?.RowCount ?? 0; }
        public int MaxColumn { get=>this.Range?.ColumnCount ?? 0; }


        public DataTable? ToDataTable(CellPos cellPos)
        {
            IRange? range = this[new RangePos(cellPos, new CellPos(this.MaxRow, this.MaxColumn))];
            if(range == null) { return null; }
            DataTable? result = range.DataTable;
            if(result == null) { return null; }
            result.TableName = this.Name;
            return result;
        }

        
        public virtual void Test()
        {
            throw new Exception($"函数“void Test()”并未实现；");
        }






    }
}
