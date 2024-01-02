using KalevaAalto.Models.Excel.Enums;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public abstract class IRange
    {


        public abstract RangePos Pos { get; }
        public int RowCount  => Pos.EndPos.Row - Pos.StartPos.Row + 1; 
        public int ColumnCount => Pos.EndPos.Column - Pos.StartPos.Column + 1; 
        public abstract IStyle Style { get; }
        public abstract ErrorType ErrorType { get; }
        public bool IsError => this.ErrorType != ErrorType.None;
        public IRow[] Rows
        {
            get
            {
                List<IRow> result = new List<IRow>(RowCount);
                for (int row = 1; row <= RowCount; row++)
                {
                    result.Add(this[row, 1].Row);
                }
                return result.ToArray();
            }
        }
        public IColumn[] Columns
        {
            get
            {
                List<IColumn> result = new List<IColumn>(RowCount);
                for (int row = 1; row <= RowCount; row++)
                {
                    result.Add(this[row, 1].Column);
                }
                return result.ToArray();
            }
        }
        public ICell this[CellPos cellPos]
        {
            get
            {
                if (cellPos.Column > ColumnCount || cellPos.Row > RowCount)
                {
                    throw new Exception(@"行号或列号越界");
                }
                return GetCell(cellPos);
            }
        }
        public abstract bool Merge { get; set; }
        public ICell this[int row, int column] => this[new CellPos(row, column)];
        public object? Value { get => this[CellPos.DefaultStartPos].Value; set => this[CellPos.DefaultStartPos].Value = value; }

        private readonly static Regex regexColumnName = new Regex(@"^(?<columnName>\.+?)(\d+)?$");
        private readonly static Regex regexIsNumber = new Regex(@"^\d*$");
        public System.Data.DataTable? DataTable
        {
            get
            {
                System.Data.DataTable result = new System.Data.DataTable();
                Dictionary<string, int> columnNumbers = new Dictionary<string, int>();

                int rowCount = RowCount, columnCount = ColumnCount;
                for (int column = 1; column <= columnCount; column++)
                {
                    ICell cell = this[1, column];
                    if (cell.Value == null) { continue; }
                    string? columnName = cell.Value.ToString();
                    if (string.IsNullOrEmpty(columnName)) { continue; }
                    columnName = columnName.Replace(@" ", @"");
                    if (regexIsNumber.IsMatch(columnName)) { continue; }
                    uint number = 0;
                    while (columnNumbers.ContainsKey(columnName))
                    {
                        number++;
                        Match match = regexIsNumber.Match(columnName);
                        string columnNameString = match.Groups[@"columnName"].Value;
                        columnName = columnNameString + number.ToString();
                    }
                    columnNumbers.Add(columnName, column);
                }


                if (columnNumbers.Count <= 0) { return null; }


                foreach (string columnName in columnNumbers.Keys)
                {
                    result.Columns.Add(columnName, typeof(object));
                }

                for (int row = 2; row <= rowCount; row++)
                {
                    DataRow dataRow = result.Rows.Add();
                    foreach (KeyValuePair<string, int> item in columnNumbers)
                    {
                        ICell cell = this[row, item.Value];
                        if (cell.Value == null)
                        {
                            dataRow[item.Key] = DBNull.Value;
                        }
                        else if (cell.IsError)
                        {
                            dataRow[item.Key] = cell.ErrorType;
                        }
                        else
                        {
                            dataRow[item.Key] = cell.Value;
                        }
                    }
                }

                return result;
            }
        }
        protected abstract ICell GetCell(CellPos cellPos);
        protected abstract IRange GetRange(RangePos rangePos);


        public virtual void SetTable(string tableName)=>throw new NotImplementedException();
        


    }
}
