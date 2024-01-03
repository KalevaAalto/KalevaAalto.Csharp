using KalevaAalto.Models.Excel;
using KalevaAalto.Models.Excel.Enums;
using MySqlX.XDevAPI.Relational;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Cell : ICell
    {
        public static ErrorType GetErrorType(object? obj)
        {
            if(obj is null) { return ErrorType.None; }

            if(obj is OfficeOpenXml.ExcelErrorValue error)
            {
                
                switch (error.Type)
                {
                    case eErrorType.Div0:return ErrorType.Div0;
                    case eErrorType.NA: return ErrorType.NA;
                    case eErrorType.Name: return ErrorType.Name;
                    case eErrorType.Null: return ErrorType.Null;
                    case eErrorType.Num: return ErrorType.Num;
                    case eErrorType.Calc: return ErrorType.Calc;
                    case eErrorType.Ref: return ErrorType.Ref;
                    case eErrorType.Spill: return ErrorType.Spill;
                }
            }



            return ErrorType.None;
        }


        private ExcelRange _cell;


        public Cell(ExcelRange rng) { _cell = rng; }

        public override IStyle Style => new Style(_cell.Style);

        public override CellPos Pos => new CellPos(_cell.Start.Row, _cell.Start.Column);

        public override object? Value { get => _cell.Value; set => _cell.Value = value; }

        public override IColumn Column => new Column(_cell.EntireColumn);
        public override IRow Row => new Row(_cell.EntireRow);

        public override ErrorType ErrorType => GetErrorType(Value);
    }
}
