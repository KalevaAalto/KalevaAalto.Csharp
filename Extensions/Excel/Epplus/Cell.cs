using KalevaAalto.Models.Excel;
using KalevaAalto.Models.Excel.Enums;
using MySqlX.XDevAPI.Relational;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
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


        private ExcelRange rng;


        public Cell(ExcelRange rng) { this.rng = rng; }

        public override IStyle Style => new Style(rng.Style);

        public override CellPos Pos => new CellPos(rng.Start.Row, rng.Start.Column);

        public override object? Value { get => rng.Value; set => rng.Value = value; }

        public override IColumn Column => new Column(rng.EntireColumn);
        public override IRow Row => new Row(rng.EntireRow);

        public override ErrorType ErrorType => GetErrorType(this.Value);
    }
}
