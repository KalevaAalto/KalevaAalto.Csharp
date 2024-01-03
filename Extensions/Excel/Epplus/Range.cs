using KalevaAalto.Models.Excel;
using KalevaAalto.Models.Excel.Enums;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Range : IRange
    {
        private readonly ExcelWorksheet _worksheet;
        private ExcelRange _rng=> _worksheet.Cells[_pos.StartPos.Row, _pos.StartPos.Column, _pos.EndPos.Row, _pos.EndPos.Column];
        private readonly RangePos _pos;
        public override ErrorType ErrorType => Cell.GetErrorType(this.Value);

        internal Range(ExcelRange rng) 
        {
            _pos = new RangePos(new CellPos(rng.Start.Row, rng.Start.Column), new CellPos(rng.End.Row, rng.End.Column));
            _worksheet = rng.Worksheet;
        }

        public override IStyle Style => new Style(_rng.Style);

        public override RangePos Pos => _pos;

        protected override ICell GetCell(CellPos cellPos)=>new Cell(_rng[_pos.StartPos.Row + cellPos.Row - 1, _pos.StartPos.Column + cellPos.Column - 1]);
        

        protected override IRange GetRange(RangePos rangePos)=>
            new Range(_rng[_pos.StartPos.Row + rangePos.StartPos.Row -1, _pos.StartPos.Column + rangePos.StartPos.Column - 1, _pos.StartPos.Row + rangePos.EndPos.Row - 1, _pos.StartPos.Column + rangePos.EndPos.Column - 1]);
        
        public override bool Merge { get => _rng.Merge; set => _rng.Merge=value; }
        public override void SetTable(string tableName)=>_worksheet.Tables.Add(_rng, tableName);
        

    }
}
