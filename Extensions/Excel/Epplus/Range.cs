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
        private readonly ExcelWorksheet worksheet;
        private ExcelRange rng=> this.worksheet.Cells[this.pos.StartPos.Row,this.pos.StartPos.Column,this.pos.EndPos.Row,this.pos.EndPos.Column];
        private readonly RangePos pos;
        public override ErrorType ErrorType => Cell.GetErrorType(this.Value);

        internal Range(ExcelRange rng) 
        {
            this.pos = new RangePos(new CellPos(rng.Start.Row, rng.Start.Column), new CellPos(rng.End.Row, rng.End.Column));
            this.worksheet = rng.Worksheet;
        }

        public override IStyle Style => new Style(rng.Style);

        public override RangePos Pos => this.pos;

        protected override ICell GetCell(CellPos cellPos)
        {
            return new Cell(this.rng[this.pos.StartPos.Row + cellPos.Row - 1, this.pos.StartPos.Column + cellPos.Column - 1]);
        }

        protected override IRange GetRange(RangePos rangePos)
        {
            return new Range(this.rng[this.pos.StartPos.Row + rangePos.StartPos.Row -1, this.pos.StartPos.Column + rangePos.StartPos.Column - 1, this.pos.StartPos.Row + rangePos.EndPos.Row - 1, this.pos.StartPos.Column + rangePos.EndPos.Column - 1]);
        }
        public override bool Merge { get => this.rng.Merge; set => this.rng.Merge=value; }
        public override void SetTable(string tableName)
        {
            this.rng.Worksheet.Tables.Add(this.rng,tableName);
        }

    }
}
