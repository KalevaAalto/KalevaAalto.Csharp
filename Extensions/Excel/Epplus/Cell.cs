using KalevaAalto.Interfaces.Excel;
using KalevaAalto.Models.Excel;
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
        private ExcelRange rng;


        public Cell(ExcelRange rng) { this.rng = rng; }

        public override IStyle Style => new Style(rng.Style);

        public override CellPos Pos => new CellPos(rng.Start.Row, rng.Start.Column);

        public override object? Value { get => rng.Value; set => rng.Value = value; }

        public override IColumn Column => new Column(rng.EntireColumn);
        public override IRow Row => new Row(rng.EntireRow);


    }
}
