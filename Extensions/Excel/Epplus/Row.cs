using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel;
using OfficeOpenXml;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Row : IRow
    {
        private readonly ExcelRangeRow row;
        public Row(ExcelRangeRow row)
        {
            this.row = row;
        }
        public override double Height { get => row.Height; set => row.Height = value; }

        public override int Pos => row.EndRow;

        public override IStyle Style => new Style(row.Style);
    }
}
