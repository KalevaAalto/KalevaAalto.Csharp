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
        private readonly ExcelRangeRow _row;
        public Row(ExcelRangeRow row)
        {
            _row = row;
        }
        public override double Height { get => _row.Height; set => _row.Height = value; }

        public override int Pos => _row.EndRow;

        public override IStyle Style => new Style(_row.Style);
    }
}
