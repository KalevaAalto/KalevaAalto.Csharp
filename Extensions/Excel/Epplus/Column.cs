using KalevaAalto.Models.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Column : IColumn
    {
        private ExcelRangeColumn _column;
        public Column(ExcelRangeColumn column)
        {
            _column = column;
        }

        public override int Pos => _column.EndColumn;
        public override IStyle Style => new Style(_column.Style);
        public override double Width { get => _column.Width; set => _column.Width = value ; }

    }
}
