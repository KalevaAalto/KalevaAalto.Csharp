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
        private ExcelRangeColumn column;
        public Column(ExcelRangeColumn column)
        {
            this.column = column;
        }

        public override int Pos => column.EndColumn;
        public override IStyle Style => new Style(column.Style);
        public override double Width { get => column.Width; set => column.Width = value ; }

    }
}
