using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel;

namespace KalevaAalto.Interfaces.Excel
{
    public abstract class ICell
    {
        public abstract CellPos Pos { get; }
        public abstract IStyle Style { get; }
        public abstract object? Value { get; set; }
        public abstract IRow Row { get; }
        public abstract IColumn Column { get; }
    }
}
