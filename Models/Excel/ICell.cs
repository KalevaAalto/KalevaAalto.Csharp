using KalevaAalto.Models.Excel.Enums;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public abstract class ICell
    {
        public abstract CellPos Pos { get; }
        public abstract IStyle Style { get; }
        public abstract object? Value { get; set; }
        public abstract IRow Row { get; }
        public abstract IColumn Column { get; }
        public abstract ErrorType ErrorType { get; }
        public bool IsError => this.ErrorType != ErrorType.None;
    }
}
