using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Interfaces.Excel
{
    public abstract class IColumn
    {
        public abstract int Pos { get; }
        public abstract double Width { get; set; }
        public abstract IStyle Style { get; }
    }
}
