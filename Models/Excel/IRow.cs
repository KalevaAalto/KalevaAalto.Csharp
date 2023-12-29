using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public abstract class IRow
    {
        public abstract int Pos { get; }
        public abstract double Height { get; set; }
        public abstract IStyle Style { get; }



    }
}
