using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel.Enums;

namespace KalevaAalto.Models.Excel
{

    public abstract class IStyle
    {
        public abstract Color BorderColor { get; set; }
        public abstract BorderStyle BorderStyle { get; set; }
        public abstract Color FontColor { get; set; }
        public abstract FontWeight FontWeight { get; set; }
        public abstract HorizontalAlignment HorizontalAlignment { get; set; }
        public abstract VerticalAlignment VerticalAlignment { get; set; }
        public abstract string NumberFormatString { get; set; }
        /// <summary>
        /// 自动换行
        /// </summary>
        public abstract bool WarpText { get; set; }
        public abstract double FontSize { get; set; }
        public abstract string FontFamily { get; set; }
        public abstract UnderLineType UnderLineType { get; set; }
    }
}
