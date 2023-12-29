using Microsoft.ML;
using OfficeOpenXml.Style;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using KalevaAalto.Models.Excel.Enums;
using KalevaAalto.Models.Excel;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Style : IStyle
    {
        private readonly ExcelStyle style;


        internal Style(ExcelStyle style)
        {
            this.style = style;
        }

        public override BorderStyle BorderStyle
        {
            get
            {
                switch (style.Border.Top.Style)
                {
                    case ExcelBorderStyle.Thick:
                        return BorderStyle.Thick;
                    case ExcelBorderStyle.Thin:
                        return BorderStyle.Thin;
                    default:
                        return BorderStyle.None;
                }
            }
            set
            {
                switch (value)
                {
                    case BorderStyle.Thin:
                        style.Border.Top.Style = ExcelBorderStyle.Thin;
                        style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        style.Border.Left.Style = ExcelBorderStyle.Thin;
                        style.Border.Right.Style = ExcelBorderStyle.Thin;
                        break;
                    case BorderStyle.Thick:
                        style.Border.Top.Style = ExcelBorderStyle.Thick;
                        style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                        style.Border.Left.Style = ExcelBorderStyle.Thick;
                        style.Border.Right.Style = ExcelBorderStyle.Thick;
                        break;
                    default:
                        style.Border.Top.Style = ExcelBorderStyle.None;
                        style.Border.Bottom.Style = ExcelBorderStyle.None;
                        style.Border.Left.Style = ExcelBorderStyle.None;
                        style.Border.Right.Style = ExcelBorderStyle.None;
                        break;

                }
            }
        }


        public override Color BorderColor
        {
            get => Color.Black;
            set
            {
                style.Border.Top.Color.SetColor(value);
                style.Border.Bottom.Color.SetColor(value);
                style.Border.Left.Color.SetColor(value);
                style.Border.Right.Color.SetColor(value);
            }
        }

        public override Color FontColor { get => Color.Black; set => style.Font.Color.SetColor(value); }

        public override double FontSize { get => style.Font.Size; set => style.Font.Size = (float)value; }
        public override FontWeight FontWeight
        {
            get => style.Font.Bold ? FontWeight.Thick : FontWeight.Thin;
            set
            {
                switch (value)
                {
                    case FontWeight.Thin:
                        style.Font.Bold = false;
                        break;
                    case FontWeight.Thick:
                        style.Font.Bold = true;
                        break;
                    default:
                        style.Font.Bold = false;
                        break;
                }
            }
        }

        public override string FontFamily { get => style.Font.Name; set => style.Font.Name = value; }


        private readonly static Dictionary<ExcelHorizontalAlignment, HorizontalAlignment> horizontalAlignments = new Dictionary<ExcelHorizontalAlignment, HorizontalAlignment>
        {
            {ExcelHorizontalAlignment.Left, HorizontalAlignment.Left},
            {ExcelHorizontalAlignment.Center, HorizontalAlignment.Center},
            {ExcelHorizontalAlignment.Right, HorizontalAlignment.Right},
        };
        public override HorizontalAlignment HorizontalAlignment
        {
            get => horizontalAlignments[style.HorizontalAlignment];
            set => style.HorizontalAlignment = horizontalAlignments.First(it => it.Value == value).Key;
        }
        private readonly static Dictionary<ExcelVerticalAlignment, VerticalAlignment> verticalAlignments = new Dictionary<ExcelVerticalAlignment, VerticalAlignment>
        {
            {ExcelVerticalAlignment.Bottom, VerticalAlignment.Bottom},
            {ExcelVerticalAlignment.Center, VerticalAlignment.Center},
            {ExcelVerticalAlignment.Top, VerticalAlignment.Top},
        };

        public override VerticalAlignment VerticalAlignment
        {
            get => verticalAlignments[style.VerticalAlignment];
            set => style.VerticalAlignment = verticalAlignments.First(it => it.Value == value).Key;
        }

        public override string NumberFormatString { get => style.Numberformat.Format; set => style.Numberformat.Format = value; }
        public override bool WarpText { get => style.WrapText; set => style.WrapText = value; }

    }
}
