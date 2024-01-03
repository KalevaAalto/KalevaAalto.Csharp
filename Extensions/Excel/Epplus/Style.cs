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
using System.Collections.Immutable;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    internal class Style : IStyle
    {
        private readonly ExcelStyle _style;


        internal Style(ExcelStyle style)
        {
            _style = style;
        }

        public override BorderStyle BorderStyle
        {
            get
            {
                switch (_style.Border.Top.Style)
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
                        _style.Border.Top.Style = ExcelBorderStyle.Thin;
                        _style.Border.Bottom.Style = ExcelBorderStyle.Thin;
                        _style.Border.Left.Style = ExcelBorderStyle.Thin;
                        _style.Border.Right.Style = ExcelBorderStyle.Thin;
                        break;
                    case BorderStyle.Thick:
                        _style.Border.Top.Style = ExcelBorderStyle.Thick;
                        _style.Border.Bottom.Style = ExcelBorderStyle.Thick;
                        _style.Border.Left.Style = ExcelBorderStyle.Thick;
                        _style.Border.Right.Style = ExcelBorderStyle.Thick;
                        break;
                    default:
                        _style.Border.Top.Style = ExcelBorderStyle.None;
                        _style.Border.Bottom.Style = ExcelBorderStyle.None;
                        _style.Border.Left.Style = ExcelBorderStyle.None;
                        _style.Border.Right.Style = ExcelBorderStyle.None;
                        break;

                }
            }
        }


        public override Color BorderColor
        {
            get => Color.Black;
            set
            {
                _style.Border.Top.Color.SetColor(value);
                _style.Border.Bottom.Color.SetColor(value);
                _style.Border.Left.Color.SetColor(value);
                _style.Border.Right.Color.SetColor(value);
            }
        }

        public override Color FontColor { get => Color.Black; set => _style.Font.Color.SetColor(value); }

        public override double FontSize { get => _style.Font.Size; set => _style.Font.Size = (float)value; }
        public override FontWeight FontWeight
        {
            get => _style.Font.Bold ? FontWeight.Thick : FontWeight.Thin;
            set
            {
                switch (value)
                {
                    case FontWeight.Thin:
                        _style.Font.Bold = false;
                        break;
                    case FontWeight.Thick:
                        _style.Font.Bold = true;
                        break;
                    default:
                        _style.Font.Bold = false;
                        break;
                }
            }
        }

        public override string FontFamily { get => _style.Font.Name; set => _style.Font.Name = value; }


        private readonly static ImmutableDictionary<ExcelHorizontalAlignment, HorizontalAlignment> s_horizontalAlignments = new Dictionary<ExcelHorizontalAlignment, HorizontalAlignment>
        {
            {ExcelHorizontalAlignment.Left, HorizontalAlignment.Left},
            {ExcelHorizontalAlignment.Center, HorizontalAlignment.Center},
            {ExcelHorizontalAlignment.Right, HorizontalAlignment.Right},
        }.ToImmutableDictionary();
        public override HorizontalAlignment HorizontalAlignment
        {
            get => s_horizontalAlignments[_style.HorizontalAlignment];
            set => _style.HorizontalAlignment = s_horizontalAlignments.First(it => it.Value == value).Key;
        }
        private readonly static ImmutableDictionary<ExcelVerticalAlignment, VerticalAlignment> s_verticalAlignments = new Dictionary<ExcelVerticalAlignment, VerticalAlignment>
        {
            {ExcelVerticalAlignment.Bottom, VerticalAlignment.Bottom},
            {ExcelVerticalAlignment.Center, VerticalAlignment.Center},
            {ExcelVerticalAlignment.Top, VerticalAlignment.Top},
        }.ToImmutableDictionary();

        public override VerticalAlignment VerticalAlignment
        {
            get => s_verticalAlignments[_style.VerticalAlignment];
            set => _style.VerticalAlignment = s_verticalAlignments.First(it => it.Value == value).Key;
        }

        public override string NumberFormatString { get => _style.Numberformat.Format; set => _style.Numberformat.Format = value; }
        public override bool WarpText { get => _style.WrapText; set => _style.WrapText = value; }
        public override UnderLineType UnderLineType 
        { 
            get => _style.Font.UnderLine? UnderLineType.Solid: UnderLineType.None;
            set
            {
                switch (value)
                {
                    case UnderLineType.None:_style.Font.UnderLine = false;break;
                    case UnderLineType.Solid:_style.Font.UnderLine = true;break;
                    default:throw new NotFiniteNumberException();
                }
            }
        }
    }
}
