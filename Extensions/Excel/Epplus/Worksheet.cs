using KalevaAalto.Models.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    public class Worksheet : IWorksheet
    {
        private ExcelWorksheet _worksheet;


        public Worksheet(ExcelWorksheet excelWorksheet)
        {
            _worksheet = excelWorksheet;
        }


        public override string Name { get => _worksheet.Name; set => _worksheet.Name = value; }

        public override IRange? Range
        {
            get
            {
                var dimension = _worksheet.Dimension;
                return this.GetRange(new RangePos(Models.Excel.CellPos.DefaultStartPos, new Models.Excel.CellPos(dimension.Rows, dimension.Columns)));
            }
        }

        protected override ICell GetCell(Models.Excel.CellPos cellPos)
        {
            return new Cell(_worksheet.Cells[cellPos.Row, cellPos.Column]);
        }

        protected override IRange GetRange(RangePos rangePos)
        {
            return new Range(_worksheet.Cells[rangePos.StartPos.Row, rangePos.StartPos.Column, rangePos.EndPos.Row, rangePos.EndPos.Column]);
        }


        public override void Test()
        {
            // 设置打印方向为横向
            _worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

            // 设置纸张大小为A5
            _worksheet.PrinterSettings.PaperSize = ePaperSize.A5;

            // 设置页面居中
            _worksheet.PrinterSettings.HorizontalCentered = true;
            //worksheet.PrinterSettings.VerticalCentered = true;

            // 将工作表调整为一页
            _worksheet.PrinterSettings.FitToPage = true;
            //worksheet.PrinterSettings.FitToWidth = 1;
            //worksheet.PrinterSettings.FitToHeight = 1;
        }


        public override void CleanDataValidation()
        {
            throw new NotImplementedException();
        }

        public override void CleanDrawings()
        {
            throw new NotImplementedException();
        }

        public override void ClearFormulas()
        {
            throw new NotImplementedException();
        }



    }
}
