﻿using KalevaAalto.Interfaces.Excel;
using KalevaAalto.Models.Excel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static KalevaAalto.Main;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    public class Worksheet : IWorksheet
    {
        private ExcelWorksheet worksheet { get; }


        public Worksheet(ExcelWorksheet excelWorksheet)
        {
            worksheet = excelWorksheet;
        }


        public override string Name { get => worksheet.Name; set => worksheet.Name = value; }

        public override IRange? Range
        {
            get
            {
                var dimension = worksheet.Dimension;
                return this.GetRange(new RangePos(Models.Excel.CellPos.DefaultStartPos, new Models.Excel.CellPos(dimension.Rows, dimension.Columns)));
            }
        }

        protected override ICell GetCell(Models.Excel.CellPos cellPos)
        {
            return new Cell(worksheet.Cells[cellPos.Row, cellPos.Column]);
        }

        protected override IRange GetRange(RangePos rangePos)
        {
            return new Range(this.worksheet.Cells[rangePos.StartPos.Row, rangePos.StartPos.Column, rangePos.EndPos.Row, rangePos.EndPos.Column]);
        }


        public override void Test()
        {
            // 设置打印方向为横向
            worksheet.PrinterSettings.Orientation = eOrientation.Landscape;

            // 设置纸张大小为A5
            worksheet.PrinterSettings.PaperSize = ePaperSize.A5;

            // 将工作表调整为一页
            worksheet.PrinterSettings.FitToPage = true;
            //worksheet.PrinterSettings.FitToWidth = 1;
            //worksheet.PrinterSettings.FitToHeight = 1;
        }



    }
}