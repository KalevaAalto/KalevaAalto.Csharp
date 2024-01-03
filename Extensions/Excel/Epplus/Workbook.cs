using KalevaAalto;
using KalevaAalto.Models.Excel;
using Npgsql.Replication;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Extensions.Excel.Epplus
{
    public class Workbook : IWorkbook
    {
        private ExcelPackage? _package;
        

        public Workbook(string fileName) : base(fileName) { }
        protected override void Init()
        {
            if (FileExist) _package = new ExcelPackage(new FileInfo(FileName));
            else _package = new ExcelPackage();
        }


        public override IWorksheet[] Worksheets { get => _package!.Workbook.Worksheets.Cast<ExcelWorksheet>().Select(it => new Worksheet(it)).ToArray(); }


        protected override void _Save()
        {
            _package!.SaveAs(FileName);
        }

        public override IWorksheet AddWorksheet(string name)
        {
            return new Worksheet(_package!.Workbook.Worksheets.Add(name));
        }

        public override void Dispose()
        {
            _package!.Dispose();
        }
    }
}
