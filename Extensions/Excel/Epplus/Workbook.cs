using KalevaAalto;
using KalevaAalto.Interfaces.Excel;
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
        private ExcelPackage? package = null;
        

        public Workbook(string fileName) : base(fileName) { }
        protected override void Init()
        {
            if (FileExist)
            {
                package = new ExcelPackage(new FileInfo(FileName));
            }
            else
            {
                package = new ExcelPackage();
            }
        }


        public override IWorksheet[] Worksheets { get => package!.Workbook.Worksheets.Cast<ExcelWorksheet>().Select(it => new Worksheet(it)).ToArray(); }


        protected override void _Save()
        {
            package!.SaveAs(FileName);
        }

        public override IWorksheet AddWorksheet(string name)
        {
            return new Worksheet(package!.Workbook.Worksheets.Add(name));
        }

        public override void Dispose()
        {
            this.package?.Dispose();
        }
    }
}
