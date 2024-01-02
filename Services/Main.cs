using KalevaAalto.Models.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto
{
    public static partial class Static
    {
        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="fileName">工作簿路径</param>
        public static IWorkbook GetWorkbook(string fileName) => new Extensions.Excel.Epplus.Workbook(fileName);

        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="path">工作簿所在文件夹路径</param>
        /// <param name="name">工作簿名称</param>
        public static IWorkbook GetWorkbook(string path, string name) => new Extensions.Excel.Epplus.Workbook(Path.Combine(path, $"{name}.xlsx"));


        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="fileName">工作簿路径</param>
        public static async Task<IWorkbook> GetWorkbookAsync(string fileName) =>await Task.Run(() => { return new Extensions.Excel.Epplus.Workbook(fileName); });

        /// <summary>
        /// 获取工作簿
        /// </summary>
        /// <param name="path">工作簿所在文件夹路径</param>
        /// <param name="name">工作簿名称</param>
        public static async Task<IWorkbook> GetWorkbookAsync(string path, string name) => await Task.Run(() => { return new Extensions.Excel.Epplus.Workbook(Path.Combine(path, $"{name}.xlsx")); });
    }
}
