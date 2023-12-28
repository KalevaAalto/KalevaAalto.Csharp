using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto;

public static class Serives
{
    /// <summary>
    /// 获取工作簿
    /// </summary>
    /// <param name="fileName">工作簿路径</param>
    public static Interfaces.Excel.IWorkbook GetWorkbook(string fileName)=>new Extensions.Excel.Epplus.Workbook(fileName);

    /// <summary>
    /// 获取工作簿
    /// </summary>
    /// <param name="path">工作簿所在文件夹路径</param>
    /// <param name="name">工作簿名称</param>
    /// <returns></returns>
    public static Interfaces.Excel.IWorkbook GetWorkbook(string path,string name)=>new Extensions.Excel.Epplus.Workbook(Path.Combine(path,$"{name}.xlsx"));

}
