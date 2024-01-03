using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using SqlSugar;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlTypes;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection.Metadata;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace KalevaAalto;
/// <summary>
/// KalevaAalto个人的常用库
/// </summary>
public static partial class Static
{
        
    //未找到
    public const int Notfound = -1;

    public readonly static Regex RegexMonthString = new Regex(@"(?<year>\d{4})[年\-\/]?(?<month>\d{1,2})[月]?");
    public readonly static Regex RegexNumber = new Regex(@"[+-]?\d+(\.\d+)?");

    /// <summary>
    /// 获取一个带时间的测试名称
    /// </summary>
    public static string TestName
    {
        get
        {
            return @"test" + NowNumberString + GetRandomString(@"abcdefghijklmnopqrstuvwxyz",16);
        }
    }



    /// <summary>
    /// 进程初始化
    /// </summary>
    public static void ProcessInit()
    {
        //注册更多的字符编码集
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        
        //注册Epplus
        OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial;
    }


    /// <summary>
    /// 进程初始化
    /// </summary>
    public static async Task ProcessInitAsync()
    {
        await Task.WhenAll(
            //注册更多的字符编码集
            Task.Run(()=> Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)),
            //注册Epplus
            Task.Run(() => OfficeOpenXml.ExcelPackage.LicenseContext = OfficeOpenXml.LicenseContext.NonCommercial)
            );
    }









    #region 字符串


    public const string EmptyString = @"";


    public static bool IsNullOrEmpty(this string? str)
    {
        return string.IsNullOrEmpty(str);
    }


    public static string Reverse(this string str)
    {
        StringBuilder result = new StringBuilder();

        for(int i = str.Length- 1; i >= 0; i--)
        {
            result.Append(str[i]);
        }

        return result.ToString();
    }


    public static string RegexFormat(string str, GroupCollection groups)
    {
        string rs = str;
        rs = rs.Replace("$$", groups[0].Value);
        for (int i = 0; i < groups.Count; i++) rs = rs.Replace($"${i.ToString()}", groups[i].Value);
        foreach (Group group in groups) rs = rs.Replace("${" + group.Name + "}", group.Value);
        return rs;
    }

    /// <summary>
    /// 查找字符串中有多少个子字符串
    /// </summary>
    /// <param name="str">要查找的字符串</param>
    /// <param name="subStr">要查找的子字符串</param>
    /// <returns>返回str中包含的subStr的个数</returns>
    public static int SubStringCount(this string str, string subStr)
    {
        if(str.Length == 0 || subStr.Length == 0 || subStr.Length > str.Length)
        {
            return 0;
        }

        int count = 0;
        int index = 0;

        while (true)
        {
            index = str.IndexOf(subStr, index);
            if (index == -1)
            {
                break;
            }
            count++;
            index += subStr.Length;  // 下次查找的起始位置
        }
        return count;
    }



    private readonly static HashSet<char> permittedFullAngleChars = new HashSet<char> { 
        '，', '？', '“', '”', '：' ,
    };

    /// <summary>
    /// 将字符串中的所有全角字符转换为半角字符
    /// </summary>
    /// <param name="str">要转换的字符串</param>
    /// <returns>转换后的字符串</returns>
    public static string TurnHalfChar(this string str)
    {
        StringBuilder rs = new StringBuilder();
        for (int i = 0; i < str.Length; i++)
        {
            char c = str[i];
            if (!permittedFullAngleChars.Contains(c) && c >= 65281 && c <= 65374) // 如果是全角字符
            {
                rs.Append((char)(c - 65248));// 将其转换为对应的半角字符
            }
            else
            {
                rs.Append(c);
            }
        }
        return rs.ToString();
    }

    public static string RegexReplace(this string str,string pattern,string replacement)
    {
        return Regex.Replace(str,pattern,replacement);
    }

    public static string Join(this string str, string[] strings)
    {
        StringBuilder sb = new StringBuilder();
        foreach (string s in strings)
        {
            sb.Append(s);
            sb.Append(str);
        }
        if (sb.Length > 0)
        {
            sb.Remove(sb.Length - str.Length, str.Length);
        }

        return sb.ToString();
    }

    public static string Join(this string str, List<string> strings)
    {
        return str.Join(strings.ToArray());
    }

    public static string Join(this string str, HashSet<string> strings)
    {
        return str.Join(strings.ToArray());
    }

    /// <summary>
    /// 将字符串数组以字符串的方式呈现出来
    /// </summary>
    /// <param name="strs">要呈现的字符串数组</param>
    /// <returns>返回数组的字符串形式</returns>
    public static string Round(this string[] strs)
    {
        StringBuilder result = new StringBuilder();
        foreach(string str in strs)
        {
            result.Append('【');
            result.Append(str);
            result.Append('】');
        }
        return result.ToString();
    }


    /// <summary>
    /// 将字符串可变数组以字符串的方式呈现出来
    /// </summary>
    /// <param name="strs">要呈现的字符串可变数组</param>
    /// <returns>返回可变数组的字符串形式</returns>
    public static string Round(this List<string> strs)
    {
        StringBuilder result = new StringBuilder();
        foreach (string str in strs)
        {
            result.Append('【');
            result.Append(str);
            result.Append('】');
        }
        return result.ToString();
    }


    /// <summary>
    /// 将字符串哈希集以字符串的方式呈现出来
    /// </summary>
    /// <param name="strs">要呈现的字符串哈希集</param>
    /// <returns>返回哈希集的字符串形式</returns>
    public static string Round(this HashSet<string> strs)
    {
        StringBuilder result = new StringBuilder();
        foreach (string str in strs)
        {
            result.Append('【');
            result.Append(str);
            result.Append('】');
        }
        return result.ToString();
    }


    /// <summary>
    /// 将字节数组转化为字符串
    /// </summary>
    /// <param name="content">要转换的字符数组</param>
    /// <returns>将字节数组转化为字符串</returns>
    public static string ToStringL(this byte[] content,Encoding? encoding = null)
    {
        if(encoding is null)
        {
            encoding = Encoding.UTF8;
        }
        return encoding.GetString(content);
    }

    /// <summary>
    /// 将字符串转化为字符数组
    /// </summary>
    /// <param name="content">要转换的字符串</param>
    /// <returns>将字符串转化为字符数组</returns>
    public static byte[] ToByte(this string content, Encoding? encoding = null)
    {
        if (encoding is null)
        {
            encoding = Encoding.UTF8;
        }
        return encoding.GetBytes(content);
    }


        
    /// <summary>
    /// 将字符串转化为整型数字
    /// </summary>
    /// <param name="str">字符串</param>
    /// <returns>返回转换出来的数字</returns>
    public static int StringToInt(this string str)
    {
        Match match = RegexNumber.Match(str);
        if (match.Success)
        {
            decimal result = Convert.ToDecimal(match.Groups[0].Value);
            return Convert.ToInt32(result);
        }
        else
        {
            return 0;
        }
    }

    /// <summary>
    /// 将字符串转化为双精度浮点数
    /// </summary>
    /// <param name="str">字符串</param>
    /// <returns>返回转换出来的数字</returns>
    public static double StringToDouble(this string str)
    {
        Match match = RegexNumber.Match(str);
        if (match.Success)
        {
            decimal result = Convert.ToDecimal(match.Groups[0].Value);
            return Convert.ToDouble(result);
        }
        else
        {
            return 0.0;
        }
    }

    /// <summary>
    /// 将字符串转化为十进制数字
    /// </summary>
    /// <param name="str">字符串</param>
    /// <returns>返回转换出来的数字</returns>
    public static decimal StringToDecimal(this string str)
    {
        Match match = RegexNumber.Match(str);
        if (match.Success)
        {
            return Convert.ToDecimal(match.Groups[0].Value);
        }
        else
        {
            return 0m;
        }
    }


    #endregion



    #region 计时

    /// <summary>
    /// 停止计时器，并获取计时器所计到的时间的字符串描述
    /// </summary>
    public static string ClockString(this Stopwatch stopwatch)
    {
        stopwatch.Stop();
        double ticks = (double)stopwatch.ElapsedTicks / 10;    //微秒级别
        string unit = @"微秒";


        if (ticks < 30000)
        {
            return $"({ticks.ToString("#,##0.00")}{unit})";
        }
        else
        {
            ticks /= 1000;
            unit = @"毫秒";
        }

        if (ticks < 30000)
        {
            return $"({ticks.ToString("#,##0.00")}{unit})";
        }
        else
        {
            ticks /= 1000;
            unit = @"秒";
        }

        if (ticks < 300)
        {
            return $"({ticks.ToString("#,##0.00")}{unit})";
        }
        else
        {
            ticks /= 60;
            unit = @"分钟";
        }

        return $"({ticks.ToString("#,##0.00")}{unit})";


    }






    public readonly static Action<string> LogConsole = (str) => Console.WriteLine(str);
#if DEBUG
    public readonly static Action<string>? LogTest = (str)=> Console.WriteLine(str);
#else
    public readonly static Action<string>? LogTest = (str) => Trace.WriteLine(str);
#endif





    #endregion








    #region 运算符扩展

    public static int  Find(this string[] values, string value)
    {
        for(int i =0;i < values.Length; i++)
        {
            if (values[i] == value)
            {
                return i;
            }
        }
        return -1;
    }







    #endregion





}
