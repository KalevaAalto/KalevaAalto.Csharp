using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto;

public static partial class Main
{
    /// <summary>
    /// 获取某个日期所在月份的第一天
    /// </summary>
    /// <param name="date">要解析的日期</param>
    /// <returns>返回日期date日期所在月份的第一天</returns>
    public static DateTime MonthFirstDay(this DateTime date)
    {
        return new DateTime(date.Year, date.Month, 1);
    }

    /// <summary>
    /// 获取某个日期所在月份的最后一天
    /// </summary>
    /// <param name="date">要解析的日期</param>
    /// <returns>返回日期date日期所在月份的最后一天</returns>
    public static DateTime MonthLastDay(this DateTime date)
    {
        return date.AddMonths(1).MonthFirstDay().AddDays(-1);
    }

    /// <summary>
    /// 获取某个日期所在年份的第一天
    /// </summary>
    /// <param name="date">要解析的日期</param>
    /// <returns>返回日期date日期所在年份的第一天</returns>
    public static DateTime YearFirstDay(this DateTime date)
    {
        return new DateTime(date.Year, 1, 1);
    }

    /// <summary>
    /// 获取某个日期所在年份的最后一天
    /// </summary>
    /// <param name="date">要解析的日期</param>
    /// <returns>返回日期date日期所在年份的最后一天</returns>
    public static DateTime YearLastDay(this DateTime date)
    {
        return new DateTime(date.Year, 12, 31);
    }

    /// <summary>
    /// 获取时间和日期的全数字字符串
    /// </summary>
    /// <param name="date"></param>
    /// <returns></returns>
    public static string ToNumberString(this DateTime datetime)
    {
        return datetime.ToString(@"yyyyMMddHHmmss");
    }


    public static string MonthString(this DateTime date)
    {
        return date.ToString(@"yyyyMM");
    }



    public static string NowNumberString
    {
        get
        {
            return DateTime.Now.ToNumberString();
        }
    }


    public static DateTime CurrentMonth
    {
        get
        {
            return DateTime.Today.MonthFirstDay();
        }
    }


    public static string ToStandardDateString(this DateTime date)
    {
        return date.ToString(@"yyyy-MM-dd");
    }

    public static string ToStandardChineseDateString(this DateTime date)
    {
        return date.ToString(@"yyyy年MM月dd日");
    }

    public static DateTime GetMonthFromString(string monthString)
    {

        Match match = RegexMonthString.Match(monthString);
        if (match.Success)
        {
            return new DateTime(Convert.ToInt16(match.Groups[@"year"].Value), Convert.ToInt16(match.Groups[@"month"].Value), 1);
        }
        else
        {
            throw new Exception($"“{monthString}”不是合法的月份字符串；");
        }
    }



}
