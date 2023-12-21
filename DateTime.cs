using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto
{
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


        public static DateTime currentMonth
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

            Match match = regexMonthString.Match(monthString);
            if (match.Success)
            {
                return new DateTime(Convert.ToInt16(match.Groups[@"year"].Value), Convert.ToInt16(match.Groups[@"month"].Value), 1);
            }
            else
            {
                throw new Exception($"“{monthString}”不是合法的月份字符串；");
            }
        }



        private static string[] dateStringFormats = {
                @"yyyy-MM-dd HH:mm:ss",
                @"yyyy-M-d H:m:s",
                @"yyyy-MM-dd HH:mm",
                @"yyyy-M-d H:m",
                @"yyyy/MM/dd HH:mm:ss",
                @"yyyy/M/d H:m:s",
                @"yyyy/MM/dd HH:mm",
                @"yyyy/M/d H:m",
                @"yyyy年MM月dd日 HH时mm分ss秒",
                @"yyyy年M月d日 H时m分s秒",
                @"yyyy年MM月dd日 HH时mm分",
                @"yyyy年M月d日 H时m分",


                @"yyyy-MM-dd",
                @"yyyy-M-d",
                @"yyyy-MM",
                @"yyyy-M",
                @"yyyy/MM/dd",
                @"yyyy/M/d",
                @"yyyy/MM",
                @"yyyy/M",
                @"yyyy年MM月dd日",
                @"yyyy年M月d日",
                @"yyyy年MM月",
                @"yyyy年M月",

                @"HH:mm:ss",
                @"H:m:s",
                @"HH:mm",
                @"H:m",
                @"HH时mm分ss秒",
                @"H时m分s秒",
                @"HH时mm分",
                @"H时m分",

                @"yyyyMMddHHmmss",
                @"yyyyMMddHHmm",
                @"yyyyMMdd",
                @"yyyyMM",

            };



        public static DateTime? GetDateTime(this string dateString)
        {
            if (DateTime.TryParseExact(dateString, dateStringFormats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime result))
            {
                return result;
            }

            Match match = regexNumber.Match(dateString);
            if (match.Success)
            {
                return DateTime.MinValue.AddDays(System.Convert.ToDouble(match.Groups[@"decimal"].Value));
            }



            return null;
        }












    }
}
