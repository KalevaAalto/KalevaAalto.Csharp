using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Models.Excel
{
    public static class NumberFormatTemplate
    {

        public const string Normal = @"@";
        public const string Int = @"#,##0";
        public const string Decimal = @"#,##0.00";

        public static string GetDecimal(int count)
        {
            if (count <= 0)
            {
                return Int;
            }

            return @"#,##0." + new string('0', count);

        }

        public const string DateTime = @"yyyy-MM-dd HH:mm:ss";
        public const string ChineseDateTime = @"yyyy年MM月dd日HH时mm分ss秒";
        public const string Date = @"yyyy-MM-dd";
        public const string ChineseDate = @"yyyy年MM月dd日";
        public const string Time = @"HH:mm:ss";
        public const string ChineseTime = @"HH时mm分ss秒";


        public const string Percent = @"0.00%";

        public static string GetPercent(int count)
        {
            if (count <= 0)
            {
                return @"0%";
            }
            return @"0." + new string('0', count) + '%';

        }


    }
}
