using MySqlX.XDevAPI.Common;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto.Static;

/// <summary>
/// KalevaAalto个人的数学专用库
/// </summary>
public static partial class Main
{

    /// <summary>
    /// 随机数生成器
    /// </summary>
    private readonly static Random entityRandom = new Random();

    /// <summary>
    /// 获取一个随机整数
    /// </summary>
    /// <param name="min">随机数范围的最小值</param>
    /// <param name="max">随机数范围的最大值</param>
    /// <returns>返回一个在最小值为Min和最大值为Max之间均匀分布的随机整数</returns>
    public static int GetRandomInt(int min,int max)
    {
        if (max == min)
        {
            return min;
        }
        else if(max < min)
        {
            int tempInt = min;
            min = max;
            max = tempInt;
        }
        return entityRandom.Next(min, max + 1);
    }


    /// <summary>
    /// 获取一个随机字符串
    /// </summary>
    /// <param name="randomCharList">随机字符串要选取的字符列表</param>
    /// <param name="length">随机字符串的长度</param>
    /// <returns>返回一个长度为Length且字符是从RandomCharList中选取的字符串。</returns>
    public static string GetRandomString(string randomCharList,int length)
    {
        //如果长度小于或等于0，则返回空字符串
        if (length <= 0) return string.Empty;

        //获取随机字符串长度
        int randomCharListLength = randomCharList.Length;

        //创建一个可变字符串对象
        StringBuilder result = new StringBuilder();

        //从字符列表中选取字符，并生成字符串
        for (int i = 0; i < length; i++)
        {
            result.Append(randomCharList[entityRandom.Next(randomCharListLength)]);
        }

        //返回随机字符串
        return result.ToString();
    }


    /// <summary>
    /// 获取一个随机字符串
    /// </summary>
    /// <param name="length">随机字符串的长度</param>
    /// <returns>返回一个长度为Length且字符是从大写字母、小写字母、数字中选取的字符串。</returns>
    public static string GetRandomString(int length)
    {
        return GetRandomString(@"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789",length);
    }


    /// <summary>
    /// 获取一个随机字符串
    /// </summary>
    /// <param name="minLength">随机字符串的最小长度</param>
    ///  <param name="maxLength">随机字符串的最大长度</param>
    /// <returns>返回一个长度为Length且字符是从大写字母、小写字母、数字中选取的字符串。</returns>
    public static string GetRandomString(int minLength,int maxLength)
    {
        return GetRandomString(@"ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789", GetRandomInt(minLength, maxLength));
    }


    public static string ToPercentage(this decimal Num)
    {
        return (Num*100).ToString(@"#,##.00") + '%';
    }

    public static T Around<T>(this T value,T value1,T value2) where T : struct,IComparable<T>
    {
        if (value.CompareTo(value1)<0)
        {
            return value1;
        }
        else if(value.CompareTo(value2) > 0)
        {
            return value2;
        }
        else
        {
            return value;
        }
    }


    private readonly static string[] cnNumbers = { @"零", @"壹", @"贰", @"叁", @"肆", @"伍", @"陆", @"柒", @"捌", @"玖" };
    private readonly static string[] cnUnits = { @"", @"拾", @"佰", @"仟" };
    private readonly static string[] cnGroupUnits = { @"", @"万", @"亿", @"万亿" };

    #region 金额大写
    public static string ConvertToChineseMoney(decimal amount)
    {
        

        StringBuilder result = new StringBuilder();
        if (amount < 0)
        {
            result.Append('负');
            amount *= -1;
        }

        string moneyStr = amount.ToString(@"0.00"); // 转换为带两位小数的字符串

        string[] parts = moneyStr.Split('.');
        string integerPart = parts[0];
        string decimalPart = parts.Length > 1 ? parts[1] : @"00";

        string integerStr = ConvertIntegerToChinese(integerPart);
        string decimalStr = ConvertDecimalToChinese(decimalPart);

        result.Append(integerStr);
        result.Append('元');
        result.Append(decimalStr);


        return result.ToString();
    }


    public static string[] dpNum(string str)
    {
        List<string> result = new List<string>();
        StringBuilder sb = new StringBuilder();

        for(int i = 0;i < str.Length; i++)
        {
            if (i % 4 == 0 && sb.Length > 0)
            {
                result.Add(sb.ToString()); ;
                sb = new StringBuilder();
            }
            sb.Append(str[i]);
        }

        if (sb.Length > 0)
        {
            result.Add(sb.ToString()); ;
        }

        


        return result.ToArray();
    }

    private static string ConvertIntegerToChinese(string integerPart)
    {
        if (integerPart == @"0")
        {
            return @"零";
        }
        StringBuilder result = new StringBuilder();

        bool isZero = true;
        bool isAllZero = true;
        for(int i = 0; i < integerPart.Length; i++)
        {
            int groupNumber = (integerPart.Length - i - 1) / 4;
            int unitNumber = (integerPart.Length - i - 1) % 4;

            if (integerPart[i] == '0')
            {
                if (isZero && unitNumber>0)
                {
                    result.Append('零');
                    isZero = false;
                }
            }
            else
            {
                result.Append(cnNumbers[integerPart[i]-'0']);
                result.Append(cnUnits[unitNumber]);
                isZero = true;
                isAllZero = false;
            }





            if(unitNumber == 0 && !isAllZero)
            {
                if (result[result.Length - 1] == '零')
                {
                    result.Remove(result.Length -1,1);
                }

                result.Append(cnGroupUnits[groupNumber]);
                isZero = false;
            }



            
        }





        return result.ToString();
    }

    private static string ConvertDecimalToChinese(string decimalPart)
    {
        if (decimalPart == @"00")
        {
            return @"整";
        }
        StringBuilder result = new StringBuilder();
        char c;

        c = decimalPart[0];
        result.Append(cnNumbers[c-'0']);
        result.Append(@"角");

        c = decimalPart[1];
        if (c != '0')
        {
            result.Append(cnNumbers[c - '0']);
            result.Append(@"分");
        }
        

        return result.ToString();
    }
    #endregion



}
