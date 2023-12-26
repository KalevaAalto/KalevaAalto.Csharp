using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto
{
    public static partial class Main
    {
        /// <summary>
        /// 检查该类型是否为可空类型
        /// </summary>
        /// <param name="type">要检查的类型</param>
        public static bool IsNullable(this Type type)
        {
            return type.IsClass || type.IsArray || (type.IsValueType && type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>));
        }


        public static object? GetValue(this Type type,object? obj)
        {
            return new ObjectConvert(type).GetValue(obj);
        }


        



        private readonly static HashSet<Type> standardDataTableType = new HashSet<Type> { 

            //无符号整数
            typeof(byte),typeof(ushort), typeof(uint), typeof(ulong) ,
            typeof(byte?),typeof(ushort?), typeof(uint?), typeof(ulong?) ,

            //整数
            typeof(short), typeof(int), typeof(long) ,
            typeof(short?), typeof(int?), typeof(long?) ,

            //浮点数
            typeof(float) , typeof(double) ,
            typeof(float?) , typeof(double?) ,

            //十字制数
            typeof(decimal),typeof(decimal?),

            //布尔
            typeof(bool),typeof(bool?),
            
            //字符
            typeof(char),typeof(char?),

            //时间
            typeof(DateTime),typeof(DateTime?),

            //字符串
            typeof(string),
            
            //字节数组
            typeof(byte[]),
        };
        public static bool IsStandardDataTableType(this Type type)
        {
            return standardDataTableType.Contains(type);
        }












        /// <summary>
        /// 定义所有的整数类型的集合
        /// </summary>
        private readonly static HashSet<Type> integerType = new HashSet<Type> { typeof(short),typeof(int), typeof(long) };
        /// <summary>
        /// 定义所有的整数类型的集合
        /// </summary>
        private readonly static HashSet<Type> integerNullableType = new HashSet<Type> { typeof(short?), typeof(int?), typeof(long?) };
        /// <summary>
        /// 检查该类型是否为整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsInteger(this Type type)
        {
            return integerType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableInteger(this Type type)
        {
            return integerNullableType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableInteger(this Type type)
        {
            return integerType.Contains(type) || integerNullableType.Contains(type);
        }











        /// <summary>
        /// 定义所有的整数类型的集合
        /// </summary>
        private readonly static HashSet<Type> uintegerType = new HashSet<Type> { typeof(byte), typeof(ushort), typeof(uint), typeof(ulong) };
        /// <summary>
        /// 定义所有的整数类型的集合
        /// </summary>
        private readonly static HashSet<Type> uintegerNullableType = new HashSet<Type> { typeof(byte?), typeof(ushort?), typeof(uint?), typeof(ulong?) };
        /// <summary>
        /// 检查该类型是否为非负整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsUInteger(this Type type)
        {
            return uintegerType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为非负整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableUInteger(this Type type)
        {
            return uintegerNullableType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为非负整数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableUInteger(this Type type)
        {
            return uintegerType.Contains(type)|| uintegerNullableType.Contains(type);
        }







        /// <summary>
        /// 定义所有的小数类型的集合
        /// </summary>
        private readonly static HashSet<Type> floatType = new HashSet<Type> { typeof(float), typeof(double) };
        /// <summary>
        /// 定义所有的小数类型的集合
        /// </summary>
        private readonly static HashSet<Type> floatNullableType = new HashSet<Type> { typeof(float?), typeof(double?) };
        /// <summary>
        /// 检查该类型是否为浮点类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsFloat(this Type type)
        {
            return floatType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为浮点类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableFloat(this Type type)
        {
            return floatNullableType.Contains(type);
        }
        /// <summary>
        /// 检查该类型是否为浮点类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableFloat(this Type type)
        {
            return floatType.Contains(type)|| floatNullableType.Contains(type);
        }









        /// <summary>
        /// 检查该类型是否为小数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsDecimal(this Type type)
        {
            return type == typeof(decimal);
        }
        /// <summary>
        /// 检查该类型是否为小数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableDecimal(this Type type)
        {
            return type == typeof(decimal?);
        }
        /// <summary>
        /// 检查该类型是否为小数类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableDecimal(this Type type)
        {
            return type == typeof(decimal)|| type == typeof(decimal?);
        }






        


        /// <summary>
        /// 检查该类型是否为布尔类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsBool(this Type type)
        {
            return type == typeof(bool);
        }
        /// <summary>
        /// 检查该类型是否为布尔类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableBool(this Type type)
        {
            return type == typeof(bool?);
        }
        /// <summary>
        /// 检查该类型是否为布尔类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableBool(this Type type)
        {
            return type == typeof(bool) || type == typeof(bool?);
        }


        /// <summary>
        /// 检查该类型是否为字符类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsChar(this Type type)
        {
            return type == typeof(char);
        }
        /// <summary>
        /// 检查该类型是否为字符类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableChar(this Type type)
        {
            return type == typeof(char?);
        }
        /// <summary>
        /// 检查该类型是否为字符类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableChar(this Type type)
        {
            return type == typeof(char?);
        }





        /// <summary>
        /// 检查该类型是否为时间类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsDateTime(this Type type)
        {
            return type == typeof(DateTime);
        }
        /// <summary>
        /// 检查该类型是否为时间类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsNullableDateTime(this Type type)
        {
            return type == typeof(DateTime?);
        }
        /// <summary>
        /// 检查该类型是否为时间类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsOrNullableDateTime(this Type type)
        {
            return type == typeof(DateTime) || type == typeof(DateTime?);
        }





        /// <summary>
        /// 检查该类型是否为字符串类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsString(this Type type)
        {
            return type == typeof(string);
        }



        /// <summary>
        /// 检查该类型是否为字节数组类型，是的话返回true，不是的话就返回false
        /// </summary>
        public static bool IsByteArray(this Type type)
        {
            return type == typeof(byte[]);
        }



        public static bool IsNumber(this Type type)
        {
            return type.IsInteger()
                || type.IsUInteger()
                || type.IsFloat()
                || type.IsDecimal();
        }


        public static bool IsNullableNumber(this Type type)
        {
            return type.IsNullableInteger()
                || type.IsNullableUInteger()
                || type.IsNullableFloat()
                || type.IsNullableDecimal();
        }


        public static bool IsOrNullableNumber(this Type type)
        {
            return type.IsInteger()
                || type.IsUInteger()
                || type.IsFloat()
                || type.IsDecimal()
                || type.IsNullableInteger()
                || type.IsNullableUInteger()
                || type.IsNullableFloat()
                || type.IsNullableDecimal();
        }












        public readonly static Regex regexNumber = new Regex(@"(?<decimal>(?<int>[\-]?(?<uint>\d+))(\.\d+)?)");
        public readonly static Regex regexTrue = new Regex(@"[Tt]([Rr][Uu][Ee])?");
        public readonly static Regex regexFalse = new Regex(@"[Ff]([Aa][Ll][Ss][Ee])?");


        public readonly static Dictionary<Type, object?> numberZero = new Dictionary<Type, object?> 
        {
            { typeof(short) , (short)0 },
            { typeof(short?) , null },
            { typeof(int) , 0 },
            { typeof(int?) , null },
            { typeof(long) , 0L },
            { typeof(long?) , null },

            { typeof(byte) , (byte)0 },
            { typeof(byte?) , null },
            { typeof(ushort) , (ushort)0 },
            { typeof(ushort?) , null },
            { typeof(uint) , 0U },
            { typeof(uint?) , null },
            { typeof(ulong) , 0UL },
            { typeof(ulong?) , null },

            { typeof(float) , 0.0F },
            { typeof(float?) , null },
            { typeof(double) , 0.0D },
            { typeof(double?) , null },

            { typeof(decimal) , 0.0M },
            { typeof(decimal?) , null },

            { typeof(bool) , false },
            { typeof(bool?) , false },

            { typeof(char), (char)0 },
            { typeof(char?), null },

            { typeof(DateTime), DateTime.MinValue },
            { typeof(DateTime?), null },
        };


        public static object? TypeParse(this Type type,object? objSource)
        {
            if(objSource is null)
            {
                return numberZero.ContainsKey(type) ? numberZero[type] : null;
            }

            Type objSourceType = objSource.GetType();
            if(objSourceType == type)
            {
                return objSource;
            }

            if (type.IsOrNullableNumber())
            {
                if (!objSourceType.IsOrNullableNumber())
                {
                    string? objSourceString = objSource.ToString();
                    if (string.IsNullOrEmpty(objSourceString))
                    {
                        return numberZero[type];
                    }
                    Match match = regexNumber.Match(objSourceString);
                    if (!match.Success)
                    {
                        return numberZero[type];
                    }

                    if (type.IsOrNullableUInteger())
                    {
                        objSource = match.Groups[@"uint"].Value;
                    }
                    else if (type.IsOrNullableInteger())
                    {
                        objSource = match.Groups[@"int"].Value;
                    }
                    else
                    {
                        objSource = match.Groups[@"decimal"].Value;
                    }


                }

                if(type == typeof(byte) || type == typeof(byte?))
                {
                    return System.Convert.ToByte(objSource);
                }
                else if (type == typeof(ushort) || type == typeof(ushort?))
                {
                    return System.Convert.ToUInt16(objSource);
                }
                else if (type == typeof(uint) || type == typeof(uint?))
                {
                    return System.Convert.ToUInt32(objSource);
                }
                else if (type == typeof(ulong) || type == typeof(ulong?))
                {
                    return System.Convert.ToUInt64(objSource);
                }
                else if (type == typeof(short) || type == typeof(short?))
                {
                    return System.Convert.ToInt16(objSource);
                }
                else if (type == typeof(int) || type == typeof(int?))
                {
                    return System.Convert.ToInt32(objSource);
                }
                else if (type == typeof(long) || type == typeof(long?))
                {
                    return System.Convert.ToInt64(objSource);
                }
                else if (type == typeof(float) || type == typeof(float?))
                {
                    return System.Convert.ToSingle(objSource);
                }
                else if (type == typeof(double) || type == typeof(double?))
                {
                    return System.Convert.ToDouble(objSource);
                }
                else if (type == typeof(decimal) || type == typeof(decimal?))
                {
                    return System.Convert.ToDecimal(objSource);
                }
            }

            if (type.IsOrNullableBool())
            {
                if (objSourceType.IsOrNullableInteger())
                {
                    return !(System.Convert.ToInt64(objSource) == 0L);
                }
                else if (objSourceType.IsOrNullableUInteger())
                {
                    return !(System.Convert.ToUInt64(objSource) == 0UL);
                }
                else if (objSourceType.IsOrNullableFloat())
                {
                    return !(System.Convert.ToDouble(objSource) == 0.0D);
                }
                else if (objSourceType.IsOrNullableDecimal())
                {
                    return !(System.Convert.ToDecimal(objSource) == 0.0M);
                }
                else if (objSourceType.IsOrNullableChar())
                {
                    return !(System.Convert.ToChar(objSource) == (char)0);
                }
                else if (objSourceType.IsOrNullableChar())
                {
                    return !(System.Convert.ToChar(objSource) == (char)0);
                }
                else if (objSourceType.IsString())
                {
                    string? objSourceString = objSource.ToString();
                    if(!string.IsNullOrEmpty(objSourceString) && regexTrue.IsMatch(objSourceString))
                    {
                        return true;
                    }
                    else if (type.IsNullableBool())
                    {
                        if(!string.IsNullOrEmpty(objSourceString) && regexFalse.IsMatch(objSourceString))
                        {
                            return false;
                        }
                        else
                        {
                            return null;
                        }
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    if (type.IsNullableBool())
                    {
                        return null;
                    }
                    else
                    {
                        return false;
                    }
                }

            }

            if (type.IsOrNullableChar())
            {
                string? objSourceString = objSource.ToString();
                if (string.IsNullOrEmpty(objSourceString))
                {
                    if (type.IsChar())
                    {
                        return (char)0;
                    }
                    else
                    {
                        return null;
                    }
                }
                else
                {
                    return objSourceString[0];
                }
            }

            if (type.IsOrNullableDateTime())
            {
                if (objSourceType.IsOrNullableNumber())
                {
                    double addValue = System.Convert.ToDouble(objSource);
                    if (addValue < 0)
                    {
                        return numberZero[type];
                    }
                    else
                    {
                        return DateTime.MinValue.AddDays(addValue);
                    }
                }
                else if (objSourceType.IsOrNullableDateTime())
                {
                    return System.Convert.ToDateTime(objSource);
                }
                else
                {
                    string? objSourceString = objSource.ToString();

                    if (string.IsNullOrEmpty(objSourceString))
                    {
                        return numberZero[type];
                    }
                    else
                    {
                        DateTime? dateTime = objSourceString.GetDateTime();
                        if(dateTime is null)
                        {
                            return numberZero[type];
                        }
                        else
                        {
                            return dateTime;
                        }
                    }
                }
                
            }

            if (type.IsString())
            {
                return objSource.ToString();
            }



            return null;
        }




    }
}
