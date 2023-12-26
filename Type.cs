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




    }
}
