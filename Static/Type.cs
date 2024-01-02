using System;
using System.Collections.Generic;
using System.Collections.Immutable;
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


        public static object? GetValue(this Type type,object? obj) => new Models.ObjectConvert(type).GetValue(obj);
        public static string? GetString(object? obj) => (string?)typeof(string).GetValue(obj);
        public static int GetInt32(object? obj) => (int)typeof(int).GetValue(obj)!;
        public static decimal GetDecimal(object? obj) => (decimal)typeof(decimal).GetValue(obj)!;
        public static DateTime GetDateTime(object? obj) => (DateTime)typeof(DateTime).GetValue(obj)!;





        private readonly static ImmutableHashSet<Type> s_integerType = new HashSet<Type> { typeof(sbyte),typeof(short),typeof(int), typeof(long) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_integerNullableType = new HashSet<Type> { typeof(sbyte?), typeof(short?), typeof(int?), typeof(long?) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_integerOrNullableType = new HashSet<Type>(s_integerType.Intersect(s_integerNullableType)).ToImmutableHashSet();
        public static bool IsInteger(this Type type) => s_integerType.Contains(type);
        public static bool IsNullableInteger(this Type type) => s_integerNullableType.Contains(type);
        public static bool IsOrNullableInteger(this Type type) => s_integerOrNullableType.Contains(type);



        private readonly static ImmutableHashSet<Type> s_uintegerType = new HashSet<Type> { typeof(byte), typeof(ushort), typeof(uint), typeof(ulong) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_uintegerNullableType = new HashSet<Type> { typeof(byte?), typeof(ushort?), typeof(uint?), typeof(ulong?) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_uintegerOrNullableType = new HashSet<Type>(s_uintegerType.Intersect(s_uintegerNullableType)).ToImmutableHashSet();
        public static bool IsUInteger(this Type type) => s_uintegerType.Contains(type);
        public static bool IsNullableUInteger(this Type type) => s_uintegerNullableType.Contains(type);
        public static bool IsOrNullableUInteger(this Type type) => s_uintegerOrNullableType.Contains(type);





        private readonly static ImmutableHashSet<Type> s_floatType = new HashSet<Type> { typeof(float), typeof(double) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_floatNullableType = new HashSet<Type> { typeof(float?), typeof(double?) }.ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_floatOrNullableType = new HashSet<Type> (s_floatType.Intersect(s_floatNullableType)).ToImmutableHashSet();
        public static bool IsFloat(this Type type) => s_floatType.Contains(type);
        public static bool IsNullableFloat(this Type type) => s_floatNullableType.Contains(type);
        public static bool IsOrNullableFloat(this Type type) => s_floatOrNullableType.Contains(type);



        public static bool IsDecimal(this Type type) => type == typeof(decimal);
        public static bool IsNullableDecimal(this Type type) => type == typeof(decimal?);
        public static bool IsOrNullableDecimal(this Type type) => type == typeof(decimal) || type == typeof(decimal?);


        public static bool IsBool(this Type type) => type == typeof(bool);
        public static bool IsNullableBool(this Type type) => type == typeof(bool?);
        public static bool IsOrNullableBool(this Type type) => type == typeof(bool) || type == typeof(bool?);



        public static bool IsChar(this Type type) => type == typeof(char);
        public static bool IsNullableChar(this Type type) => type == typeof(char?);
        public static bool IsOrNullableChar(this Type type) => type == typeof(char) || type == typeof(char?);





        public static bool IsDateTime(this Type type) => type == typeof(DateTime);
        public static bool IsNullableDateTime(this Type type) => type == typeof(DateTime?);
        public static bool IsOrNullableDateTime(this Type type) => type == typeof(DateTime) || type == typeof(DateTime?);




        public static bool IsString(this Type type) => type == typeof(string);
        public static bool IsByteArray(this Type type) =>type == typeof(byte[]);



        private readonly static ImmutableHashSet<Type> s_numberType = new HashSet<Type>(s_integerType.Intersect(s_uintegerType).Intersect(s_floatType).Add(typeof(decimal))).ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_numberNullableType = new HashSet<Type>(s_integerNullableType.Intersect(s_uintegerNullableType).Intersect(s_floatNullableType).Add(typeof(decimal?))).ToImmutableHashSet();
        private readonly static ImmutableHashSet<Type> s_numberOrNullableType = new HashSet<Type>(s_numberType.Intersect(s_numberNullableType)).ToImmutableHashSet();
        public static bool IsNumber(this Type type) => s_numberType.Contains(type);
        public static bool IsNullableNumber(this Type type) => s_numberNullableType.Contains(type);
        public static bool IsOrNullableNumber(this Type type) => s_numberOrNullableType.Contains(type);



        private readonly static ImmutableHashSet<Type> s_standardDataTableType = new HashSet<Type>(
            s_numberOrNullableType.Intersect(new Type[]
            {
            typeof(bool),typeof(bool?),
            typeof(char),typeof(char?),
            typeof(DateTime),typeof(DateTime?),

            typeof(string),
            typeof(byte[]),
            })
            ).ToImmutableHashSet();
        public static bool IsStandardDataTableType(this Type type) => s_standardDataTableType.Contains(type);


    }
}
