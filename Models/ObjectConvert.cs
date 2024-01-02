using MimeKit.Cryptography;
using NetTaste;
using System.Collections.Immutable;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;


namespace KalevaAalto.Models;
public class ObjectConvert
{
    private Type _type;
    public ObjectConvert(Type type)=>_type = type;


    private bool isNullable =>_type.IsNullable();

    private object? defaultValue
    {
        get
        {
            if (isNullable)return null;
            else return _type.IsValueType ? Activator.CreateInstance(_type) : null;
        }
    }

    



    public object? GetValue(object? obj)
    {
        if (obj is null) return defaultValue;
        else if (_type == obj.GetType()) return obj;
        else if (_type.IsOrNullableUInteger()) return ToUInteger(obj);
        else if (_type.IsOrNullableInteger()) return ToInteger(obj);
        else if (_type.IsOrNullableFloat()) return ToFloat(obj);
        else if (_type.IsOrNullableDecimal()) return ToDecimal(obj);
        else if (_type.IsOrNullableBool()) return ToBool(obj);
        else if (_type.IsOrNullableChar()) return ToChar(obj);
        else if (_type.IsOrNullableDateTime()) return ToDateTime(obj);
        else if (_type == typeof(string)) return ToString(obj);
        else if (_type == typeof(byte[])) return ToByteArray(obj);
        else return defaultValue;
    }







    private readonly static Regex regexNumber = new Regex(@"[+-]?\d+(\.\d+)?");
    private object? ToUIntegerLastReturn(object tem)
    {
        if (_type == typeof(byte) || _type == typeof(byte?)) return Convert.ToByte(tem);
        else if (_type == typeof(ushort) || _type == typeof(ushort?)) return Convert.ToUInt16(tem);
        else if (_type == typeof(uint) || _type == typeof(uint?)) return Convert.ToUInt32(tem);
        else if (_type == typeof(ulong) || _type == typeof(ulong?)) return Convert.ToUInt64(tem);
        else throw new NotSupportedException();
    }
    private object? ToUInteger(object obj)
    {
        Type objType = obj.GetType();
        ulong minValue = ulong.MinValue;
        ulong maxValue = ulong.MaxValue;
        if (_type == typeof(byte) || _type == typeof(byte?))
        {
            minValue = byte.MinValue;
            maxValue = byte.MaxValue;
        }
        else if (_type == typeof(ushort) || _type == typeof(ushort?))
        {
            minValue = ushort.MinValue;
            maxValue = ushort.MaxValue;
        }
        else if (_type == typeof(uint) || _type == typeof(uint?))
        {
            minValue = uint.MinValue;
            maxValue = uint.MaxValue;
        }
        else if (_type == typeof(ulong) || _type == typeof(ulong?))
        {
            minValue = ulong.MinValue;
            maxValue = ulong.MaxValue;
        }


        if (objType.IsOrNullableUInteger())
        {
            ulong tem = Convert.ToUInt64(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToUIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableInteger())
        {
            if (Convert.ToInt64(obj) < 0) return defaultValue;
            else
            {
                ulong tem = Convert.ToUInt64(obj);
                if (tem < minValue) return isNullable ? null : minValue;
                else if (tem > maxValue) return isNullable ? null : maxValue;
                else return ToUIntegerLastReturn(tem);
            }
        }
        else if (objType.IsOrNullableFloat())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToUIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableDecimal())
        {
            decimal tem = (decimal)obj;
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToUIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableBool())
        {
            bool tem = (bool)obj;
            if (_type == typeof(byte) || _type == typeof(byte?)) return tem ? (byte)1 : (byte)0;
            else if (_type == typeof(ushort) || _type == typeof(ushort?)) return tem ? (ushort)1 : (ushort)0;
            else if (_type == typeof(uint) || _type == typeof(uint?)) return tem ? 1 : (uint)0;
            else if (_type == typeof(ulong) || _type == typeof(ulong?)) return tem ? 1 : (ulong)0;
        }
        else if (objType.IsOrNullableChar())
        {
            ushort tem = Convert.ToUInt16(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToUIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableDateTime())
        {
            double tem = ((DateTime)obj - DateTime.MinValue).TotalDays;
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToUIntegerLastReturn(tem);
        }
        else if (objType == typeof(string))
        {
            if (string.IsNullOrEmpty((string)obj)) return defaultValue;

            Match match = regexNumber.Match((string)obj);
            if (match.Success)
            {
                double tem = Convert.ToDouble(match.Value);
                if (tem < minValue) return isNullable ? null : minValue;
                else if (tem > maxValue) return isNullable ? null : maxValue;
                else return ToUIntegerLastReturn(tem);
            }
            else return defaultValue;
        }
        else if (objType == typeof(byte[]))
        {
            //修改中

        }

        return defaultValue;

    }


    private object? ToIntegerLastReturn(object tem)
    {
        if (_type == typeof(sbyte) || _type == typeof(sbyte?)) return Convert.ToSByte(tem);
        else if (_type == typeof(short) || _type == typeof(short?)) return Convert.ToInt16(tem);
        else if (_type == typeof(int) || _type == typeof(int?)) return Convert.ToInt32(tem);
        else if (_type == typeof(long) || _type == typeof(long?)) return Convert.ToInt64(tem);
        else throw new NotSupportedException();
    }
    private object? ToInteger(object obj)
    {
        Type objType = obj.GetType();
        long minValue = long.MinValue;
        long maxValue = long.MaxValue;
        if (_type == typeof(sbyte) || _type == typeof(sbyte?))
        {
            minValue = sbyte.MinValue;
            maxValue = sbyte.MaxValue;
        }
        else if (_type == typeof(short) || _type == typeof(short?))
        {
            minValue = short.MinValue;
            maxValue = short.MaxValue;
        }
        else if (_type == typeof(int) || _type == typeof(int?))
        {
            minValue = int.MinValue;
            maxValue = int.MaxValue;
        }
        else if (_type == typeof(long) || _type == typeof(long?))
        {
            minValue = long.MinValue;
            maxValue = long.MaxValue;
        }


        if (objType.IsOrNullableUInteger())
        {
            ulong tem = Convert.ToUInt64(obj);
            if (tem > (ulong)maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableInteger())
        {
            long tem = Convert.ToInt64(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableFloat())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableDecimal())
        {
            decimal tem = (decimal)obj;
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableBool())
        {
            bool tem = (bool)obj;
            if (_type == typeof(sbyte) || _type == typeof(sbyte?)) return tem ? (sbyte)1 : (sbyte)0;
            else if (_type == typeof(short) || _type == typeof(short?)) return tem ? (short)1 : (short)0;
            else if (_type == typeof(int) || _type == typeof(int?)) return tem ? 1 : 0;
            else if (_type == typeof(long) || _type == typeof(long?)) return tem ? 1 : (long)0;
        }
        else if (objType.IsOrNullableChar())
        {
            ushort tem = Convert.ToUInt16(obj);
            if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType.IsOrNullableDateTime())
        {
            double tem = (Convert.ToDateTime(obj) - DateTime.MinValue).TotalDays;
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToIntegerLastReturn(tem);
        }
        else if (objType == typeof(string))
        {
            if (string.IsNullOrEmpty((string)obj)) return defaultValue;

            Match match = regexNumber.Match((string)obj);
            if (match.Success)
            {
                double tem = Convert.ToDouble(match.Value);
                if (tem < minValue) return isNullable ? null : minValue;
                else if (tem > maxValue) return isNullable ? null : maxValue;
                else return ToIntegerLastReturn(tem);
            }
            else return defaultValue;
        }
        else if (objType == typeof(byte[]))
        {
            throw new NotImplementedException();

        }
        

        return defaultValue;

    }



    private object? ToFloatLastReturn(object tem)
    {
        if (_type == typeof(float) || _type == typeof(float?)) return Convert.ToSingle(tem);
        else if (_type == typeof(double) || _type == typeof(double?)) return Convert.ToDouble(tem);
        else return new NotImplementedException();
    }
    private object? ToFloat(object obj)
    {
        Type objType = obj.GetType();
        double minValue = double.MinValue;
        double maxValue = double.MaxValue;
        if (_type == typeof(float) || _type == typeof(float?))
        {
            minValue = float.MinValue;
            maxValue = float.MaxValue;
        }
        else if (_type == typeof(double) || _type == typeof(double?))
        {
            minValue = double.MinValue;
            maxValue = double.MaxValue;
        }


        if (objType.IsOrNullableUInteger()) return ToFloatLastReturn(obj);
        else if (objType.IsOrNullableInteger()) return ToFloatLastReturn(obj);
        else if (objType.IsOrNullableFloat())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < minValue) return isNullable ? null : minValue;
            else if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToFloatLastReturn(tem);
        }
        else if (objType.IsOrNullableDecimal())
        {
            decimal tem = Convert.ToDecimal(obj);
            if (tem < Convert.ToDecimal(minValue)) return isNullable ? null : minValue;
            else if (tem > Convert.ToDecimal(maxValue)) return isNullable ? null : maxValue;
            else return ToFloatLastReturn(tem);
        }
        else if (objType.IsOrNullableBool())
        {
            bool tem = Convert.ToBoolean(obj);
            if (_type == typeof(float) || _type == typeof(float?)) return tem ? (float)1 : (float)0;
            else if (_type == typeof(double) || _type == typeof(double?)) return tem ? (double)1 : (double)0;
        }
        else if (objType.IsOrNullableChar())
        {
            ushort tem = Convert.ToUInt16(obj);
            if (tem > maxValue) return isNullable ? null : maxValue;
            else return ToFloatLastReturn(tem);
        }
        else if (objType.IsOrNullableDateTime())
        {
            double tem = ((DateTime)obj - DateTime.MinValue).TotalDays;
            return ToFloatLastReturn(tem);
        }
        else if (objType == typeof(string))
        {
            if (string.IsNullOrEmpty((string)obj)) return defaultValue;

            Match match = regexNumber.Match((string)obj);
            if (match.Success)
            {
                double tem = Convert.ToDouble(match.Value);
                if (tem < minValue) return isNullable ? null : minValue;
                else if (tem > maxValue) return isNullable ? null : maxValue;
                else return ToFloatLastReturn(tem);
            }
            else return defaultValue;
        }

        return defaultValue;

    }

    private object? ToDecimal(object obj)
    {
        Type objType = obj.GetType();


        if (objType.IsOrNullableUInteger()) return Convert.ToDecimal(obj);
        else if (objType.IsOrNullableInteger()) return Convert.ToDecimal(obj);
        else if (objType.IsOrNullableFloat())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < Convert.ToDouble(decimal.MinValue)) return isNullable ? null : decimal.MinValue;
            else if (tem > Convert.ToDouble(decimal.MaxValue)) return isNullable ? null : decimal.MaxValue;
            else return Convert.ToDecimal(tem);
        }
        else if (objType.IsOrNullableDecimal()) return (decimal)obj;
        else if (objType.IsOrNullableBool())
        {
            bool tem = Convert.ToBoolean(obj);
            return tem ? (decimal)1 : (decimal)0;
        }
        else if (objType.IsOrNullableChar()) return Convert.ToDecimal(obj);
        else if (objType.IsOrNullableDateTime()) return Convert.ToDecimal(((DateTime)obj - DateTime.MinValue).TotalDays);
        else if (objType == typeof(string))
        {
            if (string.IsNullOrEmpty((string)obj)) return defaultValue;

            Match match = regexNumber.Match((string)obj);
            if (match.Success) return Convert.ToDecimal(match.Value);
            else return defaultValue;
        }

        return defaultValue;

    }



    private readonly static ImmutableHashSet<char> s_trueChars = new HashSet<char> { 't', 'T', '对', '是', '1' }.ToImmutableHashSet();
    private readonly static ImmutableHashSet<char> s_falseChars = new HashSet<char> { 'f', 'F', '错', '否', '0' }.ToImmutableHashSet();
    private readonly static ImmutableHashSet<string> s_trueStrings = new HashSet<string> { @"TRUE", @"True", @"T", @"true", @"t", @"是的", @"是", @"对的", @"对", @"1" }.ToImmutableHashSet();
    private readonly static ImmutableHashSet<string> s_falseStrings = new HashSet<string> { @"FALSE", @"False", @"F", @"false", @"f", @"否", @"错", @"错的", @"不对", @"0" }.ToImmutableHashSet();
    private object? ToBool(object obj)
    {
        Type objType = obj.GetType();

        if (objType.IsOrNullableNumber()) return Convert.ToDouble(obj) == 0;
        else if (objType.IsOrNullableBool()) return (bool)obj;
        else if (objType.IsOrNullableChar())
        {
            if (s_trueChars.Contains(Convert.ToChar(obj))) return true;
            else if (s_falseChars.Contains(Convert.ToChar(obj))) return false;
            else return defaultValue;
        }
        else if (objType == typeof(string))
        {
            if (s_trueStrings.Contains(Convert.ToString(obj) ?? string.Empty)) return true;
            else if (s_falseStrings.Contains(Convert.ToString(obj) ?? string.Empty)) return false;
            else return defaultValue;
        }
        return defaultValue;
    }

    private object? ToChar(object obj)
    {
        Type objType = obj.GetType();
        char minValue = char.MinValue;
        char maxValue = char.MaxValue;

        if (objType.IsOrNullableUInteger())
        {
            ulong tem = Convert.ToUInt64(obj);
            if (tem < minValue)return defaultValue;
            else if (tem > maxValue)return defaultValue;
            else return Convert.ToChar(tem);
        }
        else if (objType.IsOrNullableInteger())
        {
            long tem = Convert.ToInt64(obj);
            if (tem < minValue) return defaultValue;
            else if (tem > maxValue) return defaultValue;
            else return Convert.ToChar(tem);
        }
        else if (objType.IsOrNullableFloat())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < minValue) return defaultValue;
            else if (tem > maxValue) return defaultValue;
            else return Convert.ToChar(tem);
        }
        else if (objType.IsOrNullableDecimal())
        {
            decimal tem = (decimal)obj;
            if (tem < minValue) return defaultValue;
            else if (tem > maxValue) return defaultValue;
            else return Convert.ToChar(tem);
        }
        else if (objType.IsOrNullableChar()) return (char)obj;
        else if (objType == typeof(string))
        {
            if (string.IsNullOrEmpty((string)obj)) return defaultValue;
            else return ((string)obj)[0];
        }
        else if (objType == typeof(byte[]))
        {
            byte[] tem = (byte[])obj;
            if (tem.Length == 0) return defaultValue;
            else if (tem.Length == 1) return (char)tem[0];
            else return ((char)tem[1] << 8) + tem[0];
        }

        return defaultValue;
    }

    private object? ToDateTime(object obj)
    {
        Type objType = obj.GetType();

        if (objType.IsOrNullableNumber())
        {
            double tem = Convert.ToDouble(obj);
            if (tem < 0)return isNullable ? null : DateTime.MinValue;
            else if (tem > (DateTime.MaxValue - DateTime.MinValue).TotalDays) return isNullable ? null : DateTime.MaxValue;
            else return DateTime.FromOADate(tem);
        }
        else if (objType.IsOrNullableDateTime()) return (DateTime)obj;
        else if (objType == typeof(string))
        {
            string tem = (string)obj;
            if (string.IsNullOrEmpty(tem)) return defaultValue;

            if (DateTime.TryParse(tem, out DateTime dateTime)) return dateTime;
            else return defaultValue;
        }

        return defaultValue;
    }




    private object? ToString(object obj)
    {
        Type objType = obj.GetType();

        if (objType.IsOrNullableNumber() || objType.IsOrNullableBool() || objType.IsOrNullableChar() || objType.IsOrNullableDateTime() || objType.IsString())
        {
            return Convert.ToString(obj);
        }
        else if (objType == typeof(byte[]))
        {
            byte[] tem = (byte[])obj;
            if (tem.Length == 0) return defaultValue;
            else
            {
                StringBuilder stringBuilder = new StringBuilder(@"0x");
                foreach (byte b in tem) stringBuilder.AppendFormat(@"{0:X2}", b);
                return stringBuilder.ToString();
            }
        }
        return defaultValue;
    }


    private object? ToByteArray(object obj)
    {
        Type objType = obj.GetType();


        if (objType == typeof(byte) || objType == typeof(byte?))return BitConverter.GetBytes((byte)obj);
        else if (objType == typeof(ushort) || objType == typeof(ushort?))return BitConverter.GetBytes((ushort)obj);
        else if (objType == typeof(uint) || objType == typeof(uint?))return BitConverter.GetBytes((uint)obj);
        else if (objType == typeof(ulong) || objType == typeof(ulong?))return BitConverter.GetBytes((ulong)obj);
        else if (objType == typeof(sbyte) || objType == typeof(sbyte?))return BitConverter.GetBytes((sbyte)obj);
        else if (objType == typeof(short) || objType == typeof(short?))return BitConverter.GetBytes((short)obj);
        else if (objType == typeof(int) || objType == typeof(int?)) return BitConverter.GetBytes((int)obj);
        else if (objType == typeof(long) || objType == typeof(long?)) return BitConverter.GetBytes((long)obj);
        else if (objType == typeof(float) || objType == typeof(float?)) return BitConverter.GetBytes((float)obj);
        else if (objType == typeof(double) || objType == typeof(double?)) return BitConverter.GetBytes((double)obj);
        else if (objType == typeof(decimal) || objType == typeof(decimal?))
        {
            int[] bits = decimal.GetBits((decimal)obj);

            // 将每个整数部分和小数部分的字节数组连接在一起
            byte[] result = new byte[bits.Length * sizeof(int)];
            for (int i = 0; i < bits.Length; i++) BitConverter.GetBytes(bits[i]).CopyTo(result, i * sizeof(int));
            return result;
        }
        else if (objType.IsOrNullableBool()) return BitConverter.GetBytes((bool)obj);
        else if (objType.IsOrNullableChar()) return BitConverter.GetBytes((char)obj);
        else if (objType == typeof(string)) return Encoding.Unicode.GetBytes(Convert.ToString(obj) ?? string.Empty);
        else if (objType == typeof(byte[])) return obj;

        using (MemoryStream stream = new MemoryStream())
        {
            BinaryFormatter formatter = new BinaryFormatter();
#pragma warning disable SYSLIB0011 // 类型或成员已过时
            formatter.Serialize(stream, obj);
#pragma warning restore SYSLIB0011 // 类型或成员已过时
            return stream.ToArray();
        }

    }

}