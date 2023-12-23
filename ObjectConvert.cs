

using Google.Protobuf;
using K4os.Compression.LZ4.Internal;
using System.Collections;
using System.Drawing.Text;
using System.Runtime.CompilerServices;
using System.Runtime.Serialization.Formatters.Binary;
using System.Text;
using System.Text.RegularExpressions;

namespace KalevaAalto
{
    public class ObjectConvert
    {
        private Type type;

        public ObjectConvert(Type type)
        {
            this.type = type;
        }

        private bool isNullable
        {
            get
            {
                return this.type.IsNullable();
            }
        }

        private object? defaultValue
        {
            get
            {
                if (this.isNullable)
                {
                    return null;
                }
                else
                {
                    return type.IsValueType ? Activator.CreateInstance(type) : null;
                }
            }
        }

        public object? GetValue(object? obj)
        {
            if(obj is null)
            {
                return this.defaultValue;
            }
            else if(this.type == obj.GetType())
            {
                return obj;
            }
            else if (this.type.IsOrNullableUInteger())
            {
                return this.ToUInteger(obj);
            }
            else if (this.type.IsOrNullableInteger())
            {
                return this.ToInteger(obj);
            }
            else if (this.type.IsOrNullableFloat())
            {
                return this.ToFloat(obj);
            }
            else if (this.type.IsOrNullableDecimal())
            {
                return this.ToDecimal(obj);
            }
            else if (this.type.IsOrNullableBool())
            {
                return this.ToBool(obj);
            }
            else if (this.type.IsOrNullableChar())
            {
                return this.ToChar(obj);
            }
            else if (this.type.IsOrNullableDateTime())
            {
                return this.ToDateTime(obj);
            }
            else if(this.type == typeof(string))
            {
                return this.ToString(obj);
            }
            else if(this.type == typeof(byte[]))
            {
                return this.ToByteArray(obj);
            }


            return this.defaultValue;
        }




        


        private readonly static Regex regexNumber = new Regex(@"[+-]?\d+(\.\d+)?");
        private object? ToUIntegerLastReturn(object tem)
        {
            if (this.type == typeof(byte) || this.type == typeof(byte?))
            {
                return System.Convert.ToByte(tem);
            }
            else if (this.type == typeof(ushort) || this.type == typeof(ushort?))
            {
                return System.Convert.ToUInt16(tem);
            }
            else if (this.type == typeof(uint) || this.type == typeof(uint?))
            {
                return System.Convert.ToUInt32(tem);
            }
            else if (this.type == typeof(ulong) || this.type == typeof(ulong?))
            {
                return System.Convert.ToUInt64(tem);
            }
            return this.defaultValue;
        }
        private object? ToUInteger(object obj)
        {
            Type objType = obj.GetType();
            ulong minValue = ulong.MinValue;
            ulong maxValue = ulong.MaxValue;
            if(this.type == typeof(byte) || this.type == typeof(byte?))
            {
                minValue = byte.MinValue;
                maxValue = byte.MaxValue;
            }
            else if (this.type == typeof(ushort) || this.type == typeof(ushort?))
            {
                minValue = ushort.MinValue;
                maxValue = ushort.MaxValue;
            }
            else if (this.type == typeof(uint) || this.type == typeof(uint?))
            {
                minValue = uint.MinValue;
                maxValue = uint.MaxValue;
            }
            else if (this.type == typeof(ulong) || this.type == typeof(ulong?))
            {
                minValue = ulong.MinValue;
                maxValue = ulong.MaxValue;
            }


            if (objType.IsOrNullableUInteger())
            {
                ulong tem = System.Convert.ToUInt64(obj);
                if(tem < minValue)
                {
                    return isNullable? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToUIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableInteger())
            {
                if (System.Convert.ToInt64(obj) < 0)
                {
                    return this.defaultValue;
                }
                else
                {
                    ulong tem = System.Convert.ToUInt64(obj);
                    if (tem < minValue)
                    {
                        return isNullable ? null : minValue;
                    }
                    else if (tem > maxValue)
                    {
                        return isNullable ? null : maxValue;
                    }
                    else
                    {
                        return this.ToUIntegerLastReturn(tem);
                    }
                }
            }
            else if (objType.IsOrNullableFloat())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToUIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDecimal())
            {
                decimal tem = (decimal)obj;
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToUIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableBool())
            {
                bool tem = (bool)obj;
                if (type == typeof(byte) || type == typeof(byte?))
                {
                    return tem ? (byte)1 : (byte)0;
                }
                else if (type == typeof(ushort) || type == typeof(ushort?))
                {
                    return tem ? (ushort)1 : (ushort)0;
                }
                else if (type == typeof(uint) || type == typeof(uint?))
                {
                    return tem ? (uint)1 : (uint)0;
                }
                else if (type == typeof(ulong) || type == typeof(ulong?))
                {
                    return tem ? (ulong)1 : (ulong)0;
                }
            }
            else if (objType.IsOrNullableChar())
            {
                ushort tem = System.Convert.ToUInt16(obj);
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToUIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDateTime())
            {
                double tem = ((DateTime)obj - DateTime.MinValue).TotalDays;
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToUIntegerLastReturn(tem);
                }
            }
            else if(objType == typeof(string))
            {
                if (string.IsNullOrEmpty((string)obj))
                {
                    return this.defaultValue;
                }

                Match match = regexNumber.Match((string)obj);
                if (match.Success)
                {
                    double tem = System.Convert.ToDouble(match.Value);
                    if (tem < minValue)
                    {
                        return isNullable ? null : minValue;
                    }
                    else if (tem > maxValue)
                    {
                        return isNullable ? null : maxValue;
                    }
                    else
                    {
                        return this.ToUIntegerLastReturn(tem);
                    }
                }
                else
                {
                    return this.defaultValue;
                }
            }
            else if(objType == typeof(byte[]))
            {
                //修改中

            }

            return this.defaultValue;

        }


        private object? ToIntegerLastReturn(object tem)
        {
            if (this.type == typeof(short) || this.type == typeof(short?))
            {
                return System.Convert.ToInt16(tem);
            }
            else if (this.type == typeof(int) || this.type == typeof(int?))
            {
                return System.Convert.ToInt32(tem);
            }
            else if (this.type == typeof(long) || this.type == typeof(long?))
            {
                return System.Convert.ToInt64(tem);
            }
            return this.defaultValue;
        }
        private object? ToInteger(object obj)
        {
            Type objType = obj.GetType();
            long minValue = long.MinValue;
            long maxValue = long.MaxValue;
            if (this.type == typeof(short) || this.type == typeof(short?))
            {
                minValue = short.MinValue;
                maxValue = short.MaxValue;
            }
            else if (this.type == typeof(int) || this.type == typeof(int?))
            {
                minValue = int.MinValue;
                maxValue = int.MaxValue;
            }
            else if (this.type == typeof(long) || this.type == typeof(long?))
            {
                minValue = long.MinValue;
                maxValue = long.MaxValue;
            }


            if (objType.IsOrNullableUInteger())
            {
                ulong tem = System.Convert.ToUInt64(obj);
                if (tem > (ulong)maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableInteger())
            {
                long tem = System.Convert.ToInt64(obj);
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableFloat())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDecimal())
            {
                decimal tem = (decimal)obj;
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableBool())
            {
                bool tem = (bool)obj;
                if (type == typeof(short) || type == typeof(short?))
                {
                    return tem ? (short)1 : (short)0;
                }
                else if (type == typeof(int) || type == typeof(int?))
                {
                    return tem ? (int)1 : (int)0;
                }
                else if (type == typeof(long) || type == typeof(long?))
                {
                    return tem ? (long)1 : (long)0;
                }
            }
            else if (objType.IsOrNullableChar())
            {
                ushort tem = System.Convert.ToUInt16(obj);
                if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDateTime())
            {
                double tem = (System.Convert.ToDateTime(obj) - DateTime.MinValue).TotalDays;
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToIntegerLastReturn(tem);
                }
            }
            else if (objType == typeof(string))
            {
                if (string.IsNullOrEmpty((string)obj))
                {
                    return this.defaultValue;
                }

                Match match = regexNumber.Match((string)obj);
                if (match.Success)
                {
                    double tem = System.Convert.ToDouble(match.Value);
                    if (tem < minValue)
                    {
                        return isNullable ? null : minValue;
                    }
                    else if (tem > maxValue)
                    {
                        return isNullable ? null : maxValue;
                    }
                    else
                    {
                        return this.ToIntegerLastReturn(tem);
                    }
                }
                else
                {
                    return this.defaultValue;
                }
            }
            else if (objType == typeof(byte[]))
            {
                //修改中

            }

            return this.defaultValue;

        }



        private object? ToFloatLastReturn(object tem)
        {
            if (this.type == typeof(float) || this.type == typeof(float?))
            {
                return System.Convert.ToSingle(tem);
            }
            else if (this.type == typeof(double) || this.type == typeof(double?))
            {
                return System.Convert.ToDouble(tem);
            }
            return this.defaultValue;
        }
        private object? ToFloat(object obj)
        {
            Type objType = obj.GetType();
            double minValue = double.MinValue;
            double maxValue = double.MaxValue;
            if (this.type == typeof(float) || this.type == typeof(float?))
            {
                minValue = float.MinValue;
                maxValue = float.MaxValue;
            }
            else if (this.type == typeof(double) || this.type == typeof(double?))
            {
                minValue = double.MinValue;
                maxValue = double.MaxValue;
            }


            if (objType.IsOrNullableUInteger())
            {
                double tem = System.Convert.ToUInt64(obj);
                return this.ToFloatLastReturn(tem);
            }
            else if (objType.IsOrNullableInteger())
            {
                double tem = System.Convert.ToUInt64(obj);
                return this.ToFloatLastReturn(tem);
            }
            else if (objType.IsOrNullableFloat())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < minValue)
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToFloatLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDecimal())
            {
                decimal tem = System.Convert.ToDecimal(obj);
                if (tem < System.Convert.ToDecimal(minValue))
                {
                    return isNullable ? null : minValue;
                }
                else if (tem > System.Convert.ToDecimal(maxValue))
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToFloatLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableBool())
            {
                bool tem = System.Convert.ToBoolean(obj);
                if (type == typeof(float) || type == typeof(float?))
                {
                    return tem ? (float)1 : (float)0;
                }
                else if (type == typeof(double) || type == typeof(double?))
                {
                    return tem ? (double)1 : (double)0;
                }
            }
            else if (objType.IsOrNullableChar())
            {
                ushort tem = System.Convert.ToUInt16(obj);
                if (tem > maxValue)
                {
                    return isNullable ? null : maxValue;
                }
                else
                {
                    return this.ToFloatLastReturn(tem);
                }
            }
            else if (objType.IsOrNullableDateTime())
            {
                double tem = ((DateTime)obj - DateTime.MinValue).TotalDays;
                return this.ToFloatLastReturn(tem);
            }
            else if (objType == typeof(string))
            {
                if (string.IsNullOrEmpty((string)obj))
                {
                    return this.defaultValue;
                }

                Match match = regexNumber.Match((string)obj);
                if (match.Success)
                {
                    double tem = System.Convert.ToDouble(match.Value);
                    if (tem < minValue)
                    {
                        return isNullable ? null : minValue;
                    }
                    else if (tem > maxValue)
                    {
                        return isNullable ? null : maxValue;
                    }
                    else
                    {
                        return this.ToFloatLastReturn(tem);
                    }
                }
                else
                {
                    return this.defaultValue;
                }
            }

            return this.defaultValue;

        }

        private object? ToDecimal(object obj)
        {
            Type objType = obj.GetType();


            if (objType.IsOrNullableUInteger())
            {
                return System.Convert.ToDecimal(obj);
            }
            else if (objType.IsOrNullableInteger())
            {
                return System.Convert.ToDecimal(obj);
            }
            else if (objType.IsOrNullableFloat())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < System.Convert.ToDouble(decimal.MinValue))
                {
                    return isNullable ? null : decimal.MinValue;
                }
                else if (tem > System.Convert.ToDouble(decimal.MaxValue))
                {
                    return isNullable ? null : decimal.MaxValue;
                }
                else
                {
                    return System.Convert.ToDecimal(tem);
                }
            }
            else if (objType.IsOrNullableDecimal())
            {
                return (decimal)obj;
            }
            else if (objType.IsOrNullableBool())
            {
                bool tem = System.Convert.ToBoolean(obj);
                return tem ? (decimal)1 : (decimal)0;
            }
            else if (objType.IsOrNullableChar())
            {
                return System.Convert.ToDecimal(obj);
            }
            else if (objType.IsOrNullableDateTime())
            {
                return System.Convert.ToDecimal(((DateTime)obj - DateTime.MinValue).TotalDays);
            }
            else if (objType == typeof(string))
            {
                if (string.IsNullOrEmpty((string)obj))
                {
                    return this.defaultValue;
                }

                Match match = regexNumber.Match((string)obj);
                if (match.Success)
                {
                    return System.Convert.ToDecimal(match.Value);
                }
                else
                {
                    return this.defaultValue;
                }
            }

            return this.defaultValue;

        }


        private object? ToBool(object? obj)
        {
            if (obj == null)
            {
                return this.defaultValue;
            }
            Type objType = obj.GetType();


            if (objType.IsOrNullableNumber())
            {
                double tem = System.Convert.ToDouble(obj);
                return tem == (double)0;
            }
            else if(objType.IsOrNullableBool())
            {
                return (bool)obj;
            }
            else if (objType.IsOrNullableChar())
            {
                if(new HashSet<char> { 't','T','对','是','1' }.Contains(System.Convert.ToChar(obj)))
                {
                    return true;
                }
                else if (new HashSet<char> { 'f', 'F', '错', '否', '0' }.Contains(System.Convert.ToChar(obj)))
                {
                    return false;
                }
                else
                {
                    return this.defaultValue;
                }
            }
            else if (objType == typeof(string))
            {
                if (new HashSet<string> { @"TRUE",@"True",@"T",@"true",@"t",@"是的",@"是",@"对的",@"对",@"1" }.Contains(System.Convert.ToString(obj) ?? string.Empty))
                {
                    return true;
                }
                else if (new HashSet<string> { @"FALSE", @"False", @"F", @"false", @"f",  @"否", @"错" ,@"错的", @"不对", @"0" }.Contains(System.Convert.ToString(obj) ?? string.Empty))
                {
                    return false;
                }
                else
                {
                    return this.defaultValue;
                }
            }
            return this.defaultValue;
        }

        private object? ToChar(object? obj)
        {
            if (obj == null)
            {
                return this.defaultValue;
            }
            Type objType = obj.GetType();
            char minValue = char.MinValue;
            char maxValue = char.MaxValue;

            if (objType.IsOrNullableUInteger())
            {
                ulong tem = System.Convert.ToUInt64(obj);
                if (tem < minValue)
                {
                    return this.defaultValue;
                }
                else if (tem > maxValue)
                {
                    return this.defaultValue;
                }
                else
                {
                    return System.Convert.ToChar(tem);
                }
            }
            else if (objType.IsOrNullableInteger())
            {
                long tem = System.Convert.ToInt64(obj);
                if (tem < minValue)
                {
                    return this.defaultValue;
                }
                else if (tem > maxValue)
                {
                    return this.defaultValue;
                }
                else
                {
                    return System.Convert.ToChar(tem);
                }
            }
            else if (objType.IsOrNullableFloat())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < minValue)
                {
                    return this.defaultValue;
                }
                else if (tem > maxValue)
                {
                    return this.defaultValue;
                }
                else
                {
                    return System.Convert.ToChar(tem);
                }
            }
            else if (objType.IsOrNullableDecimal())
            {
                decimal tem = (decimal)obj;
                if (tem < minValue)
                {
                    return this.defaultValue;
                }
                else if (tem > maxValue)
                {
                    return this.defaultValue;
                }
                else
                {
                    return System.Convert.ToChar(tem);
                }
            }
            else if (objType.IsOrNullableChar())
            {
                return (char)obj;
            }
            else if (objType == typeof(string))
            {
                if (string.IsNullOrEmpty((string)obj))
                {
                    return this.defaultValue;
                }
                return ((string)obj)[0];
            }
            else if(objType == typeof(byte[]))
            {
                byte[] tem = (byte[])obj;
                if (tem.Length == 0)
                {
                    return this.defaultValue;
                }
                else if(tem.Length == 1)
                {
                    return (char)tem[0];
                }
                else
                {
                    return ((char)tem[1] << 8) + tem[0];
                }
            }





            return this.defaultValue;
        }

        private object? ToDateTime(object? obj)
        {
            if (obj == null)
            {
                return this.defaultValue;
            }
            Type objType = obj.GetType();

            if (objType.IsOrNullableNumber())
            {
                double tem = System.Convert.ToDouble(obj);
                if (tem < 0)
                {
                    return this.isNullable? null : DateTime.MinValue;
                }
                else if(tem > (DateTime.MaxValue - DateTime.MinValue).TotalDays)
                {
                    return this.isNullable ? null : DateTime.MaxValue;
                }
                else
                {
                    return DateTime.MinValue.AddDays(tem);
                }
            }
            else if (objType.IsOrNullableDateTime())
            {
                return (DateTime)obj;
            }
            else if (objType == typeof(string))
            {
                string tem = (string)obj;
                if (string.IsNullOrEmpty(tem))
                {
                    return this.defaultValue;
                }

                if (DateTime.TryParse(tem, out DateTime dateTime))
                {
                    return dateTime;
                }
                else
                {
                    return this.defaultValue;
                }

            }

            return this.defaultValue;
        }




        private object? ToString(object? obj)
        {
            if (obj == null)
            {
                return this.defaultValue;
            }
            Type objType = obj.GetType();

            if (objType.IsOrNullableNumber())
            {
                return System.Convert.ToString(obj);
            }
            else if (objType.IsOrNullableBool())
            {
                return System.Convert.ToString(obj);
            }
            else if (objType.IsOrNullableChar())
            {
                return System.Convert.ToString(obj);
            }
            else if (objType.IsOrNullableDateTime())
            {
                return System.Convert.ToString(obj);
            }
            else if (objType == typeof(string))
            {
                return System.Convert.ToString(obj);
            }
            else if(objType == typeof(byte[]))
            {
                byte[] tem = (byte[])obj;
                if(tem.Length == 0)
                {
                    return this.defaultValue;
                }
                else
                {
                    StringBuilder stringBuilder = new StringBuilder(@"0x");
                    foreach (byte b in tem)
                    {
                        stringBuilder.AppendFormat(@"{0:X2}", b);
                    }
                    return stringBuilder.ToString();
                }
            }


            return this.defaultValue;
        }


        private object? ToByteArray(object? obj)
        {
            if (obj == null)
            {
                return this.defaultValue;
            }
            Type objType = obj.GetType();

            
            if(objType == typeof(byte) || objType == typeof(byte?))
            {
                return BitConverter.GetBytes((byte)obj);
            }
            else if (objType == typeof(ushort) || objType == typeof(ushort?))
            {
                return BitConverter.GetBytes((ushort)obj);
            }
            else if (objType == typeof(uint) || objType == typeof(uint?))
            {
                return BitConverter.GetBytes((uint)obj);
            }
            else if (objType == typeof(ulong) || objType == typeof(ulong?))
            {
                return BitConverter.GetBytes((ulong)obj);
            }
            else if (objType == typeof(short) || objType == typeof(short?))
            {
                return BitConverter.GetBytes((short)obj);
            }
            else if (objType == typeof(int) || objType == typeof(int?))
            {
                return BitConverter.GetBytes((int)obj);
            }
            else if (objType == typeof(long) || objType == typeof(long?))
            {
                return BitConverter.GetBytes((long)obj);
            }
            else if (objType == typeof(float) || objType == typeof(float?))
            {
                return BitConverter.GetBytes((float)obj);
            }
            else if (objType == typeof(double) || objType == typeof(double?))
            {
                return BitConverter.GetBytes((double)obj);
            }
            else if (objType == typeof(decimal) || objType == typeof(decimal?))
            {
                int[] bits = decimal.GetBits((decimal)obj);

                // 将每个整数部分和小数部分的字节数组连接在一起
                byte[] result = new byte[bits.Length * sizeof(int)];
                for (int i = 0; i < bits.Length; i++)
                {
                    BitConverter.GetBytes(bits[i]).CopyTo(result, i * sizeof(int));
                }
                return result;
            }
            else if (objType.IsOrNullableBool())
            {
                return BitConverter.GetBytes((bool)obj);
            }
            else if (objType.IsOrNullableChar())
            {
                return BitConverter.GetBytes((char)obj);
            }
            else if (objType == typeof(string))
            {
                return Encoding.Unicode.GetBytes(System.Convert.ToString(obj) ?? string.Empty);
            }
            else if (objType == typeof(byte[]))
            {
                return obj;
            }

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

}
