using Microsoft.Extensions.FileSystemGlobbing;
using Microsoft.Win32;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using OfficeOpenXml;
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

namespace KalevaAalto
{

        
    /// <summary>
    /// KalevaAalto个人的常用库
    /// </summary>
    public static partial class Main
    {
        
        //未找到
        public const int notfound = -1;

        public readonly static Regex regexMonthString = new Regex(@"(?<year>\d{4})[年\-\/]?(?<month>\d{1,2})[月]?");

        /// <summary>
        /// 获取一个带时间的测试名称
        /// </summary>
        public static string testName
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
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }


        



        #region 文件流

        /// <summary>
        /// 定义默认字符集为Unicode
        /// </summary>
        public static System.Text.Encoding defaultEncoding = Encoding.Unicode;


        /// <summary>
        /// 解析文本文档的字符集
        /// </summary>
        /// <param name="stream">要分析字符集的数据流</param>
        /// <returns>返回路径为fileName的文本文档所使用的字符集</returns>
        public static System.Text.Encoding GetEncoding(this Stream stream)
        {
            
            if (stream.Length > int.MaxValue)
            {
                throw new Exception(@"文件过大，超过2GB，无法解析字符集！");
            }

            Func<byte[], bool> IsUtf8Bytes = (data) =>
            {
                int charByteCounter = 1; //计算当前正分析的字符应还有的字节数
                byte curByte; //当前分析的字节.
                for (int i = 0; i < data.Length; i++)
                {
                    curByte = data[i];
                    if (charByteCounter == 1)
                    {
                        if (curByte >= 0x80)
                        {
                            //判断当前
                            while (((curByte <<= 1) & 0x80) != 0) charByteCounter++;

                            //标记位首位若为非0 则至少以2个1开始 如:110XXXXX…1111110X
                            if (charByteCounter == 1 || charByteCounter > 6) return false;

                        }
                    }
                    else
                    {
                        //若是UTF-8 此时第一位必须为1
                        if ((curByte & 0xC0) != 0x80) return false;
                        charByteCounter--;
                    }
                }
                if (charByteCounter > 1) throw new Exception("非预期的byte格式");

                return true;
            };

            BinaryReader binaryReader = new BinaryReader(stream, Encoding.Default);
            byte[] ss = binaryReader.ReadBytes((int)stream.Length);
            binaryReader.Close();

            if (IsUtf8Bytes(ss) || (ss.Length >= 3 && ss[0] == 0xEF && ss[1] == 0xBB && ss[2] == 0xBF))
            {
                return Encoding.UTF8;
            }
            else if (ss.Length >= 4 && ss[0] == 0xFF && ss[1] == 0xFE && ss[0] == 0x00 && ss[1] == 0x00)
            {
                return Encoding.UTF32;
            }
            else if (ss.Length >= 4 && ss[0] == 0x00 && ss[1] == 0x00 && ss[0] == 0xFE && ss[1] == 0xFF)
            {
                return Encoding.GetEncoding(12001);
            }
            else if (ss.Length >= 2 && ss[0] == 0xFE && ss[1] == 0xFF)
            {
                return Encoding.BigEndianUnicode;
            }
            else if (ss.Length >= 2 && ss[0] == 0xFF && ss[1] == 0xFE)
            {
                return Encoding.Unicode;
            }

            

            return Encoding.GetEncoding(936);

        }


        /// <summary>
        /// 解析文本文档的字符集
        /// </summary>
        /// <param name="fileName">要分析字符集的文本文档的路径</param>
        /// <returns>返回路径为fileName的文本文档所使用的字符集</returns>
        public static Encoding GetEncoding(string fileName)
        {
            using (FileStream fileStream = new FileStream(fileName, FileMode.Open, FileAccess.Read))
            {
                return fileStream.GetEncoding();
            }
        }
       

        /// <summary>
        /// 从文本文档中获取字符串
        /// </summary>
        /// <param name="fileName">要获取字符串的文本文档的路径</param>
        /// <param name="encoding">要获取字符串的字符集，无则自动解析字符集</param>
        /// <returns>返回从路径fileName的文本文档所获取的字符串</returns>
        public static string GetStringFromFile(string fileName, Encoding? encoding = null)
        {
            if (encoding is null)
            {
                encoding = GetEncoding(fileName);
            }
            return File.ReadAllText(fileName, encoding);

        }

        /// <summary>
        /// 从文本文档中获取字符串
        /// </summary>
        /// <param name="fileNameInfo">要获取字符串的文件名信息对象</param>
        /// <param name="encoding">要获取字符串的字符集，无则自动解析字符集</param>
        /// <returns>返回从路径fileName的文本文档所获取的字符串</returns>
        public static string GetStringFromFile(FileNameInfo fileNameInfo, Encoding? encoding = null)
        {
            return GetStringFromFile(fileNameInfo.fileName,encoding);
        }

        /// <summary>
        /// 将字符串str以encoding为字符集，保存在路径为fileName的文件中
        /// </summary>
        /// <param name="str">要保存的字符串</param>
        /// <param name="fileName">要获取字符串的文本文档的路径</param>
        /// <param name="encoding">要获取字符串的字符集，无则选取默认字符串</param>
        public static void SaveToFile(this string str, string fileName, Encoding? encoding = null)
        {
            //如果字符集为空，则选则默认字符集
            if (encoding is null)
            {
                encoding = defaultEncoding;
            }

            //将字符串str以encoding为字符集，保存在路径为fileName的文件中
            File.WriteAllText(fileName, str, encoding);
        }

        /// <summary>
        /// 将字符串str以encoding为字符集，保存在路径为fileName的文件中
        /// </summary>
        /// <param name="str">要保存的字符串</param>
        /// <param name="fileNameInfo">要获取字符串的文本文档的文件名对象</param>
        /// <param name="encoding">要获取字符串的字符集，无则选取默认字符串</param>
        public static void SaveToFile(this string str, FileNameInfo fileNameInfo, Encoding? encoding = null)
        {
            str.SaveToFile(fileNameInfo.fileName,encoding);
        }

        /// <summary>
        /// 用来解析文件路径的的文件名类
        /// </summary>
        public class FileNameInfo
        {
            /// <summary>
            /// 文件所在文件夹的路径
            /// </summary>
            public string path { get; } = string.Empty;

            /// <summary>
            /// 文件名（不包含后缀名）
            /// </summary>
            public string name { get; set; } = string.Empty;

            /// <summary>
            /// 文件后缀名
            /// </summary>
            public string suffix { get; } = string.Empty;

            /// <summary>
            /// 文件名类的状态，异常则为false，此时的文件名对象无法正常使用
            /// </summary>
            public bool status { get; } = false;

            /// <summary>
            /// 返回文件是否存在
            /// </summary>
            public bool exists
            {
                get
                {
                    return File.Exists(this.fileName);
                }
            }


            public char sign { get; } = '\\';

            /// <summary>
            /// 创建解析文件路径所需要用到的正则表达式对象
            /// </summary>
            private readonly static Regex regex = new Regex(@"^((?<path>.+)[\\/])?(?<name>[^\\/]+?)(\.(?<suffix>[^\\\\./]+))?$");

            /// <summary>
            /// 以一个文件的文件路径来创建文件名对象
            /// </summary>
            /// <param name="fileName">文件的文件路径</param>
            public FileNameInfo(string? fileName)
            {
                if(fileName is null)
                {
                    return;
                }

                //拆分路径
                Match match = regex.Match(fileName);
                if (!match.Success)
                {
                    return;
                }

                if (fileName.Contains('/'))
                {
                    this.sign = '/';
                }


                this.path = match.Groups[@"path"].Value;
                this.name = match.Groups[@"name"].Value;
                this.suffix = match.Groups[@"suffix"].Value.ToLower();
                this.status = true;
            }

            /// <summary>
            /// 手动文件名对象
            /// </summary>
            /// <param name="path">文件所在文件夹的路径</param>
            /// <param name="name">文件名（不包含后缀名）</param>
            /// <param name="suffix">文件后缀名</param>
            public FileNameInfo(string path,string name,string? suffix = null)
            {
                this.path = path;
                this.name = name;
                this.suffix = suffix ?? string.Empty;
                this.status = true;
            }

            /// <summary>
            /// 返回此文件名对象的文件路径
            /// </summary>
            public string fileName
            {
                get
                {
                    if(this.path.Length > 0)
                    {
                        return this.path + this.sign + this.partialFileName;
                    }
                    else
                    {
                        return this.partialFileName;
                    }
                    
                }
            }

            /// <summary>
            /// 返回此文件名对象的文件名（有后缀名）
            /// </summary>
            public string partialFileName
            {
                get
                {
                    if(this.suffix.Length > 0)
                    {
                        return this.name + '.' + this.suffix;
                    }
                    else
                    {
                        return this.name;
                    }
                    
                }
            }


            public FileInfo fileInfo
            {
                get
                {
                    return new FileInfo(this.fileName);
                }
            }

        }


        /// <summary>
        /// 解析的文件路径数组，并返回文件名数组
        /// </summary>
        /// <param name="fileNames">要解析的文件路径数组</param>
        /// <returns>返回从文件路径数组fileNames中解析出的文件名数组</returns>
        public static FileNameInfo[] GetFileNameInfos(this string?[]? fileNames)
        {
            //如果文件路径数组fileNames为空值，则返回空文件名数组
            if (fileNames is null)
            {
                return new FileNameInfo[] { };
            }

            //创建文件名动态数组
            List<FileNameInfo> result = new List<FileNameInfo>();

            //逐个解析文件路径
            foreach (string? fileName in fileNames)
            {
                //如果文件路径fileName为空值，则跳过该文件路径
                if (fileName is null)
                {
                    continue;
                }

                //以文件路径fileName创建文件名对象，并添加至动态数组中
                result.Add(new FileNameInfo(fileName));
            }

            //返回生成的文件名数组
            return result.ToArray();
        }



        /// <summary>
        /// 检查文件夹是否有访问权限
        /// </summary>
        /// <param name="directoryPath">文件夹路径</param>
        /// <returns>返回文件夹是否有访问权限</returns>
        static bool HasDirectoryAccess(string directoryPath)
        {
            try
            {
                // 使用 Directory.GetFiles() 只是为了检查是否有访问权限，不会实际获取文件列表
                string[] files = Directory.GetFiles(directoryPath);
                return true;
            }
            catch (UnauthorizedAccessException)
            {
                return false;
            }
        }

        /// <summary>
        /// 递归获取文件夹及其子文件夹中的所有文件列表
        /// </summary>
        /// <param name="folderPath">要解析的文件夹路径</param>
        /// <returns>返回从文件夹folderPath中所有的文件及其及文件所拥有的文件夹</returns>
        public static string[] GetFilesRecursively(string folderPath)
        {
            List<string> files = new List<string>();
            foreach (var file in Directory.GetFiles(folderPath))
            {
                FileSystemInfo fileSystemInfo = new FileInfo(file);
                if (File.Exists(file))
                {
                    files.Add(file);
                }
            }
            foreach (var folder in Directory.GetDirectories(folderPath))
            {
                FileSystemInfo fileSystemInfo = new FileInfo(folder);
                if (Directory.Exists(folder) && HasDirectoryAccess(folder))
                {
                    files.AddRange(GetFilesRecursively(folder));
                }
            }
            return files.ToArray();
        }

        /// <summary>
        /// 获取文件夹的空间大小
        /// </summary>
        /// <param name="folderPath">文件夹路径</param>
        /// <returns>返回文件夹的空间大小</returns>
        public static long GetFileFolderSize(string folderPath)
        {
            string[] files = GetFilesRecursively(folderPath);
            long result = 0;
            foreach (string file in files)
            {
                // 读取文件大小并累加到总大小中
                FileInfo fileInfo = new FileInfo(file);
                result += fileInfo.Length;
            }
            return result;
        }


        /// <summary>
        /// 获取空数组
        /// </summary>
        /// <typeparam name="T">数组的类</typeparam>
        /// <returns>返回一个空的T数组</returns>
        public static T[] GetEmptyArray<T>()
        {
            return Array.Empty<T>();
        }




        #endregion


        #region 字符串


        public const string emptyString = @"";


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
            Match match = regexNumber.Match(str);
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
            Match match = regexNumber.Match(str);
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
            Match match = regexNumber.Match(str);
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






        public readonly static Action<string> logConsole = (str) => Console.WriteLine(str);
#if DEBUG
        public readonly static Action<string>? logTest = (str)=> Console.WriteLine(str);
#else
        public readonly static Action<string>? logTest = null;
#endif

        public class Workflow
        {
            private Stopwatch globalStopwatch = new Stopwatch();
            private Stopwatch stopwatch = new Stopwatch();
            private string subName;
            private Action<string>? log;
            private string workingContent = String.Empty;
            public List<Task> tasks { get; set; } = new List<Task>();
            public string WorkingContent
            {
                get
                {
                    return this.workingContent;
                }
                set
                {
                    if (!string.IsNullOrEmpty(this.workingContent) && this.log is not null)
                    {
                        this.log($"进程：{this.subName}：{this.workingContent}成功！！！" + stopwatch.ClockString());
                    }
                    this.stopwatch.Restart();
                    this.workingContent = value;
                }
            }
            public Workflow(Action<string>? log, string subName)
            {
                this.log = log;
                this.globalStopwatch.Restart();
                this.subName = subName;
                if (this.log is not null)
                {
                    this.log(@"-----------------------------------------");
                }
            }

            public void AddTask(Task task,string? workingContent = null)
            {
                Stopwatch stopwatch = Stopwatch.StartNew();
                task.Start();
                this.tasks.Add(task);
                task.Wait();
                if (!string.IsNullOrEmpty(this.workingContent) && this.log is not null)
                {
                    this.log($"进程：{this.subName}：{this.workingContent}成功！！！" + stopwatch.ClockString());
                }

            }
            
            public void Log(string str)
            {
                if(this.log != null)
                {
                    this.log(str);
                }
            }


            public Workflow(string subName, Action<string>? log)
            {
                this.log = log;
                this.globalStopwatch.Restart();
                this.subName = subName;
                if (this.log is not null)
                {
                    this.log(@"-----------------------------------------");
                }
                    
            }

            public Workflow(string subName)
            {
                this.log = null;
                this.globalStopwatch.Restart();
                this.subName = subName;
                if(this.log is not null)
                {
                    this.log(@"-----------------------------------------");
                }
            }
            public void End()
            {

                Task.WhenAll(this.tasks).Wait();
                if(this.log is not null)
                {
                    if (!string.IsNullOrEmpty(this.workingContent))
                    {
                        this.log($"进程：{this.subName}：{this.workingContent}成功！！！" + stopwatch.ClockString());
                    }
                    this.log(@"==============================================");
                    this.log($"进程：{subName}：结束！！！！" + this.globalStopwatch.ClockString());
                }

            }
            public void Stop(string message = KalevaAalto.Main.emptyString)
            {
                if(this.log is not null)
                {
                    if (string.IsNullOrEmpty(message))
                    {
                        this.log($"进程：{this.subName}：{this.workingContent}成功！！！" + stopwatch.ClockString());
                    }
                    else
                    {
                        this.log($"进程：{this.subName}：{this.workingContent}成功，{message}！！！" + stopwatch.ClockString());
                    }
                }
                this.workingContent = string.Empty;
            }

            public void Error(Exception error)
            {
                if(this.log is not null)
                {
                    this.log(@"==============================================");
                    Regex regex = new Regex(@"\s+");
                    string errorMessage = regex.Replace(error.Source + "：" + error.Message, " ");
                    this.log($"进程：{this.subName}：{(string.IsNullOrEmpty(this.workingContent) ? this.subName : this.workingContent)}：异常：{errorMessage}");
                }
            }

            public void Error(string errorMessage)
            {
                if (this.log is not null)
                {
                    this.log(@"==============================================");
                    this.log($"进程：{this.subName}：{(string.IsNullOrEmpty(this.workingContent) ? this.subName : this.workingContent)}：异常：{errorMessage}");
                }
                    
            }

        }



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


        #region 异常处理
        /// <summary>
        /// 以字符串的形式抛出异常
        /// </summary>
        /// <param name="errorMessage">抛出异常所表示的字符串</param>
        /// <param name="isSign">是否添加"失败，请重试......"描述，默认为否</param>
        public static void ThrowError(string errorMessage, bool isSign = false)
        {
            if (isSign) throw new Exception(errorMessage + "失败，请重试......");
            else throw new Exception(errorMessage);
        }
        #endregion


    }
}