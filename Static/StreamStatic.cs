using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KalevaAalto
{
    public static partial class Main
    {
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
        /// <param name="fileName">要获取字符串的文本文档的路径</param>
        /// <param name="encoding">要获取字符串的字符集，无则自动解析字符集</param>
        /// <returns>返回从路径fileName的文本文档所获取的字符串</returns>
        public static Task<string> GetStringFromFileAsync(string fileName,CancellationToken token = default, Encoding? encoding = null)
        {
            if (encoding is null)
            {
                encoding = GetEncoding(fileName);
            }
            return File.ReadAllTextAsync(fileName, encoding, token);
        }

        /// <summary>
        /// 从文本文档中获取字符串
        /// </summary>
        /// <param name="fileNameInfo">要获取字符串的文件名信息对象</param>
        /// <param name="encoding">要获取字符串的字符集，无则自动解析字符集</param>
        /// <returns>返回从路径fileName的文本文档所获取的字符串</returns>
        public static string GetStringFromFile(Models.FileSystem.FileNameInfo fileNameInfo, Encoding? encoding = null)
        {
            return GetStringFromFile(fileNameInfo.FileName, encoding);
        }

        /// <summary>
        /// 从文本文档中获取字符串
        /// </summary>
        /// <param name="fileNameInfo">要获取字符串的文件名信息对象</param>
        /// <param name="encoding">要获取字符串的字符集，无则自动解析字符集</param>
        /// <returns>返回从路径fileName的文本文档所获取的字符串</returns>
        public static Task<string> GetStringFromFileAsync(Models.FileSystem.FileNameInfo fileNameInfo, CancellationToken token = default, Encoding? encoding = null)
        {
            return GetStringFromFileAsync(fileNameInfo.FileName, token, encoding);
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
        public static void SaveToFile(this string str, Models.FileSystem.FileNameInfo fileNameInfo, Encoding? encoding = null)
        {
            str.SaveToFile(fileNameInfo.FileName, encoding);
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
    }
}
