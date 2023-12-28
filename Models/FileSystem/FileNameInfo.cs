using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace KalevaAalto.Models.FileSystem
{
    /// <summary>
    /// 用来解析文件路径的的文件名类
    /// </summary>
    public class FileNameInfo
    {
        /// <summary>
        /// 文件所在文件夹的路径
        /// </summary>
        public string Path { get; } = string.Empty;

        /// <summary>
        /// 文件名（不包含后缀名）
        /// </summary>
        public string Name { get; set; } = string.Empty;

        /// <summary>
        /// 文件后缀名
        /// </summary>
        public string Suffix { get; } = string.Empty;

        /// <summary>
        /// 文件名类的状态，异常则为false，此时的文件名对象无法正常使用
        /// </summary>
        public bool Status { get; } = false;

        /// <summary>
        /// 返回文件是否存在
        /// </summary>
        public bool Exists
        {
            get
            {
                return File.Exists(this.FileName);
            }
        }


        public char Sign { get; } = '\\';

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
            if (fileName is null)
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
                this.Sign = '/';
            }


            this.Path = match.Groups[@"path"].Value;
            this.Name = match.Groups[@"name"].Value;
            this.Suffix = match.Groups[@"suffix"].Value.ToLower();
            this.Status = true;
        }

        /// <summary>
        /// 手动文件名对象
        /// </summary>
        /// <param name="path">文件所在文件夹的路径</param>
        /// <param name="name">文件名（不包含后缀名）</param>
        /// <param name="suffix">文件后缀名</param>
        public FileNameInfo(string path, string name, string? suffix = null)
        {
            this.Path = path;
            this.Name = name;
            this.Suffix = suffix ?? string.Empty;
            this.Status = true;
        }

        /// <summary>
        /// 返回此文件名对象的文件路径
        /// </summary>
        public string FileName
        {
            get
            {
                if (this.Path.Length > 0)
                {
                    return this.Path + this.Sign + this.PartialFileName;
                }
                else
                {
                    return this.PartialFileName;
                }

            }
        }

        /// <summary>
        /// 返回此文件名对象的文件名（有后缀名）
        /// </summary>
        public string PartialFileName
        {
            get
            {
                if (this.Suffix.Length > 0)
                {
                    return this.Name + '.' + this.Suffix;
                }
                else
                {
                    return this.Name;
                }

            }
        }


        public FileInfo FileInfo
        {
            get
            {
                return new FileInfo(this.FileName);
            }
        }

    }
}



namespace KalevaAalto
{
    public static partial class Main
    {
        /// <summary>
        /// 解析的文件路径数组，并返回文件名数组
        /// </summary>
        /// <param name="fileNames">要解析的文件路径数组</param>
        /// <returns>返回从文件路径数组fileNames中解析出的文件名数组</returns>
        public static Models.FileSystem.FileNameInfo[] GetFileNameInfos(this string?[]? fileNames)
        {
            //如果文件路径数组fileNames为空值，则返回空文件名数组
            if (fileNames is null)
            {
                return new Models.FileSystem.FileNameInfo[] { };
            }

            //创建文件名动态数组
            List<Models.FileSystem.FileNameInfo> result = new List<Models.FileSystem.FileNameInfo>();

            //逐个解析文件路径
            foreach (string? fileName in fileNames)
            {
                //如果文件路径fileName为空值，则跳过该文件路径
                if (fileName is null)
                {
                    continue;
                }

                //以文件路径fileName创建文件名对象，并添加至动态数组中
                result.Add(new Models.FileSystem.FileNameInfo(fileName));
            }

            //返回生成的文件名数组
            return result.ToArray();
        }





    }
}