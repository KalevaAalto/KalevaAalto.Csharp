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


        private string _path = string.Empty;
        private string _name = string.Empty;
        private string _suffix = string.Empty;
        private bool _status = false;

        private char _sign = '\\';
        private readonly static Regex s_regex = new Regex(@"^((?<path>.+)[\\/])?(?<name>[^\\/]+?)(\.(?<suffix>[^\\\\./]+))?$");
        public FileNameInfo(string? fileName)
        {
            if (fileName is null) return;
            Match match = s_regex.Match(fileName);
            if (!match.Success)return;
            if (fileName.Contains('/')) _sign = '/';

            _path = match.Groups[@"path"].Value;
            _name = match.Groups[@"name"].Value;
            _suffix = match.Groups[@"suffix"].Value.ToLower();
            _status = true;
        }

        public FileNameInfo(string path, string name, string? suffix = null)
        {
            _path = path;
            _name = name;
            _suffix = suffix ?? string.Empty;
            _status = true;
        }

        public string Path => _path;
        public string Name => _name;
        public string Suffix => _suffix;
        public bool Status => _status;
        public bool Exists => File.Exists(FileName);
        public string FileName => string.IsNullOrEmpty(_path) ? _path + _sign + PartialFileName : PartialFileName;
        public long Size => FileInfo.Length;
        public string PartialFileName => string.IsNullOrEmpty(_suffix) ? _name + '.' + _suffix : _name;
        public FileInfo FileInfo => new FileInfo(FileName);



    }
}



namespace KalevaAalto
{
    public static partial class Static
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