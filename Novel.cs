using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using System.Data;
using static System.Net.Mime.MediaTypeNames;
using System.IO;
using System.IO.Compression;
using System.Xml;
using System.Xml.Linq;
using System.Net.Http.Headers;
using System.Reflection;
using K4os.Compression.LZ4.Streams.Abstractions;
using Org.BouncyCastle.Bcpg;
using static KalevaAalto.Main;
using System.Data.SqlTypes;

namespace KalevaAalto
{

    public class Title
    {
        public string Name;
        public int pos;
        public Title(string name, int pos)
        {
            Name = name;
            this.pos = pos;
        }
    }

    /// <summary>
    /// 小说章节类
    /// </summary>
    public class NovelChapter
    {
        /// <summary>
        /// 小说章节名称
        /// </summary>
        public string name { get; set; }

        /// <summary>
        /// 小说章节内容，分段落
        /// </summary>
        public List<string> paragraphs { get; private set; } = new List<string>();

        private static Regex defaultRegex = new Regex(@"\S[\S ]*");

        /// <summary>
        /// 根据小说章节内容来分段落
        /// </summary>
        /// <param name="content">小说章节内容</param>
        /// <returns></returns>
        public static string[] Split(string content)
        {
            List<string> result = new List<string>();

            // 获取第一个匹配
            Match match = defaultRegex.Match(content);

            // 循环遍历所有匹配
            while (match.Success)
            {
                // 获取分组的内容(如 (.*))
                string group = match.Groups[0].Value;
                result.Add(group.Trim());

                // 获取下一个匹配
                match = match.NextMatch();
            }
            return result.ToArray();
        }


        /// <summary>
        /// 根据小说章节名和内容来创建小说章节对象
        /// </summary>
        /// <param name="name">小说章节名</param>
        /// <param name="content">小说章节内容</param>
        public NovelChapter(string name, string content = "")
        {
            this.name = name;
            this.Append(content);
        }

        /// <summary>
        /// 根据小说章节名和内容来创建小说章节对象
        /// </summary>
        /// <param name="name">小说章节名</param>
        /// <param name="contents">小说章节内容</param>
        public NovelChapter(string name, string[] contents)
        {
            this.name = name;
            this.Append(contents);
        }


        /// <summary>
        /// 给小说章节添加段落
        /// </summary>
        /// <param name="content">要添加的内容</param>
        public void Append(string content)
        {
            this.paragraphs.InsertRange(this.paragraphs.Count, Split(content));
        }

        /// <summary>
        /// 给小说章节添加段落
        /// </summary>
        /// <param name="contents">要添加的内容</param>
        public void Append(string[] contents)
        {
            foreach(string content in contents)
            {
                this.paragraphs.InsertRange(this.paragraphs.Count, Split(content));
            }
        }


        /// <summary>
        /// 清空小说章节的所有段落
        /// </summary>
        public void Clear()
        {
            this.paragraphs.Clear();
        }


        /// <summary>
        /// 将小说章节转换为字符串
        /// </summary>
        /// <param name="lineBreak">段落区分字符串，默认为换行符加空隔符</param>
        /// <returns>返回小说章节的字符串形式</returns>
        public string ToString(string lineBreak = "\n    ")
        {
            StringBuilder rs = new StringBuilder(this.name);


            //添加内容
            foreach (string line in this.paragraphs)
            {
                rs.Append(lineBreak);
                rs.Append(line);
            }


            return rs.ToString();
        }


        /// <summary>
        /// 小说章节的长度
        /// </summary>
        public int Length
        {
            get
            {
                int result = 0;
                foreach (string content in this.paragraphs)
                {
                    result += content.Length;
                }
                return result;
            }
        }


        public XmlDocument htmlDocument
        {
            get
            {
                XmlDocument text_xml = new XmlDocument();
                text_xml.AppendChild(text_xml.CreateXmlDeclaration(@"1.0", @"utf-8", null));
                XmlElement html = text_xml.CreateElement(@"html");
                text_xml.AppendChild(html);
                {
                    //添加标题
                    XmlElement head = text_xml.CreateElement(@"head");
                    html.AppendChild(head);
                    {
                        XmlElement title = text_xml.CreateElement(@"title");
                        head.AppendChild(title);
                        title.InnerText = this.name;
                    }

                    //添加章节内容
                    XmlElement body = text_xml.CreateElement(@"body");
                    html.AppendChild(body);
                    {
                        XmlElement h2 = text_xml.CreateElement(@"h2");
                        body.AppendChild(h2);
                        h2.InnerText = this.name;

                        foreach (string line in this.paragraphs)
                        {
                            XmlElement p = text_xml.CreateElement(@"p");
                            body.AppendChild(p);
                            p.InnerText = line;
                        }

                    }
                }
                return text_xml;
            }
        }

    }


    public class Novel
    {
        /// <summary>
        /// 小说文本格式类
        /// </summary>
        public enum NovelFileFormat
        {
            txt, epub, xml
        }

        public static Dictionary<NovelFileFormat, string> novelFileFormatStrings = new Dictionary<NovelFileFormat, string> {
            { NovelFileFormat.txt , @"txt"},
            { NovelFileFormat.epub , @"epub"},
            { NovelFileFormat.xml , @"xml"},
        };

        public static Dictionary<string, NovelFileFormat> novelFileFormats = new Dictionary<string, NovelFileFormat> {
            {@"txt",NovelFileFormat.txt},
            {@"epub",NovelFileFormat.epub},
            {@"xml",NovelFileFormat.xml},
        };



        /// <summary>
        /// 小说名称
        /// </summary>
        public string name { get; set; } = String.Empty;

        /// <summary>
        /// 小说序章
        /// </summary>
        public List<string> prologue { get; private set; } = new List<string>();

        /// <summary>
        /// 小说序章的小说章节对象
        /// </summary>
        public NovelChapter prologueChapter
        {
            get
            {
                return new NovelChapter(@"序章",this.prologue.ToArray());
            }
        }

        public List<NovelChapter> chapters { get; private set; } = new List<NovelChapter>();

        /// <summary>
        /// 生成空小说对象
        /// </summary>
        public Novel()
        {

        }

        



        /// <summary>
        /// 根据小说名称来创建小说对象
        /// </summary>
        /// <param name="name">小说名称</param>
        public Novel(string name)
        {
            this.name = name;
        }

        //获取小说内容
        public Novel(string novelName, string content, string pattern)
        {
            this.name = novelName;
            string[] contentLines = NovelChapter.Split(content);
            Regex regex = new Regex("^" + pattern + "$");

            int pos = 0;


            //获取小说序章
            while (pos < contentLines.Length && !regex.IsMatch(contentLines[pos]))
            {
                this.prologue.Add(contentLines[pos]);
                pos++;
            }


            //获取小说章节
            while (pos < contentLines.Length)
            {
                if (regex.IsMatch(contentLines[pos]))
                {
                    this.chapters.Add(new NovelChapter(contentLines[pos]));
                }
                else
                {
                    NovelChapter last_chapter = this.chapters.Last();
                    last_chapter.Append(contentLines[pos]);
                }
                pos++;
            }

        }


        /// <param name="lineBreak">段落区分字符串</param>
        /// <returns>返回小说序章的字符串形式</returns>
        public string PrologueString(string lineBreak = "\n    ")
        {
            StringBuilder result = new StringBuilder();
            //添加序章
            foreach (string line in this.prologue)
            {
                result.Append(lineBreak);
                result.Append(line);
                
            }
            
            return result.ToString();
        }

        /// <param name="lineBreak">段落区分字符串</param>
        /// <returns>返回小说的字符串形式</returns>
        public string ToString(string lineBreak = "\n")
        {
            StringBuilder result = new StringBuilder(lineBreak + lineBreak);

            //添加序章
            result.Append(this.PrologueString());

            result.Append(lineBreak + lineBreak);

            foreach (NovelChapter chapter in this.chapters)
            {
                result.Append(chapter.ToString(lineBreak));
                result.Append(lineBreak + lineBreak);
            }


            return result.ToString();
        }

        /// <summary>
        /// 小说序章的字数长度
        /// </summary>
        public int prologueLength
        {
            get
            {
                int result = 0;
                foreach (string line in this.prologue)
                {
                    result += line.Length;
                }
                return result;
            }
        }

        /// <summary>
        /// 小说的字数长度
        /// </summary>
        public int Length
        {
            get
            {
                int result = this.prologueLength;
                foreach (NovelChapter chapter in this.chapters)
                {
                    result += chapter.name.Length;
                    result += chapter.Length;
                }
                return result;
            }

        }


        /// <param name="index">小说章节号</param>
        /// <returns>返回小说章节号所对应的小说章节</returns>
        public NovelChapter this[int index]
        {
            get
            {
                if (index <= this.chapters.Count && index > 0)
                {
                    return this.chapters[index - 1];
                }
                else
                {
                    return new NovelChapter("序章", this.prologue.ToArray());
                }
            }

        }

        /// <param name="index">小说章节名</param>
        /// <returns>返回小说章节名所对应的小说章节</returns>
        public NovelChapter this[string name]
        {
            get
            {
                foreach (NovelChapter chapter in this.chapters)
                {
                    if (chapter.name == name) return chapter;
                }
                return new NovelChapter("序章", this.prologue.ToArray());
            }

        }


        /// <summary>
        /// 删除章节
        /// </summary>
        /// <param name="name">要删除的章节名</param>
        public void deleteChapter(string name)
        {
            if(name == @"序章")
            {
                this.prologue.Clear();
                return;
            }
            foreach (NovelChapter chapter in this.chapters)
            {
                if (chapter.name == name) 
                { 
                    this.chapters.Remove(chapter); break; 
                }
            }
        }

        /// <summary>
        /// 删除章节
        /// </summary>
        /// <param name="name">要删除的章节序号</param>
        public void deleteChapter(int index)
        {
            if (index <= this.chapters.Count && index > 0)
            {
                this.chapters.Remove(this.chapters[index - 1]);
            }
            else if(index == 0)
            {
                this.prologue.Clear();
            }
        }



        


        /// <summary>
        /// container.xml标准文档
        /// </summary>
        private static XmlDocument container_xml_xml
        {
            get
            {
                XmlDocument result = new XmlDocument();
                result.AppendChild(result.CreateXmlDeclaration(@"1.0", @"utf-8", null));
                {

                    XmlElement container = result.CreateElement(@"container");
                    result.AppendChild(container);
                    container.SetAttribute(@"version", @"1.0");
                    container.SetAttribute(@"xmlns", @"urn:oasis:names:tc:opendocument:xmlns:container");


                    XmlElement rootfiles = result.CreateElement(@"rootfiles");
                    container.AppendChild(rootfiles);

                    XmlElement rootfile = result.CreateElement(@"rootfile");
                    rootfiles.AppendChild(rootfile);
                    rootfile.SetAttribute(@"full-path", @"OPS/content.opf");
                    rootfile.SetAttribute(@"media-type", @"application/oebps-package+xml");

                }
                return result;
            }
        }

        public byte[] txt
        {
            get
            {
                string text = this.ToString();
                // 将字符串转换为字节数组
                return text.ToByte();
            }
        }

        /// <summary>
        /// 将小说对象保存为txt文档
        /// </summary>
        /// <param name="path">要保存的文件夹路径</param>
        /// <returns>返回程序是否运行</returns>
        public bool SaveAsTxt(string path)
        {
            this.ToString().SaveToFile(new Main.FileNameInfo(path,this.name,@"txt").fileName);
            return true;
        }

        /// <summary>
        /// 从XML文档中获取小说
        /// </summary>
        /// <param name="fileName">要读取的文件名称</param>
        /// <param name="pattern">判断标题的正则表达式</param>
        public static Novel LoadNovelFromTxt(string fileName,string pattern = @"第\d+章\-.*")
        {
            Main.FileNameInfo fileNameInfo = new Main.FileNameInfo(fileName);

            return new Novel(fileNameInfo.name, Main.GetStringFromFile(fileName), pattern);
        }


        /// <summary>
        /// 从XML文档中获取小说
        /// </summary>
        /// <param name="fileName">要读取的文件名称</param>
        /// <param name="pattern">判断标题的正则表达式</param>
        public static Novel LoadNovelFromTxt(string novelName,byte[] bytes, string pattern = @"第\d+章\-.*")
        {
            using (MemoryStream memoryStream = new MemoryStream(bytes))
            {
                Encoding encoding = memoryStream.GetEncoding();
                string content = encoding.GetString(memoryStream.ToArray());
                return new Novel(novelName,content, pattern);
            }
        }

        /// <summary>
        /// 将小说对象保存为xml文档
        /// </summary>
        /// <param name="path">要保存的文件夹路径</param>
        /// <returns>返回程序是否运行</returns>
        public bool SaveAsXml(string path)
        {
            this.xml.Save(new Main.FileNameInfo(path,this.name,@"xml").fileName);
            return true;
        }

        public XmlDocument xml
        {
            get
            {
                XmlDocument xml = new XmlDocument();

                XmlElement novel = xml.CreateElement(@"novel");
                xml.AppendChild(novel);

                #region 添加名称
                XmlElement title = xml.CreateElement(@"title");
                novel.AppendChild(title);
                title.InnerText = this.name;
                #endregion

                #region 添加序章
                XmlElement prologue = xml.CreateElement(@"prologue");
                novel.AppendChild(prologue);
                foreach (string str in this.prologue)
                {
                    XmlElement p = xml.CreateElement(@"p");
                    prologue.AppendChild(p);
                    p.InnerText = str;
                }
                #endregion


                #region 添加章节
                XmlElement chapters = xml.CreateElement(@"chapters");
                novel.AppendChild(chapters);
                foreach (NovelChapter chapter in this.chapters)
                {
                    XmlElement xml_chapter = xml.CreateElement(@"chapter");
                    chapters.AppendChild(xml_chapter);

                    XmlElement chapter_title = xml.CreateElement(@"title");
                    xml_chapter.AppendChild(chapter_title);
                    chapter_title.InnerText = chapter.name;

                    XmlElement chapter_content = xml.CreateElement(@"content");
                    xml_chapter.AppendChild(chapter_content);

                    foreach (string str in chapter.paragraphs)
                    {
                        XmlElement p = xml.CreateElement(@"p");
                        chapter_content.AppendChild(p);
                        p.InnerText = str;
                    }

                }
                #endregion

                return xml;
            }
        }


        /// <summary>
        /// 从XML文档中获取小说
        /// </summary>
        /// <param name="xml">小说的XML文档</param>
        public static Novel LoadNovelFromXml(XmlDocument xml, string novelName = Main.emptyString)
        {
            Novel result = new Novel(novelName);

            if (xml.DocumentElement is null)
            {
                throw new Exception(@"此XML文档为空！");
            }

            #region 录入标题
            if (string.IsNullOrEmpty(result.name))
            {
                var title = xml.DocumentElement.SelectSingleNode("title");
                if (title is null)
                {
                    throw new Exception(@"此XML中找不到小说标题！！！");
                }
                else
                {
                    result.name = title.InnerText;
                }
            }
            #endregion





            #region 录入序章
            var prologue = xml.DocumentElement.SelectSingleNode("prologue");
            if (prologue is null)
            {
                throw new Exception(@"此XML中找不到小说序章！！！");
            }
            foreach (XmlNode p in prologue.ChildNodes)
            {
                result.prologue.Add(p.InnerText);
            }
            #endregion


            #region 录入章节
            var chapters = xml.DocumentElement.SelectSingleNode("chapters");
            if (chapters is null)
            {
                throw new Exception(@"此XML中找不到小说章节！！！");
            }
            foreach (XmlNode chapter_node in chapters.ChildNodes)
            {
                var chapter_name = chapter_node.SelectSingleNode("title");
                if (chapter_name is null)
                {
                    throw new Exception(@"此XML中小说章节格式有误！！！");
                }
                NovelChapter chapter = new NovelChapter(chapter_name.InnerText);
                result.chapters.Add(chapter);

                var content = chapter_node.SelectSingleNode("content");
                if (content is null)
                {
                    throw new Exception($"此XML中小说“{chapter.name}”章节格式有误！！！");
                }
                foreach (XmlNode p in content.ChildNodes)
                {
                    chapter.paragraphs.Add(p.InnerText);
                }
            }
            #endregion


            return result;


        }


        /// <summary>
        /// 从XML文档中获取小说
        /// </summary>
        /// <param name="bytes">小说的XML文档的二进制形式</param>
        public static Novel LoadNovelFromXml(byte[] bytes, string novelName = Main.emptyString)
        {
            using (MemoryStream memoryStream = new MemoryStream(bytes))
            {
                // 使用 XmlDocument 加载 Memory Stream 中的数据
                XmlDocument xmlDoc = new XmlDocument();
                xmlDoc.Load(memoryStream);

                return LoadNovelFromXml(xmlDoc, novelName);
            }
        }


        /// <summary>
        /// 从XML文档中获取小说
        /// </summary>
        /// <param name="xml">小说的XML文档</param>
        public static Novel LoadNovelFromXml(string fileName)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(fileName);
            return LoadNovelFromXml(xml,new Main.FileNameInfo(fileName).name);
        }



        public byte[] epub
        {
            get
            {
                //生成UUID
                string uuid;  //获取到一个随机的UUID码
                {
                    StringBuilder rs = new StringBuilder(@"urn:uuid:");

                    const string chars = @"abdcefghijklmnopqrstuvwxyz0123456789";    //设定可生成的的随机字符
                    Random random = new Random();
                    int[] lengths = { 8, 4, 4, 4, 12 };

                    foreach (int length in lengths)
                    {
                        rs.Append(GetRandomString(chars, length));
                        rs.Append('-');
                    }
                    if (rs.Length > 0) rs.Remove(rs.Length - 1, 1);

                    uuid = rs.ToString();
                }



                #region 创建 content.opf文档
                XmlDocument content_opf_xml = new XmlDocument();
                content_opf_xml.AppendChild(content_opf_xml.CreateXmlDeclaration(@"1.0", @"utf-8", null));
                {

                    XmlElement package = content_opf_xml.CreateElement(@"package");
                    content_opf_xml.AppendChild(package);
                    package.SetAttribute(@"xmlns", "http://www.idpf.org/2007/opf");
                    package.SetAttribute(@"unique-identifier", @"BookId");
                    package.SetAttribute(@"version", @"2.0");
                    {
                        XmlElement metadata = content_opf_xml.CreateElement(@"metadata");
                        package.AppendChild(metadata);
                        metadata.SetAttribute(@"xmlns:dc", @"http://purl.org/dc/elements/1.1/");
                        metadata.SetAttribute(@"xmlns:opf", @"http://www.idpf.org/2007/opf");
                        {
                            XmlElement identifier = content_opf_xml.CreateElement(@"dc", @"identifier", @"http://purl.org/dc/elements/1.1/");
                            metadata.AppendChild(identifier);
                            identifier.SetAttribute(@"id", @"BookId");
                            identifier.SetAttribute(@"opf:scheme", @"UUID");
                            identifier.InnerText = uuid;


                            XmlElement language = content_opf_xml.CreateElement(@"dc", @"language", @"http://purl.org/dc/elements/1.1/");
                            metadata.AppendChild(language);
                            language.InnerText = @"cn";

                            XmlElement title = content_opf_xml.CreateElement(@"dc", @"title", @"http://purl.org/dc/elements/1.1/");
                            metadata.AppendChild(title);
                            title.InnerText = this.name;

                            XmlElement meta = content_opf_xml.CreateElement(@"meta");
                            metadata.AppendChild(meta);
                            meta.SetAttribute(@"content", @"1.9.0");
                            meta.SetAttribute(@"name", @"Sigil version");

                            XmlElement date = content_opf_xml.CreateElement(@"dc", @"date", @"http://purl.org/dc/elements/1.1/");
                            metadata.AppendChild(date);
                            date.SetAttribute(@"opf:event", @"modification");
                            date.SetAttribute(@"xmlns:opf", @"http://www.idpf.org/2007/opf");
                            date.InnerText = DateTime.Now.ToString(@"yyyy-MM-dd");
                        }


                        XmlElement manifest = content_opf_xml.CreateElement(@"manifest");
                        package.AppendChild(manifest);
                        {
                            XmlElement item;

                            item = content_opf_xml.CreateElement(@"item");
                            manifest.AppendChild(item);
                            item.SetAttribute(@"href", @"toc.ncx");
                            item.SetAttribute(@"id", @"ncx");
                            item.SetAttribute(@"media-type", @"application/x-dtbncx+xml");

                            for (int i = 0; i <= this.chapters.Count; i++)
                            {
                                item = content_opf_xml.CreateElement(@"item");
                                manifest.AppendChild(item);
                                item.SetAttribute(@"href", $"Text/Section{i.ToString(@"0000")}.xhtml");
                                item.SetAttribute(@"id", $"Section{i.ToString(@"0000")}.xhtml");
                                item.SetAttribute(@"media-type", @"application/xhtml+xml");
                            }

                        }


                        XmlElement spine = content_opf_xml.CreateElement(@"spine");
                        package.AppendChild(spine);
                        spine.SetAttribute(@"toc", @"ncx");
                        {
                            for (int i = 0; i <= this.chapters.Count; i++)
                            {
                                XmlElement itemref = content_opf_xml.CreateElement(@"itemref");
                                spine.AppendChild(itemref);
                                itemref.SetAttribute(@"idref", $"Section{i.ToString(@"0000")}.xhtml");
                            }
                        }
                    }

                }

                #endregion




                #region 创建 toc_ncx_xml文档
                XmlDocument toc_ncx_xml = new XmlDocument();
                toc_ncx_xml.AppendChild(toc_ncx_xml.CreateXmlDeclaration(@"1.0", @"utf-8", null));
                XmlElement ncx = toc_ncx_xml.CreateElement(@"ncx");
                toc_ncx_xml.AppendChild(ncx);
                ncx.SetAttribute(@"xmlns", @"http://www.daisy.org/z3986/2005/ncx/");
                {
                    XmlElement head = toc_ncx_xml.CreateElement(@"head");
                    ncx.AppendChild(head);
                    {
                        XmlElement meta;

                        meta = toc_ncx_xml.CreateElement(@"meta");
                        head.AppendChild(meta);
                        meta.SetAttribute(@"content", uuid);
                        meta.SetAttribute(@"name", @"dtb:uid");

                        meta = toc_ncx_xml.CreateElement(@"meta");
                        head.AppendChild(meta);
                        meta.SetAttribute(@"content", @"1");
                        meta.SetAttribute(@"name", @"dtb:depth");

                        meta = toc_ncx_xml.CreateElement(@"meta");
                        head.AppendChild(meta);
                        meta.SetAttribute(@"content", @"0");
                        meta.SetAttribute(@"name", @"dtb:totalPageCount");

                        meta = toc_ncx_xml.CreateElement(@"meta");
                        head.AppendChild(meta);
                        meta.SetAttribute(@"content", @"0");
                        meta.SetAttribute(@"name", @"dtb:maxPageNumber");
                    }


                    XmlElement docTitle = toc_ncx_xml.CreateElement(@"docTitle");
                    ncx.AppendChild(docTitle);
                    {
                        XmlElement text = toc_ncx_xml.CreateElement(@"docTitle");
                        docTitle.AppendChild(text);
                        text.InnerText = this.name;
                    }


                    XmlElement navMap = toc_ncx_xml.CreateElement(@"navMap");
                    ncx.AppendChild(navMap);
                    {
                        Func<int, string, XmlElement> lambda = (order, title_name) =>
                        {
                            XmlElement navPoint = toc_ncx_xml.CreateElement(@"navPoint");
                            navPoint.SetAttribute(@"id", @"navPoint-" + order.ToString());
                            navPoint.SetAttribute(@"playOrder", order.ToString());
                            {
                                XmlElement navLabel = toc_ncx_xml.CreateElement(@"navLabel");
                                navPoint.AppendChild(navLabel);
                                {
                                    XmlElement text = toc_ncx_xml.CreateElement(@"text");
                                    navLabel.AppendChild(text);
                                    text.InnerText = title_name;
                                }


                                XmlElement content = toc_ncx_xml.CreateElement(@"content");
                                navPoint.AppendChild(content);
                                content.SetAttribute(@"src", $"Text/Section{order.ToString(@"0000")}.xhtml");
                            }

                            return navPoint;
                        };

                        navMap.AppendChild(lambda(0, @"序章"));
                        for (int i = 1; i <= this.chapters.Count; i++)
                        {
                            navMap.AppendChild(lambda(i, this.chapters[i - 1].name));
                        }


                    }


                }

                #endregion


                List<XmlDocument> xmls = new List<XmlDocument>();
                xmls.Add(this.prologueChapter.htmlDocument);
                foreach (NovelChapter chapter in this.chapters)
                {
                    xmls.Add(chapter.htmlDocument);
                }



                #region 建立zip文件并保存



                using (MemoryStream zipStream = new MemoryStream())
                {
                    ZipArchive zipFile = new ZipArchive(zipStream, ZipArchiveMode.Create, true);
                    Stream stream; //定义一个流
                    StreamWriter writer;
                    #region 创建文件夹
                    zipFile.CreateEntry(@"META-INF/");
                    zipFile.CreateEntry(@"OPS/");
                    zipFile.CreateEntry(@"OPS/Text/");
                    #endregion

                    #region 添加mimetype
                    var mimetype = zipFile.CreateEntry(@"mimetype");
                    stream = mimetype.Open();
                    writer = new StreamWriter(stream);
                    writer.WriteLine(@"application/epub+zip");
                    writer.Close();
                    stream.Close();
                    #endregion

                    #region 添加container.xml文档
                    var container_xml = zipFile.CreateEntry(@"META-INF/container.xml");
                    stream = container_xml.Open();
                    container_xml_xml.Save(stream);
                    stream.Close();
                    #endregion

                    #region 添加content.opf
                    var content_opf = zipFile.CreateEntry(@"OPS/content.opf");
                    stream = content_opf.Open();
                    content_opf_xml.Save(stream);
                    stream.Close();
                    #endregion

                    #region 添加toc.ncx
                    var toc_ncx = zipFile.CreateEntry(@"OPS/toc.ncx");
                    stream = toc_ncx.Open();
                    toc_ncx_xml.Save(stream);
                    stream.Close();
                    #endregion


                    #region 添加Text文档
                    for (int i = 0; i < xmls.Count; i++)
                    {
                        var text = zipFile.CreateEntry($"OPS/Text/Section{i.ToString(@"0000")}.xhtml");
                        stream = text.Open();
                        xmls[i].Save(stream);
                        stream.Close();
                    }
                    #endregion



                    zipFile.Dispose();//保存zip文档
                    zipStream.Seek(0, SeekOrigin.Begin);
                    #endregion

                    return zipStream.ToArray();
                }
                    

            }
        }

        /// <summary>
        /// 将小说对象保存为
        /// </summary>
        /// <param name="path">要保存到的文件夹</param>
        /// <returns>返回epub是否正确加载，是的话返回true，否则返回false</returns>
        public bool SaveAsEpub(string path)
        {

            Main.FileNameInfo fileNameInfo = new Main.FileNameInfo(path,this.name,@"epub");
            // 这里你可以将数据写入到 memoryStream 中，例如使用 memoryStream.Write 方法
            // 将MemoryStream的内容写入文件
            File.WriteAllBytes(fileNameInfo.fileName, this.epub);
            return true;

        }

        /// <summary>
        /// 从epub文件中获取小说
        /// </summary>
        /// <param name="fileName">epub文件的文件路径</param>
        public static Novel LoadNovelFromEpub(string fileName)
        {
            Main.FileNameInfo fileNameInfo = new Main.FileNameInfo(fileName);


            
            //检查后缀名是否为epub
            if (!fileNameInfo.status || fileNameInfo.suffix != @"epub")
            {
                throw new Exception(@"请选择epub文件！");
            }

            //检查文件是否存在
            if (!fileNameInfo.exists)
            {
                throw new Exception($"文件“{fileNameInfo.fileName}”不存在！");
            }

            
            try
            {
                return LoadNovelFromEpub(ZipFile.OpenRead(fileNameInfo.fileName),fileNameInfo.name);
            }
            catch(Exception ex)
            {
                throw new Exception($"读取文件“{fileName}”的过程中出现错误：{ex.Message}");
            }

            
        }

        /// <summary>
        /// 从epub文件中获取小说
        /// </summary>
        /// <param name="zipArchive">epub文件的zip文件流</param>
        /// <param name="novelName">小说标题名称</param>
        public static Novel LoadNovelFromEpub(ZipArchive zipArchive,string novelName = Main.emptyString)
        {
            Novel result = new Novel(novelName);
            XmlDocument xml = new XmlDocument();
            XmlNamespaceManager namespaceManager;
            #region 读取文件列表
            Dictionary<string, ZipArchiveEntry> zipArchiveEntrys = new Dictionary<string, ZipArchiveEntry>();
            foreach (ZipArchiveEntry zipArchiveEntry in zipArchive.Entries)
            {
                zipArchiveEntrys.Add(zipArchiveEntry.FullName.ToLower(), zipArchiveEntry);
            }
            #endregion

            #region 检查container.xml是否存在
            string containerXmlFileName = @"META-INF/container.xml".ToLower();
            if (!zipArchiveEntrys.ContainsKey(containerXmlFileName))
            {
                throw new Exception(@"找不到文件“container.xml”！");
            }
            #endregion




            #region 打开container.xml并找到content.opf的位置
            xml.Load(zipArchiveEntrys[containerXmlFileName].Open());

            // 创建命名空间管理器
            namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace(@"ns", @"urn:oasis:names:tc:opendocument:xmlns:container");

            // 使用 XPath 表达式查找 full-path 属性值
            string? contentOpfFileName = xml.SelectSingleNode(@"/ns:container/ns:rootfiles/ns:rootfile/@full-path", namespaceManager)?.Value?.ToLower() ?? null;

            if (contentOpfFileName is null || !zipArchiveEntrys.ContainsKey(contentOpfFileName))
            {
                throw new Exception(@"找不到文件“content.opf”！");
            }
            Main.FileNameInfo contentOpfFileNameInfo = new Main.FileNameInfo(contentOpfFileName);
            #endregion



            #region 打开content.opf
            List<string> herfs = new List<string>();
            xml.Load(zipArchiveEntrys[contentOpfFileName].Open());

            // 创建命名空间管理器
            namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace(@"ns", @"http://www.idpf.org/2007/opf");
            namespaceManager.AddNamespace(@"dc", @"http://purl.org/dc/elements/1.1/");

            #region 查找小说标题
            if (string.IsNullOrEmpty(result.name))
            {
                result.name = xml.SelectSingleNode(@"/ns:package/ns:metadata/dc:title", namespaceManager)?.InnerText ?? string.Empty;
                if (string.IsNullOrEmpty(result.name))
                {
                    result.name = @"空白文档";
                }
            }
            #endregion




            #region 获取小说章节列表
            XmlNodeList? itemNodes = xml.SelectNodes(@"/ns:package/ns:manifest/ns:item", namespaceManager);
            if (itemNodes is not null)
            {
                foreach (XmlNode itemNode in itemNodes)
                {
                    //检查属性是否存在
                    if (itemNode.Attributes is null)
                    {
                        continue;
                    }

                    XmlAttribute? mediaTypeAttribute = itemNode.Attributes[@"media-type"];
                    if (mediaTypeAttribute is null || mediaTypeAttribute.Value != @"application/xhtml+xml")
                    {
                        continue;
                    }

                    XmlAttribute? hrefAttribute = itemNode.Attributes[@"href"];

                    if (hrefAttribute is not null)
                    {
                        string herfFileName = contentOpfFileNameInfo.path + contentOpfFileNameInfo.sign + hrefAttribute.Value.ToLower();
                        if (zipArchiveEntrys.ContainsKey(herfFileName))
                        {
                            herfs.Add(herfFileName);
                        }

                    }

                }
            }
            #endregion

            //throw new Exception(herfs.Count.ToString());
            
            #endregion


            #region 读取小说章节
            foreach (string herf in herfs)
            {
                try
                {
                    using (Stream stream = zipArchiveEntrys[herf].Open())
                    {
                        using (StreamReader streamReader = new StreamReader(stream, Encoding.UTF8))
                        {
                            //读取整个文件内容
                            string content = streamReader.ReadToEnd()
                                .RegexReplace(@"&\S*;", string.Empty)
                                .RegexReplace(@"xmlns(:\S+)?=""\S+""", string.Empty)
                                .RegexReplace(@"<!.*>", string.Empty)
                                .RegexReplace(@"&", string.Empty)
                                .Trim();
                            xml.LoadXml(content);
                        }
                    }
                }catch
                {
                    throw new Exception($"读取文件“{herf}”时出错！");
                }
                

                XmlNode? titleNode = xml.SelectSingleNode(@"/html/head/title");
                XmlNodeList? pNodes = xml.SelectNodes(@"/html/body/p");
                NovelChapter novelChapter = new NovelChapter(@"第0000章-");

                #region 找到小说章节标题
                if (titleNode is not null && !string.IsNullOrEmpty(titleNode.InnerText))
                {
                    novelChapter.name = titleNode.InnerText;
                }
                
                #endregion


                #region 获取小说内容
                if (pNodes is not null)
                {
                    foreach (XmlNode pNode in pNodes)
                    {
                        novelChapter.Append(pNode.InnerText);
                    }
                }
                #endregion




                if (novelChapter.name == @"序章")
                {
                    result.prologue = novelChapter.paragraphs;
                }
                else
                {
                    result.chapters.Add(novelChapter);
                }


            }
            #endregion

            return result;

        }

        /// <summary>
        /// 从epub文件中获取小说
        /// </summary>
        /// <param name="zipArchive">epub文件的二进制模式</param>
        /// <param name="novelName">小说标题名称</param>
        public static Novel LoadNovelFromEpub(byte[] bytes, string novelName = Main.emptyString)
        {
            // 将 byte[] 转换为 ZipArchive
            using (MemoryStream memoryStream = new MemoryStream(bytes))
            using (ZipArchive zipArchive = new ZipArchive(memoryStream))
            {
                return LoadNovelFromEpub(zipArchive, novelName);
            }
        }

    }
}
