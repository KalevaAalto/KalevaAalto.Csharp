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
using System.Data.SqlTypes;
using KalevaAalto;
using SqlSugar.DbConvert;
using Google.Protobuf.WellKnownTypes;

namespace KalevaAalto.Models
{


    /// <summary>
    /// 小说章节类
    /// </summary>
    public class NovelChapter : IDisposable
    {

        private readonly static Regex s_defaultRegex = new Regex(@".*", RegexOptions.Compiled);

        private static XmlDocument containerXmlXml
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



        public const string DefaultLineBreak = @"\n    ";
        public static string[] Split(string content)
        {
            return s_defaultRegex.Matches(content).Select(match => match.Value.Trim()).Where(it=>it.Length > 0).ToArray();
        }








        string _lineBreak = DefaultLineBreak;
        string _title;
        List<string> _paragraphs = new List<string>();
        public NovelChapter(string title, string content = emptyString)
        {
            this._title = title;
            this.Content = content;
        }



        public string LineBreak { get=>this._lineBreak; set=>this._lineBreak=value; }
        public string Title { get=>this._title; set=>this._title=value; }
        public int Length =>this.Title.Length + this._paragraphs.Sum(it => it.Length);
        public string Content
        {
            get
            {
                StringBuilder rs = new StringBuilder();
                foreach (string line in this._paragraphs)
                {
                    rs.Append(this.LineBreak);
                    rs.Append(line);
                }
                return rs.ToString();
            }
            set => this.Append(value);
        }
        public XmlDocument HtmlDocument
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
                        title.InnerText = this.Title;
                    }

                    //添加章节内容
                    XmlElement body = text_xml.CreateElement(@"body");
                    html.AppendChild(body);
                    {
                        XmlElement h2 = text_xml.CreateElement(@"h2");
                        body.AppendChild(h2);
                        h2.InnerText = this.Title;

                        foreach (string line in this._paragraphs)
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

    


        public void Append(string content)
        {
            this._paragraphs.AddRange(Split(content));
        }

        public void Clear()
        {
            this._paragraphs.Clear();
        }

        public override string ToString()
        {
            return this.Title + this.Content;
        }
        public override bool Equals(object? obj)
        {
            NovelChapter? other = obj as NovelChapter;
            if (other == null) { return false; }
            return string.Equals(this.LineBreak, other.LineBreak) && string.Equals(this.Title, other.Title) && string.Equals(this.Content,other.Content);
        }
        public override int GetHashCode()
        {
            int hash = 17;
            hash = hash * 23 + string.GetHashCode(this.LineBreak);
            hash = hash * 23 + string.GetHashCode(this.Title);
            hash = hash * 23 + string.GetHashCode(this.Content);
            return hash;
        }

        public void Dispose()
        {
            this._paragraphs.Clear();
        }
    }


    public class Novel : IDisposable
    {

        public const string DefaultTitleRegexString = @"第[〇零一两二三四五六七八九十百千万亿\d]+[卷章回节集][\-\:\s]*(?<name>.*)";

        



        private string _lineBreak = NovelChapter.DefaultLineBreak;
        private string _title;
        private List<string> _prologue = new List<string>();
        private List<NovelChapter> _chapters = new List<NovelChapter>();
        public Novel(string title, string content = emptyString, string pattern = DefaultTitleRegexString)
        {
            this._title = title;
            this.SetContent(content,pattern);
        }



















       

        public string LineBreak 
        {
            get=>this._lineBreak;
            set
            {
                this._lineBreak = value;
                foreach(NovelChapter novelChapter in this._chapters) { novelChapter.LineBreak = value; }
            }
        }
        public string Title { get=>this._title; set=>this._title=value; }
        public string Prologue
        {
            get
            {
                StringBuilder result = new StringBuilder();
                this._prologue.ForEach(line => { result.Append(this.LineBreak); result.Append(line); });
                return result.ToString();
            }
            set
            {
                this._prologue.Clear();
                this._prologue.AddRange(NovelChapter.Split(value));
            }
        }
        public int PrologueLength => this._prologue.Sum(it => it.Length);
        public string Content
        {
            get
            {
                StringBuilder result = new StringBuilder(this.LineBreak + this.LineBreak);
                result.Append(this.Prologue);
                this._chapters.ForEach(novelChapter => {
                    result.Append(this.LineBreak + this.LineBreak);
                    result.Append(novelChapter.Content);
                });
                return result.ToString();
            }
            set
            {
                this.SetContent(value);
            }
        }
        public int Length=>this.Title.Length + this.PrologueLength + this._chapters.Sum(it => it.Length);
        public NovelChapter[] Chapters =>this._chapters.ToArray();
        public List<NovelChapter> ChapterList => this._chapters;


        public XmlDocument Xml
        {
            get
            {
                XmlDocument xml = new XmlDocument();

                XmlElement novel = xml.CreateElement(@"novel");
                xml.AppendChild(novel);

                #region 添加名称
                XmlElement title = xml.CreateElement(@"title");
                novel.AppendChild(title);
                title.InnerText = this.Title;
                #endregion

                #region 添加序章
                XmlElement prologue = xml.CreateElement(@"prologue");
                novel.AppendChild(prologue);
                foreach (string str in this._prologue)
                {
                    XmlElement p = xml.CreateElement(@"p");
                    prologue.AppendChild(p);
                    p.InnerText = str;
                }
                #endregion


                #region 添加章节
                XmlElement chapters = xml.CreateElement(@"chapters");
                novel.AppendChild(chapters);
                foreach (NovelChapter chapter in this._chapters)
                {
                    XmlElement xml_chapter = xml.CreateElement(@"chapter");
                    chapters.AppendChild(xml_chapter);

                    XmlElement chapter_title = xml.CreateElement(@"title");
                    xml_chapter.AppendChild(chapter_title);
                    chapter_title.InnerText = chapter.Title;

                    XmlElement chapter_content = xml.CreateElement(@"content");
                    xml_chapter.AppendChild(chapter_content);

                    foreach (string str in NovelChapter.Split(chapter.Content))
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







        public void PrologueAppend(string content)
        {
            this._prologue.AddRange(NovelChapter.Split(content));
        }
        public void SetContent(string content, string pattern = DefaultTitleRegexString)
        {
            this._prologue.Clear();
            this._chapters.ForEach(chapter => { chapter.Dispose(); });
            this._chapters.Clear();


            Regex titleRegex = new Regex($"^{pattern}$");
            int pos = 0;
            string[] contentLines = NovelChapter.Split(content);


            //获取小说序章
            for (; pos < contentLines.Length && !titleRegex.IsMatch(contentLines[pos]); pos++)
            {
                this._prologue.Add(contentLines[pos]);
            }


            //获取小说章节
            for (; pos < contentLines.Length; pos++)
            {
                if (titleRegex.IsMatch(contentLines[pos]))
                {
                    this._chapters.Add(new NovelChapter(contentLines[pos]) { LineBreak = this.LineBreak });
                }
                else
                {
                    this._chapters.Last().Append(contentLines[pos]);
                }
            }
        }


        


        public void AddChapter(NovelChapter novelChapter)
        {
            this._chapters.Add(novelChapter);
        }
        public void DeleteChapter(string title)
        {
            if (title == @"序章")
            {
                this._prologue.Clear();
            }
            else
            {
                this._chapters.ForEach(it =>
                {
                    if (it.Title == title) { this._chapters.Remove(it); return; }
                });
            }
        }
        public void DeleteChapter(int index)
        {
            if (index == 0)
            {
                this._prologue.Clear();
            }
            else if(index>0 && index<=this._chapters.Count)
            {
                this._chapters.RemoveAt(index-1);
            }
        }

        public override string ToString() =>this.Title + this.Content;






        public void SaveAsTxt(string fileName)
        {
            System.IO.File.WriteAllText(fileName, this.ToString());
        }
        public async Task SaveAsTxtAsync(string fileName)
        {
            await System.IO.File.WriteAllTextAsync(fileName, this.ToString());
        }
        public static Novel LoadNovelFromTxt(string fileName, string pattern = DefaultTitleRegexString)
        {
            return new Novel(System.IO.Path.GetFileNameWithoutExtension(fileName), GetStringFromFile(fileName), pattern);
        }
        public static async  Task<Novel> LoadNovelFromTxtAsync(string fileName, string pattern = DefaultTitleRegexString)
        {
            return new Novel(System.IO.Path.GetFileNameWithoutExtension(fileName), await GetStringFromFileAsync(fileName), pattern);
        }




        public void SaveAsXml(string fileName)
        {
            this.Xml.Save(fileName);
        }
        public static Novel LoadNovelFromXml(XmlDocument xml)
        {


            if (xml.DocumentElement is null)
            {
                throw new Exception(@"此XML文档为空！");
            }

            #region 录入标题
            var title = xml.DocumentElement.SelectSingleNode(@"title");
            if (title == null)
            {
                throw new Exception(@"此XML中找不到小说标题！！！");
            }
            Novel result = new Novel(title.InnerText);
            #endregion





            #region 录入序章
            var prologue = xml.DocumentElement.SelectSingleNode(@"prologue");
            if (prologue == null)
            {
                throw new Exception(@"此XML中找不到小说序章！！！");
            }
            result.Prologue = @"\n".Join(prologue.ChildNodes.Cast<XmlNode>().Select(it => it.InnerText).ToArray());
            #endregion


            #region 录入章节
            var chapters = xml.DocumentElement.SelectSingleNode(@"chapters");
            if (chapters is null)
            {
                throw new Exception(@"此XML中找不到小说章节！！！");
            }
            foreach (XmlNode chapter_node in chapters.ChildNodes)
            {
                var chapter_name = chapter_node.SelectSingleNode(@"title");
                if (chapter_name == null)
                {
                    throw new Exception(@"此XML中小说章节格式有误！！！");
                }
                NovelChapter chapter = new NovelChapter(chapter_name.InnerText);
                result.AddChapter(chapter);

                var content = chapter_node.SelectSingleNode("content");
                if (content == null)
                {
                    throw new Exception($"此XML中小说“{chapter.Title}”章节格式有误！！！");
                }
                foreach (XmlNode p in content.ChildNodes)
                {
                    chapter.Append(p.InnerText);
                }
            }
            #endregion


            return result;


        }
        public static Novel LoadNovelFromXml(string fileName)
        {
            XmlDocument xml = new XmlDocument();
            xml.Load(fileName);
            return LoadNovelFromXml(xml);
        }



        private static XmlDocument s_containerXmlXml
        {
            get
            {

                {
                    XDocument xml = new XDocument(
                        new XDeclaration("1.0", "utf-8", null),
                        new XElement(@"container",
                            new XAttribute(@"version",@"1.0"),
                            new XElement(@"rootfiles",
                                new XElement(@"rootfile")
                                )
                            )
                    ) ;

                }










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
                            title.InnerText = this.Title;

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

                            for (int i = 0; i <= this._chapters.Count; i++)
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
                            for (int i = 0; i <= this._chapters.Count; i++)
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
                        text.InnerText = this.Title;
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
                        for (int i = 1; i <= this._chapters.Count; i++)
                        {
                            navMap.AppendChild(lambda(i, this._chapters[i - 1].Title));
                        }


                    }


                }

                #endregion


                List<XmlDocument> xmls = new List<XmlDocument>();
                xmls.Add(new NovelChapter(@"序章",this.Prologue).HtmlDocument);
                foreach (NovelChapter chapter in this._chapters)
                {
                    xmls.Add(chapter.HtmlDocument);
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
                    s_containerXmlXml.Save(stream);
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
        public void SaveAsEpub(string fileName)
        {
            File.WriteAllBytes(fileName, this.epub);
        }

        public async Task SaveAsEpubAsync(string fileName)
        {
            await File.WriteAllBytesAsync(fileName, this.epub);
        }

        /// <summary>
        /// 从epub文件中获取小说
        /// </summary>
        /// <param name="fileName">epub文件的文件路径</param>
        public static Novel LoadNovelFromEpub(string fileName)
        {
            Models.FileSystem.FileNameInfo fileNameInfo = new Models.FileSystem.FileNameInfo(fileName);



            //检查后缀名是否为epub
            if (!fileNameInfo.Status || fileNameInfo.Suffix != @"epub")
            {
                throw new Exception(@"请选择epub文件！");
            }

            //检查文件是否存在
            if (!fileNameInfo.Exists)
            {
                throw new Exception($"文件“{fileNameInfo.FileName}”不存在！");
            }


            try
            {
                return LoadNovelFromEpub(ZipFile.OpenRead(fileNameInfo.FileName));
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件“{fileName}”的过程中出现错误：{ex.Message}");
            }


        }
        private static Novel LoadNovelFromEpub(ZipArchive zipArchive)
        {
            Novel result = new Novel(@"");
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
            Models.FileSystem.FileNameInfo contentOpfFileNameInfo = new Models.FileSystem.FileNameInfo(contentOpfFileName);
            #endregion



            #region 打开content.opf
            List<string> herfs = new List<string>();
            xml.Load(zipArchiveEntrys[contentOpfFileName].Open());

            // 创建命名空间管理器
            namespaceManager = new XmlNamespaceManager(xml.NameTable);
            namespaceManager.AddNamespace(@"ns", @"http://www.idpf.org/2007/opf");
            namespaceManager.AddNamespace(@"dc", @"http://purl.org/dc/elements/1.1/");

            #region 查找小说标题
            if (string.IsNullOrEmpty(result.Title))
            {
                result.Title = xml.SelectSingleNode(@"/ns:package/ns:metadata/dc:title", namespaceManager)?.InnerText ?? string.Empty;
                if (string.IsNullOrEmpty(result.Title))
                {
                    result.Title = @"空白文档";
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
                        string herfFileName = contentOpfFileNameInfo.Path + contentOpfFileNameInfo.Sign + hrefAttribute.Value.ToLower();
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
                }
                catch
                {
                    throw new Exception($"读取文件“{herf}”时出错！");
                }


                XmlNode? titleNode = xml.SelectSingleNode(@"/html/head/title");
                XmlNodeList? pNodes = xml.SelectNodes(@"/html/body/p");
                NovelChapter novelChapter = new NovelChapter(@"第0000章-");

                #region 找到小说章节标题
                if (titleNode != null && !string.IsNullOrEmpty(titleNode.InnerText))
                {
                    novelChapter.Title = titleNode.InnerText;
                }

                #endregion


                #region 获取小说内容
                if (pNodes != null)
                {
                    foreach (XmlNode pNode in pNodes)
                    {
                        novelChapter.Append(pNode.InnerText);
                    }
                }
                #endregion




                if (novelChapter.Title == @"序章")
                {
                    result.Prologue = novelChapter.Content;
                }
                else
                {
                    result.AddChapter(novelChapter);
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
        public static Novel LoadNovelFromEpub(byte[] bytes, string novelName = emptyString)
        {
            // 将 byte[] 转换为 ZipArchive
            using (MemoryStream memoryStream = new MemoryStream(bytes))
            using (ZipArchive zipArchive = new ZipArchive(memoryStream))
            {
                return LoadNovelFromEpub(zipArchive);
            }
        }



        public void Dispose()
        {
            this._prologue.Clear();
            this._chapters.ForEach(chapter => { chapter.Dispose(); });
            this._chapters.Clear();
        }

    }
}
