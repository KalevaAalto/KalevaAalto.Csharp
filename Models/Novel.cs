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
using OfficeOpenXml.ConditionalFormatting.Contracts;
using Microsoft.Extensions.FileSystemGlobbing.Internal;
using System.Xml.XPath;

namespace KalevaAalto.Models
{


    /// <summary>
    /// 小说章节类
    /// </summary>
    public class NovelChapter : IDisposable
    {

        private readonly static Regex s_defaultRegex = new Regex(@".*", RegexOptions.Compiled);




        public const int DefaultBlankCharNumber = 4;
        public const string DefaultLineBreak = "\n";
        public static string[] Split(string content)
        {
            return s_defaultRegex.Matches(content).Select(match => match.Value.Trim()).Where(it=>it.Length > 0).ToArray();
        }







        int _blankCharNumber = DefaultBlankCharNumber;
        string _lineBreak = DefaultLineBreak;
        string _title;
        List<string> _paragraphs = new List<string>();
        public NovelChapter(string title, string content = emptyString)
        {
            this._title = title;
            this.Content = content;
        }

        //int _blankCharNumber = DefaultBlankCharNumber;
        public int BlankCharNumber { get => this._blankCharNumber; set => this._blankCharNumber = System.Math.Max(0, value); }
        public string BlankString => new string(' ', BlankCharNumber);
        public string LineBreak { get=>this._lineBreak; set=>this._lineBreak=value; }
        public string Title { get=>this._title; set=>this._title=value; }
        public int Length =>this._paragraphs.Sum(it => it.Length);
        public string Content
        {
            get
            {
                StringBuilder rs = new StringBuilder();
                string blankString = BlankString;
                foreach (string line in this._paragraphs)
                {
                    rs.Append(blankString);
                    rs.Append(line);
                    rs.Append(LineBreak);
                }
                if (_paragraphs.Count > 0) rs.Remove(rs.Length-LineBreak.Length,LineBreak.Length);
                return rs.ToString();
            }
            set
            {
                this._paragraphs.Clear();
                this.Append(value);
            }
        }

        private static XDeclaration s_defaultDeclaration => new XDeclaration(@"1.0", @"utf-8", null);
        public XDocument HtmlDocument=> new XDocument(s_defaultDeclaration,
                    new XElement(@"html",
                        new XElement(@"head",
                            new XElement(@"title", this.Title)
                            ),
                        new XElement(@"body",
                            new XElement(@"h2",this.Title),
                            _paragraphs.Select(p=>new XElement(@"p",p))
                            )
                        )
                    );
            
        

    


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
            return Title + LineBreak + Content;
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
        private const string s_defaultNovelName = @"空白小说";
        public const string DefaultTitleRegexString = @"第[〇零一两二三四五六七八九十百千万亿\d]+[卷章回节集][\-\:\s]*(?<name>.*)";

        int _blankCharNumber = NovelChapter.DefaultBlankCharNumber;
        private string _lineBreak = NovelChapter.DefaultLineBreak;
        private string _title;
        private List<string> _prologue = new List<string>();
        private List<NovelChapter> _chapters = new List<NovelChapter>();
        public Novel()
        {
            _title = s_defaultNovelName;
        }
        public Novel(string title, string content = emptyString, string pattern = DefaultTitleRegexString)
        {
            this._title = title;
            this.SetContent(content,pattern);
        }

        
        public int BlankCharNumber 
        { 
            get => this._blankCharNumber;
            set
            {
                _blankCharNumber = System.Math.Max(0, value);
                foreach (NovelChapter novelChapter in this._chapters) { novelChapter.BlankCharNumber = value; }
            }
        }
        public string BlankString => new string(' ', BlankCharNumber);
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
                string blankString = BlankString;
                _prologue.ForEach(line => { result.Append(blankString); result.Append(line); result.Append(this.LineBreak); });
                if (_prologue.Count > 0) result.Remove(result.Length - LineBreak.Length, LineBreak.Length);
                return result.ToString();
            }
            set
            {
                _prologue.Clear();
                _prologue.AddRange(NovelChapter.Split(value));
            }
        }
        public int PrologueLength => this._prologue.Sum(it =>it.Length);
        public string Content
        {
            get
            {
                if (Length == 0) return string.Empty;

                StringBuilder result = new StringBuilder();
                result.Append(this.Prologue);
                result.Append(LineBreak + LineBreak);
                this._chapters.ForEach(novelChapter => {
                    result.Append(novelChapter.ToString()) ;
                    result.Append(this.LineBreak + this.LineBreak);
                });
                result.Remove(result.Length - LineBreak.Length*2, LineBreak.Length*2);
                return result.ToString();
            }
            set=>this.SetContent(value);
            
        }
        public int Length=>this.PrologueLength + this._chapters.Sum(it =>it.Title.Length + it.Length);
        public NovelChapter[] Chapters =>this._chapters.ToArray();
        public List<NovelChapter> ChapterList => this._chapters;


        public XDocument Xml =>new XDocument(s_defaultDeclaration,new XElement(@"novel"
                    ,new XElement(@"title",this.Title)
                    ,new XElement(@"prologue",_prologue.Select(it=>new XElement(@"p",it)))
                    ,new XElement(@"chapters"
                        ,_chapters.Select(it=>new XElement(@"chapter"
                            ,new XElement(@"title",it.Title)
                            ,new XElement(@"content", NovelChapter.Split(it.Content).Select(it => new XElement(@"p", it)))
                            ))
                        )
                    ));
            
        







        public void PrologueAppend(string content)
        {
            this._prologue.AddRange(NovelChapter.Split(content));
        }
        public void SetContent(string content, string pattern = DefaultTitleRegexString)
        {
            this._prologue.Clear();
            this._chapters.ForEach(chapter =>chapter.Dispose());
            this._chapters.Clear();


            Regex titleRegex = new Regex($"^{pattern}$");
            int pos = 0;
            string[] contentLines = NovelChapter.Split(content);


            //获取小说序章
            for (; pos < contentLines.Length && !titleRegex.IsMatch(contentLines[pos]); pos++)this._prologue.Add(contentLines[pos]);
            


            //获取小说章节
            for (; pos < contentLines.Length; pos++)
            {
                if (titleRegex.IsMatch(contentLines[pos]))this._chapters.Add(new NovelChapter(contentLines[pos]) { LineBreak = this.LineBreak,BlankCharNumber = BlankCharNumber });
                else this._chapters.Last().Append(contentLines[pos]);
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

        public override string ToString() =>Title + LineBreak + Content;






        public void SaveAsTxt(string fileName)
        {
            System.IO.File.WriteAllText(fileName, Content);
        }
        public async Task SaveAsTxtAsync(string fileName)
        {
            await System.IO.File.WriteAllTextAsync(fileName, Content);
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
            Xml.Save(fileName);
        }
        public async Task SaveAsXmlAsync(string fileName,CancellationToken token = default)
        {
            using (FileStream stream = File.Create(fileName))
            {
                await Xml.SaveAsync(stream, SaveOptions.None, token);
            }
        }


        public static Novel LoadNovelFromXml(string fileName)
        {
            XDocument xml = XDocument.Load(fileName);
            Novel result = new Novel();

            XElement? rootElement = xml.Root;
            if (rootElement is null) throw new Exception($"无法将“{fileName}”加载为XML文档；");

            Parallel.Invoke(() =>
            {
                XElement? titleElement = rootElement.Element(@"title");
                if (titleElement is not null && string.IsNullOrEmpty(titleElement.Value)) result.Title = titleElement.Value;
            }, () =>
            {
                XElement? prologueElement = rootElement.Element(@"prologue");
                if (prologueElement is not null)
                {
                    IEnumerable<XElement> pElements = prologueElement.Elements(@"p");
                    foreach (XElement pElement in pElements) result.PrologueAppend(pElement.Value);
                }
            }, () => {
                XElement? chaptersElement = rootElement.Element(@"chapters");
                if (chaptersElement is not null)
                {
                    IEnumerable<XElement> chapterElements = chaptersElement.Elements(@"chapter");
                    foreach (XElement chapterElement in chapterElements)
                    {
                        string chapterName = @"第0000章-未知章节";
                        XElement? chapterTitleElement = chapterElement.Element(@"title");
                        if (chapterTitleElement is not null) chapterName = chapterTitleElement.Value;
                        XElement? contentElement = chapterElement.Element(@"content");
                        if (contentElement is not null)
                        {
                            IEnumerable<XElement> pElements = contentElement.Elements(@"p");
                            NovelChapter chapter = new NovelChapter(chapterName);
                            result.AddChapter(chapter);
                            foreach (XElement pElement in pElements) chapter.Append(pElement.Value);
                        }
                    }
                }
            });



            return result;


        }
        public static async Task<Novel> LoadNovelFromXmlAsync(string fileName,CancellationToken token=default)
        {
            if (!File.Exists(fileName)) throw new Exception($"文件“{fileName}”不存在；");

            using (FileStream stream = File.Open(fileName, FileMode.Open))
            {
                XDocument xml = await XDocument.LoadAsync(stream,LoadOptions.None,token);
                Novel result = new Novel();

                XElement? rootElement = xml.Root;
                if (rootElement is null) throw new Exception($"无法将“{fileName}”加载为XML文档；");


                await Task.Run(() => {
                    Parallel.Invoke(() =>
                    {
                        XElement? titleElement = rootElement.Element(@"title");
                        if (titleElement is not null && string.IsNullOrEmpty(titleElement.Value)) result.Title = titleElement.Value;
                    }, () =>
                    {
                        XElement? prologueElement = rootElement.Element(@"prologue");
                        if (prologueElement is not null)
                        {
                            IEnumerable<XElement> pElements = prologueElement.Elements(@"p");
                            foreach (XElement pElement in pElements) result.PrologueAppend(pElement.Value);
                        }
                    }, () => {
                        XElement? chaptersElement = rootElement.Element(@"chapters");
                        if (chaptersElement is not null)
                        {
                            IEnumerable<XElement> chapterElements = chaptersElement.Elements(@"chapter");
                            foreach (XElement chapterElement in chapterElements)
                            {
                                string chapterName = @"第0000章-未知章节";
                                XElement? chapterTitleElement = chapterElement.Element(@"title");
                                if (chapterTitleElement is not null) chapterName = chapterTitleElement.Value;
                                XElement? contentElement = chapterElement.Element(@"content");
                                if (contentElement is not null)
                                {
                                    IEnumerable<XElement> pElements = contentElement.Elements(@"p");
                                    NovelChapter chapter = new NovelChapter(chapterName);
                                    result.AddChapter(chapter);
                                    foreach (XElement pElement in pElements) chapter.Append(pElement.Value);
                                }
                            }
                        }
                    });

                });

                return result;
            }


            


        }

        


        private static string s_uuid
        {
            get
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

                return rs.ToString();
            }
        }
        private string _uuid = s_uuid;
        private static XDeclaration s_defaultDeclaration => new XDeclaration("1.0", "utf-8", null);
        private readonly static XNamespace s_containerXmlXmlNs = @"urn:oasis:names:tc:opendocument:xmlns:container";
        private static XDocument s_containerXmlXml
        {
            get
            {
                XNamespace ns = s_containerXmlXmlNs;
                return new XDocument(
                        s_defaultDeclaration,
                        new XElement(ns + @"container",
                            new XAttribute(@"version", @"1.0"),
                            new XElement(ns + @"rootfiles",
                                new XElement(ns + @"rootfile", new XAttribute(@"full-path", @"OPS/content.opf"), new XAttribute(@"media-type", @"application/oebps-package+xml"))
                                )
                            )
                    );
            }
        }

        private readonly static XNamespace s_contentOpfXmlNs = @"http://www.idpf.org/2007/opf";
        private readonly static XNamespace s_contentOpfXmlDcNs = @"http://purl.org/dc/elements/1.1/";
        private XDocument contentOpfXml
        {
            get
            {
                XNamespace ns = s_contentOpfXmlNs;
                XNamespace dcNs = s_contentOpfXmlDcNs;


                XElement rootElement = new XElement(ns + @"package", new XAttribute(@"unique-identifier", @"BookId"), new XAttribute(@"version", @"2.0"));
                XDocument result = new XDocument(s_defaultDeclaration, rootElement);

                rootElement.Add(
                    new XElement(ns + @"metadata", new XAttribute(XNamespace.Xmlns + @"dc", dcNs),
                            new XElement(dcNs + @"identifier", new XAttribute(@"id", @"BookId"), new XAttribute(@"scheme", @"UUID"), _uuid),
                            new XElement(dcNs + @"language", @"cn"),
                            new XElement(dcNs + @"title", this.Title),
                            new XElement(ns + @"meta", new XAttribute(@"content", @"1.9.0"), new XAttribute(@"name", @"Sigil version")),
                            new XElement(dcNs + @"date", new XAttribute(@"event", @"modification"), DateTime.Today.ToString(@"yyyy-MM-dd"))
                            )
                    );
                XElement manifestElement = new XElement(ns + @"manifest", new XElement(ns + @"item", new XAttribute(@"href", @"toc.ncx"), new XAttribute(@"id", @"ncx"), new XAttribute(@"media-type", @"application/x-dtbncx+xml")));
                for(int i = 0; i <= _chapters.Count; i++)
                {
                    manifestElement.Add(new XElement(ns + @"item", new XAttribute(@"href", $"Text/Section{i.ToString(@"0000")}.xhtml"), new XAttribute(@"id", $"Section{i.ToString(@"0000")}.xhtml"), new XAttribute(@"media-type", @"application/xhtml+xml")));
                }


                rootElement.Add(manifestElement);
                rootElement.Add(
                    new XElement(ns + @"spine", new XAttribute(@"toc", @"ncx"),
                            Enumerable.Range(0, _chapters.Count).Select(i => new XElement(ns + @"itemref", new XAttribute(@"idref", $"Section{i.ToString(@"0000")}.xhtml")))
                            )
                    );

                return result;
            }
        }

        private readonly static XNamespace s_tocNcxXmlNs = @"http://www.daisy.org/z3986/2005/ncx/";
        private XDocument tocNcxXml
        {
            get
            {
                XNamespace ns = s_tocNcxXmlNs;
                return new XDocument(s_defaultDeclaration,
                    new XElement(ns + @"ncx",
                        new XElement(ns + @"head",
                            new XElement(ns + @"meta", new XAttribute(@"content", _uuid), new XAttribute(@"name", @"dtb:uid")),
                            new XElement(ns + @"meta", new XAttribute(@"content", @"1"), new XAttribute(@"name", @"dtb:depth")),
                            new XElement(ns + @"meta", new XAttribute(@"content", @"0"), new XAttribute(@"name", @"dtb:totalPageCount")),
                            new XElement(ns + @"meta", new XAttribute(@"content", @"0"), new XAttribute(@"name", @"dtb:maxPageNumber"))
                            ),
                        new XElement(ns + @"docTitle", new XElement(ns + @"docTitle", this.Title)),
                        new XElement(ns + @"navMap",
                            Enumerable.Range(0, _chapters.Count).Select(i =>
                            {
                                return new XElement(ns + @"navPoint", new XAttribute(@"id", $"navPoint-{i}"), new XAttribute(@"playOrder", i.ToString()),
                                    new XElement(ns+ @"navLabel",new XElement(ns + @"text", i == 0 ? @"序章" : _chapters[i].Title)),
                                    new XElement(ns+ @"content",new XAttribute(@"src",$"Text/Section{i.ToString(@"0000")}.xhtml"))
                                    );
                            })
                            )
                    ));
                    
            }
        }
        public byte[] GetEpub()
        {
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
                stream = zipFile.CreateEntry(@"mimetype").Open();
                writer = new StreamWriter(stream);
                writer.WriteLine(@"application/epub+zip");
                writer.Close();
                stream.Close();
                #endregion

                #region 添加container.xml文档
                stream = zipFile.CreateEntry(@"META-INF/container.xml").Open();
                s_containerXmlXml.Save(stream);
                stream.Close();
                #endregion

                #region 添加content.opf
                stream = zipFile.CreateEntry(@"OPS/content.opf").Open();
                contentOpfXml.Save(stream);
                stream.Close();
                #endregion



                #region 添加toc.ncx
                stream = zipFile.CreateEntry(@"OPS/toc.ncx").Open();
                tocNcxXml.Save(stream);
                stream.Close();
                #endregion


                #region 添加Text文档
                XDocument[] xDocuments = _chapters.AsParallel().Select(it => it.HtmlDocument).ToArray();
                for (int i = 0; i < xDocuments.Length; i++)
                {
                    stream = zipFile.CreateEntry($"OPS/Text/Section{i.ToString(@"0000")}.xhtml").Open();
                    xDocuments[i].Save(stream);
                    stream.Close();
                }
                #endregion



                zipFile.Dispose();//保存zip文档
                zipStream.Seek(0, SeekOrigin.Begin);


                return zipStream.ToArray();
            }
        }
        public async Task<byte[]> GetEpubAsync(CancellationToken cancellationToken = default)
        {

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
                stream = zipFile.CreateEntry(@"mimetype").Open();
                writer = new StreamWriter(stream);
                await writer.WriteAsync(@"application/epub+zip");
                writer.Close();
                stream.Close();
                #endregion

                #region 添加container.xml文档
                stream = zipFile.CreateEntry(@"META-INF/container.xml").Open();
                await s_containerXmlXml.SaveAsync(stream, SaveOptions.None,cancellationToken);
                stream.Close();
                #endregion

                #region 添加content.opf
                stream = zipFile.CreateEntry(@"OPS/content.opf").Open();
                await contentOpfXml.SaveAsync(stream, SaveOptions.None, cancellationToken);
                stream.Close();
                #endregion



                #region 添加toc.ncx
                stream = zipFile.CreateEntry(@"OPS/toc.ncx").Open();
                await tocNcxXml.SaveAsync(stream, SaveOptions.None, cancellationToken);
                stream.Close();
                #endregion


                #region 添加Text文档
                XDocument[] xDocuments = _chapters.AsParallel().Select(it => it.HtmlDocument).ToArray();
                for (int i = 0; i < xDocuments.Length; i++)
                {
                    stream = zipFile.CreateEntry($"OPS/Text/Section{i.ToString(@"0000")}.xhtml").Open();
                    await xDocuments[i].SaveAsync(stream, SaveOptions.None, cancellationToken);
                    stream.Close();
                }
                #endregion



                zipFile.Dispose();//保存zip文档
                zipStream.Seek(0, SeekOrigin.Begin);


                return zipStream.ToArray();
            }
        }

        public void SaveAsEpub(string fileName)
        {
            File.WriteAllBytes(fileName, this.GetEpub());
        }
        public async Task SaveAsEpubAsync(string fileName, CancellationToken cancellationToken = default)
        {
            await File.WriteAllBytesAsync(fileName,await this.GetEpubAsync(cancellationToken));
        }


        public static Novel LoadNovelFromEpub(string fileName)
        {

            if (Path.GetExtension(fileName) != @".epub")throw new Exception(@"请选择epub文件！");
            if (!File.Exists(fileName))throw new Exception($"文件“{fileName}”不存在！");
            try
            {
                return LoadNovelFromEpub(ZipFile.OpenRead(fileName));
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件“{fileName}”的过程中出现错误：{ex.Message}");
            }
        }
        public static async Task<Novel> LoadNovelFromEpubAsync(string fileName, CancellationToken token = default)
        {

            if (Path.GetExtension(fileName) != @".epub") throw new Exception(@"请选择epub文件！");
            if (!File.Exists(fileName)) throw new Exception($"文件“{fileName}”不存在！");
            try
            {
                return await LoadNovelFromEpubAsync(ZipFile.OpenRead(fileName),token);
            }
            catch (Exception ex)
            {
                throw new Exception($"读取文件“{fileName}”的过程中出现错误：{ex.Message}");
            }
        }
        private static Novel LoadNovelFromEpub(ZipArchive zipArchive)
        {
            Novel result = new Novel();
            XDocument xml;
            XmlNamespaceManager namespaceManager;
            Dictionary<string, ZipArchiveEntry> zipArchiveEntrys = zipArchive.Entries.Cast<ZipArchiveEntry>().ToDictionary(it => it.FullName.ToLower(), it=>it);


            //container.xml
            string containerXmlFileName = @"META-INF/container.xml".ToLower();
            if (!zipArchiveEntrys.ContainsKey(containerXmlFileName))throw new Exception(@"压缩文件包中找不到文件“container.xml”！");
            xml = XDocument.Load(zipArchiveEntrys[containerXmlFileName].Open());
            namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace(@"ns", s_containerXmlXmlNs.NamespaceName);
            string? contentOpfFileName = xml.XPathSelectElement(@"/ns:container/ns:rootfiles/ns:rootfile", namespaceManager)?.Attribute(@"full-path")?.Value.ToLower() ?? null;
            if (contentOpfFileName is null || !zipArchiveEntrys.ContainsKey(contentOpfFileName))throw new Exception(@"压缩文件包中找不到文件“content.opf”！");


            //content.opf
            List<string> herfs = new List<string>();
            xml = XDocument.Load(zipArchiveEntrys[contentOpfFileName].Open());

            namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace(@"ns", s_contentOpfXmlNs.NamespaceName);
            namespaceManager.AddNamespace(@"dc", s_contentOpfXmlDcNs.NamespaceName);

            string title = xml.XPathSelectElement(@"/ns:package/ns:metadata/dc:title", namespaceManager)?.Value ?? string.Empty;
            if (!string.IsNullOrEmpty(result.Title))result.Title = title;
            
            IEnumerable<XElement> itemNodes = xml.XPathSelectElements(@"/ns:package/ns:manifest/ns:item", namespaceManager);

            itemNodes.ToList().ForEach(itemNode =>
            {
                XAttribute? mediaTypeAttribute = itemNode.Attribute(@"media-type");
                if (mediaTypeAttribute is null || mediaTypeAttribute.Value != @"application/xhtml+xml")
                {
                    return;
                }

                XAttribute? hrefAttribute = itemNode.Attribute(@"href");
                if (hrefAttribute is not null)
                {
                    string herfFileName = Path.Combine(Path.GetDirectoryName(contentOpfFileName)??string.Empty,hrefAttribute.Value.ToLower());
                    if (zipArchiveEntrys.ContainsKey(herfFileName)) herfs.Add(herfFileName);
                }
            });

            
            herfs.ForEach(herf =>
            {
                xml = XDocument.Load(zipArchiveEntrys[herf].Open());

                XElement? titleNode = xml.XPathSelectElement(@"/html/head/title");
                IEnumerable<XElement> pNodes = xml.XPathSelectElements(@"/html/body/p");
                NovelChapter novelChapter = new NovelChapter(@"第0000章-未知标题");

                if (titleNode != null && !string.IsNullOrEmpty(titleNode.Value)) novelChapter.Title = titleNode.Value;


                if (pNodes != null)
                {
                    foreach (XElement pNode in pNodes) novelChapter.Append(pNode.Value);
                }

                if (novelChapter.Title == @"序章") result.Prologue = novelChapter.Content;
                else result.AddChapter(novelChapter);
            });

            return result;

        }

        private static async Task<Novel> LoadNovelFromEpubAsync(ZipArchive zipArchive,CancellationToken token = default)
        {
            Novel result = new Novel();
            XDocument xml;
            XmlNamespaceManager namespaceManager;
            Dictionary<string, ZipArchiveEntry> zipArchiveEntrys = zipArchive.Entries.Cast<ZipArchiveEntry>().ToDictionary(it => it.FullName.ToLower(), it => it);


            //container.xml
            string containerXmlFileName = @"META-INF/container.xml".ToLower();
            if (!zipArchiveEntrys.ContainsKey(containerXmlFileName)) throw new Exception(@"压缩文件包中找不到文件“container.xml”！");
            xml = await XDocument.LoadAsync(zipArchiveEntrys[containerXmlFileName].Open(),LoadOptions.None, token);
            namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace(@"ns", s_containerXmlXmlNs.NamespaceName);
            string? contentOpfFileName = xml.XPathSelectElement(@"/ns:container/ns:rootfiles/ns:rootfile", namespaceManager)?.Attribute(@"full-path")?.Value.ToLower() ?? null;
            if (contentOpfFileName is null || !zipArchiveEntrys.ContainsKey(contentOpfFileName)) throw new Exception(@"压缩文件包中找不到文件“content.opf”！");


            //content.opf
            List<string> herfs = new List<string>();
            xml =await XDocument.LoadAsync(zipArchiveEntrys[contentOpfFileName].Open(),LoadOptions.None, token);

            namespaceManager = new XmlNamespaceManager(new NameTable());
            namespaceManager.AddNamespace(@"ns", s_contentOpfXmlNs.NamespaceName);
            namespaceManager.AddNamespace(@"dc", s_contentOpfXmlDcNs.NamespaceName);

            string title = xml.XPathSelectElement(@"/ns:package/ns:metadata/dc:title", namespaceManager)?.Value ?? string.Empty;
            if (!string.IsNullOrEmpty(result.Title)) result.Title = title;

            IEnumerable<XElement> itemNodes = xml.XPathSelectElements(@"/ns:package/ns:manifest/ns:item", namespaceManager);

            itemNodes.ToList().ForEach(itemNode =>
            {
                XAttribute? mediaTypeAttribute = itemNode.Attribute(@"media-type");
                if (mediaTypeAttribute is null || mediaTypeAttribute.Value != @"application/xhtml+xml")
                {
                    return;
                }

                XAttribute? hrefAttribute = itemNode.Attribute(@"href");
                if (hrefAttribute is not null)
                {
                    string herfFileName = Path.Combine(Path.GetDirectoryName(contentOpfFileName) ?? string.Empty, hrefAttribute.Value.ToLower());
                    if (zipArchiveEntrys.ContainsKey(herfFileName)) herfs.Add(herfFileName);
                }
            });


            herfs.ForEach(herf =>
            {
                xml = XDocument.Load(zipArchiveEntrys[herf].Open());

                XElement? titleNode = xml.XPathSelectElement(@"/html/head/title");
                IEnumerable<XElement> pNodes = xml.XPathSelectElements(@"/html/body/p");
                NovelChapter novelChapter = new NovelChapter(@"第0000章-未知标题");

                if (titleNode != null && !string.IsNullOrEmpty(titleNode.Value)) novelChapter.Title = titleNode.Value;


                if (pNodes != null)
                {
                    foreach (XElement pNode in pNodes) novelChapter.Append(pNode.Value);
                }

                if (novelChapter.Title == @"序章") result.Prologue = novelChapter.Content;
                else result.AddChapter(novelChapter);
            });

            return result;

        }

        public void Clear()
        {
            this._prologue.Clear();
            this._chapters.ForEach(chapter => { chapter.Dispose(); });
            this._chapters.Clear();
        }


        public void Dispose()
        {
            this._prologue.Clear();
            this._chapters.ForEach(chapter => { chapter.Dispose(); });
            this._chapters.Clear();
        }

    }
}