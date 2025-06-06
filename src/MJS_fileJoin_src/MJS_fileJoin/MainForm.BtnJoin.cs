using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // webHelpを結合し、指定した出力ディレクトリに統合HTMLコンテンツを生成する
        private void btnJoin_Click(object sender, EventArgs e)
        {
            Cursor prevCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            // バリデーション
            if (!ValidateInput())
            {
                Cursor.Current = prevCursor;
                return;
            }

            StreamReader sr = null;
            StreamWriter sw = null;

            List<string> errorList = new List<string>();

            //テンプレート展開
            //System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            //using (Stream stream = assembly.GetManifestResourceStream("MJS_fileJoin.htmlTemplates.zip"))
            //{
            //    FileStream fs = File.Create(tbOutputDir.Text + "\\htmlTemplates.zip");
            //    stream.Seek(0, SeekOrigin.Begin);
            //    stream.CopyTo(fs);
            //    fs.Close();
            //}


            //if (Directory.Exists(tbOutputDir.Text + "\\htmlTemplates"))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\htmlTemplates", true);
            //}

            //ZipFile.ExtractToDirectory(tbOutputDir.Text + "\\htmlTemplates.zip", tbOutputDir.Text);

            //if (Directory.Exists(tbOutputDir.Text + "\\" + exportDir))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\" + exportDir, true);
            //}
            //Directory.Move(tbOutputDir.Text + "\\htmlTemplates", tbOutputDir.Text + "\\" + exportDir);

            //File.Delete(tbOutputDir.Text + "\\htmlTemplates.zip");

            //' Ver - 2023.16.08 - VyNL - ↑ - 追加'

            // 出力ディレクトリの準備
            // ここでexportDir変数に新しいフォルダ名が格納される
            PrepareOutputDirectory();

            //if (Directory.Exists(tbOutputDir.Text + "\\" + exportDir))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\" + exportDir, true);
            //}

            //Directory.CreateDirectory(tbOutputDir.Text + "\\" + exportDir);

            //CopyDirectory(lbHtmlList.Items[0].ToString(), tbOutputDir.Text + "\\" + exportDir);

            XmlDocument objToc = new XmlDocument();
            XmlNode objTocRoot = null;

            XmlDocument searchWords = new System.Xml.XmlDocument();
            searchWords.LoadXml("<div class='search'></div>");

            objToc.LoadXml(@"<result></result>");
            objTocRoot = objToc.DocumentElement;

            //各webHelpフォルダ処理

            // HTMLファイルリストの作成
            List<string> lsfiles = CreateHtmlFileList();

            //List<string> lsfiles = new List<string>();
            //foreach (string htmlDir in lbHtmlList.Items)
            //    foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
            //        lsfiles.Add(selRow["Column4"].ToString() + ".html");

            int picCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                picCount++;
                List<string> pics = new List<string>();
                foreach (string file in Directory.GetFiles(htmlDir + "\\pict", "*.*", SearchOption.AllDirectories))
                    pics.Add(Path.GetFileName(file));

                string outputDir = Path.Combine(tbOutputDir.Text, exportDir);

                //インデックスページ準備
                objTocRoot = PrepareIndexPage(htmlDir, outputDir, objTocRoot, objToc, tbChangeTitle, tbAddTop);

                foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
                {
                    if (!File.Exists(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html")))
                    {
                        errorList.Add("「" + Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html") + "」は存在しません。");
                        continue;
                    }

                    if (File.Exists(Path.Combine(outputDir, selRow["Column4"].ToString() + ".html")) && selRow["Column4"].ToString().Contains("00000"))
                    {

                        continue;
                    }

                    File.Copy(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html"), Path.Combine(outputDir, selRow["Column4"].ToString() + ".html"), true);

                    sr = new StreamReader(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html"));
                    string selHtml = sr.ReadToEnd();
                    sr.Close();

                    string[] coverKINs = { "EdgeTracker_logo50mm.png", "hyousi.png", "MJS_LOGO_255.gif" };
                    foreach (string coverKIN in coverKINs)
                    {
                        if (File.Exists(Path.Combine(htmlDir, "pict", coverKIN)) && !File.Exists(Path.Combine(outputDir, "pict", coverKIN)))
                            File.Copy(Path.Combine(htmlDir, "pict", coverKIN), Path.Combine(outputDir, "pict", coverKIN));
                    }

                    if (Regex.IsMatch(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\].+?"""))
                    {
                        //string dirName = Path.Combine("pict", selRow["Column4"].ToString().Substring(0, 3));
                        string dirName = "pict";
                        if (!Directory.Exists(Path.Combine(outputDir, dirName)))
                        {
                            Directory.CreateDirectory(Path.Combine(outputDir, dirName));
                        }

                        foreach (Match m in Regex.Matches(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\](.+?)"""))
                        {
                            if (!File.Exists(Path.Combine(outputDir, dirName, Path.GetFileNameWithoutExtension(m.Groups[1].Value) + "_" + picCount.ToString("00") + Path.GetExtension(m.Groups[1].Value))))
                            {
                                File.Copy(Path.Combine(htmlDir, "pict", m.Groups[1].Value), Path.Combine(outputDir, dirName, Path.GetFileNameWithoutExtension(m.Groups[1].Value) + "_" + picCount.ToString("00") + Path.GetExtension(m.Groups[1].Value)));
                            }
                        }

                        selHtml = Regex.Replace(selHtml, @"(<img(?: [^ />]+)* src="")pict[/\\](.+?)(\.\w+"")", "$1" + dirName + "/$2_" + picCount.ToString("00") + "$3");

                        //selHtml = Regex.Replace(selHtml, @"(<img(?: [^ />]+)* src="")pict/(.+?"")", "$1" + dirName + "/$2");
                    }

                    if (Regex.IsMatch(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>"))
                    {
                        string[] breadcrumbs = Regex.Replace(Regex.Match(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>").Groups[1].Value, "<.+?>", "").Split(new string[] { " &gt; " }, StringSplitOptions.None);
                        // get href by regex from selHtml
                        //var urls = Regex.Match(selHtml, "<a href\\s*=\\s*\"(?<url>.*?)\"").Groups["url"].Value;

                        Regex r = new Regex(@"<a.*?href=("")(?<href>.*?)(""|').*?>(?<value>.*?)</a>");
                        MatchCollection urls2 = r.Matches(selHtml);


                        for (int i = 0; i < breadcrumbs.Length; i++)
                        {
                            // get href by from urls2
                            string urls = "";
                            string title = "";
                            foreach (Match match in urls2)
                            {
                                title = match.Groups["value"].Value.ToString();
                                if (title == breadcrumbs[i])
                                {
                                    urls = match.Groups["href"].Value.ToString();
                                }
                                else
                                {
                                    urls = "";
                                }
                            }


                            string itemNodeLevel = "";
                            for (int j = 0; j <= i; j++)
                            {
                                itemNodeLevel += "/item[last()]";
                            }

                            if (objTocRoot.SelectSingleNode("." + itemNodeLevel + "[@title='" + breadcrumbs[i] + "']") == null)
                            {
                                itemNodeLevel = "";
                                for (int j = 0; j < i; j++)
                                {
                                    itemNodeLevel += "/item[last()]";
                                }
                                XmlElement addItem = (XmlElement)objTocRoot.SelectSingleNode("." + itemNodeLevel).AppendChild(objToc.CreateElement("item"));
                                addItem.SetAttribute("title", breadcrumbs[i]);

                                if (i == (breadcrumbs.Length - 1))
                                {
                                    /*if (selRow.Table.Columns["Column5"] != null && !String.IsNullOrEmpty(selRow["Column5"].ToString()))
                                    {
                                        //addItem.SetAttribute("href", selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + '#' + selRow["Column4"].ToString());
                                        addItem.SetAttribute("href", "./" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + ".html" + "#" + selRow["Column4"].ToString());
                                    }
                                    else
                                    {*/
                                    // get the href from current file
                                    if (urls != ""
                                        && urls.Contains("http") == false
                                        && urls.Contains(".html") == true
                                        && urls.Contains("#") == true)

                                    {
                                        addItem.SetAttribute("href", urls.Replace(".html", "").Replace("./", ""));
                                    }
                                    else
                                    {
                                        addItem.SetAttribute("href", selRow["Column4"].ToString());
                                    }
                                    // }

                                    XmlElement breadcrumbDisplay = objToc.CreateElement("div");
                                    string breadcrumb = "";
                                    string tocId = "";

                                    foreach (XmlElement objTocItem in addItem.SelectNodes("ancestor-or-self::item"))
                                    {
                                        if (breadcrumb != "")
                                        {
                                            breadcrumb += " > ";
                                            breadcrumbDisplay.AppendChild(objToc.CreateTextNode(" > "));
                                        }
                                        breadcrumb += ((XmlElement)objTocItem).GetAttribute("title");

                                        if (objTocItem.SelectSingleNode("@href") != null)
                                        {
                                            breadcrumbDisplay.AppendChild(objToc.CreateElement("a"));
                                            string href = "./" + ((XmlElement)objTocItem).GetAttribute("href") + ".html";
                                            if (((XmlElement)objTocItem).GetAttribute("href").Contains(".html"))
                                            {
                                                href = ((XmlElement)objTocItem).GetAttribute("href");
                                            }
                                            ((XmlElement)breadcrumbDisplay.LastChild).SetAttribute("href", href);
                                            breadcrumbDisplay.LastChild.InnerText = ((XmlElement)objTocItem).GetAttribute("title");
                                        }
                                        else
                                        {
                                            breadcrumbDisplay.AppendChild(objToc.CreateTextNode(((XmlElement)objTocItem).GetAttribute("title")));
                                        }

                                    }
                                    selHtml = Regex.Replace(selHtml, @"(?<=<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">)(?:.(?!</div>))+.(?=</div>)", breadcrumbDisplay.InnerXml);
                                    selHtml = Regex.Replace(selHtml, @"(?<=<meta name=""topic-breadcrumbs"" content="")[^""]*(?="")", breadcrumb);


                                    searchWords.DocumentElement.AppendChild(searchWords.CreateElement("div"));
                                    ((System.Xml.XmlElement)searchWords.DocumentElement.LastChild).SetAttribute("id", selRow["Column4"].ToString());
                                    string bodyStr = Regex.Replace(Regex.Replace(Regex.Replace(Regex.Replace(selHtml, "\r?\n", ""), "^.+<body[^>]*>(.+?)</body>.*$", "$1", RegexOptions.Multiline), @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">.+?</div>", ""), "<.+?>", "");

                                    string searchText = bodyStr.Replace("&", "&amp;").Replace("<", "&lt;");
                                    string displayText = searchText;
                                    if (searchText.Length >= 90)
                                    {
                                        displayText = displayText.Substring(0, 90) + " ...";
                                    }

                                    string[] wide = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン" };
                                    string[] narrow = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" };

                                    for (int p = 0; p < wide.Length; p++)
                                    {
                                        searchText = Regex.Replace(searchText, wide[p], narrow[p]);
                                    }
                                    searchText = searchText.ToLower();
                                    /*if (selRow.Table.Columns["Column5"] != null && !String.IsNullOrEmpty(selRow["Column5"].ToString()))
                                    {
                                        string searchHref_5 = selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "#" + selRow["Column4"].ToString();
                                        if(searchHref_5.Contains(".html") == false)
                                        {
                                            searchHref_5 = selRow["Column4"].ToString() + ".html" + "#" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "");
                                        }
                                        *//*searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>" 
                                            + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>" 
                                            + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;") 
                                            + "</div><div class='displayText'>" + displayText 
                                            + "</div><div class='search_word'>" + searchText + "</div>";*//*
                                        string textSearchWords = "<div class='search_breadcrumbs'>";
                                        textSearchWords += breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>";

                                        // check SelectSingleNode is null or not
                                        if (objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']") != null){
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }
                                        else if (objToc.SelectSingleNode(".//item[@href = '" + selRow["Column4"].ToString() + "']") != null)
                                        {
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + selRow["Column4"].ToString() + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }
                                        else if (objToc.SelectSingleNode(".//item[@href = '" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "']") != null)
                                        {
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }

                                        textSearchWords += "</div><div class='displayText'>" + displayText + "</div><div class='search_word'>" + searchText + "</div>";
                                        searchWords.DocumentElement.LastChild.InnerXml = textSearchWords;
                                    }
                                    else
                                    {*/
                                    string searchHref = selRow["Column4"].ToString();
                                    if (urls != "" && urls.Contains("http") == false && urls.Contains(".html") == true && urls.Contains("#") == true)
                                    {
                                        searchHref = urls.Replace(".html", "").Replace("./", "");
                                    }
                                    searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>"
                                        + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>"
                                        + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;")
                                        + "</div><div class='displayText'>" + displayText
                                        + "</div><div class='search_word'>" + searchText + "</div>";
                                    //}


                                }
                            }
                        }
                    }

                    MatchCollection mc = Regex.Matches(selHtml, @"(?<=<a href="")(?!\./|http)(?:[^""]*?/)+([^""]*?)(?="")", RegexOptions.Singleline);
                    foreach (Match m in mc)
                    {
                        string[] splitText = m.Groups[1].Value.Split('#');
                        // check if the file is in the list
                        if (lsfiles.Contains(splitText[0]))
                            if (m.Groups[1].Value.Contains("html") == true)
                            {
                                selHtml = selHtml.Replace(m.Value, "./" + m.Groups[1].Value);
                            }
                            else
                            {
                                selHtml = selHtml.Replace(m.Value, "./" + m.Groups[1].Value + "html");
                            }
                    }

                    //                    selHtml = Regex.Replace(selHtml, @"(?<=<a href="")(?!\./|http)(?:[^""]*?/)+([^""]*?)(?="")", "./$1");
                    sw = new StreamWriter(Path.Combine(outputDir, selRow["Column4"].ToString() + ".html"), false, Encoding.UTF8);
                    sw.Write(selHtml);
                    sw.Close();
                }
            }

            //全文検索ファイル出力
            sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, "search.js"), false, Encoding.UTF8);
            //            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2"), @"(?<=>)([^<]*?)'([^<]*?)(?=<)", "$1&apos;$2")));
            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;")));
            sw.Close();
            //Dictionary<string, string> mergeScript = new Dictionary<string, string>();
            foreach (XmlElement tocItem in objToc.SelectNodes(".//item[boolean(@href)]"))
            {
                if (tocItem.GetAttribute("href").Contains("#"))
                {
                    string[] parts = tocItem.GetAttribute("href").Split('#');

                    if (parts.Length >= 2)
                    {
                        string result = parts[1];
                        sr = new StreamReader(Path.Combine(tbOutputDir.Text, exportDir, result + ".html"));

                    }

                }
                else
                {
                    sr = new StreamReader(Path.Combine(tbOutputDir.Text, exportDir, tocItem.GetAttribute("href") + ".html"));
                }
                string selHtml = sr.ReadToEnd();
                sr.Close();

                string tocId = "";
                foreach (XmlElement objTocItem in tocItem.SelectNodes("ancestor-or-self::item"))
                {
                    if (tocId != "")
                    {
                        tocId += ".";
                    }
                    int precedingItemCount = objTocItem.SelectNodes("preceding-sibling::item[boolean(item)]|self::item[boolean(item)]").Count;
                    tocId += precedingItemCount.ToString();
                    if (objTocItem.SelectSingleNode("item") == null)
                    {
                        tocId += "_";
                        tocId += (objTocItem.SelectNodes("preceding-sibling::item[not(boolean(item)) and (count(preceding-sibling::item[boolean(item)]) = " + precedingItemCount + ")]").Count + 1).ToString();
                    }
                }

                selHtml = Regex.Replace(selHtml, @"(?<=gTopicId[\s]*=[\s]*"")[^""]*(?="")", tocId);
                if (tocItem.GetAttribute("href").Contains("#"))
                {
                    string[] parts = tocItem.GetAttribute("href").Split('#');

                    if (parts.Length >= 2)
                    {
                        string result = parts[1];
                        sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, result + ".html"), false, Encoding.UTF8);
                    }

                }
                else
                {
                    sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, tocItem.GetAttribute("href") + ".html"), false, Encoding.UTF8);
                }

                //string pattern = @"mergePage = {(.*?)};";
                //Match match = Regex.Match(selHtml, pattern, RegexOptions.Singleline);

                //if (match.Success)
                //{
                //    string mergePageData = match.Groups[1].Value;

                //    // Extract key-value pairs from mergePageData
                //    pattern = @"(\w+):'(\w+)'";
                //    MatchCollection matches = Regex.Matches(mergePageData, pattern);

                //    // Output the extracted key-value pairs
                //    foreach (Match m in matches)
                //    {
                //        string key = m.Groups[1].Value;
                //        string value = m.Groups[2].Value;
                //        if (!String.IsNullOrEmpty(key) && !String.IsNullOrEmpty(key)&& !mergeScript.Any(x => x.Key == key && x.Value == value))
                //            mergeScript.Add(key, value);
                //    }
                //}
                sw.Write(selHtml);
                sw.Close();
            }

            //目次出力
            createToc(objToc.DocumentElement);

            // chbListOutputがチェックされている場合にjoinList.xmlを出力する
            OutputJoinListXml();

            

            //書誌情報ファイルのマージ
            mergeHeaderFile();

            Cursor.Current = prevCursor;

            DialogResult selectMess = MessageBox.Show(tbOutputDir.Text + "\\" + exportDir + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？", "HTML出力成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectMess == DialogResult.Yes)
            {
                try
                {
                    Process.Start(tbOutputDir.Text + "\\" + exportDir + @"\index.html");
                }
                catch
                {
                    MessageBox.Show("HTMLの出力に失敗しました。", "HTML出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (checkBox2.Checked)
            {
                tabControl1.SelectedIndex = 1;
                listBox2.Items.Clear();
                listBox2.Items.Add(tbOutputDir.Text + "\\" + exportDir);
                button12.PerformClick();
            }
        }
    }
}
