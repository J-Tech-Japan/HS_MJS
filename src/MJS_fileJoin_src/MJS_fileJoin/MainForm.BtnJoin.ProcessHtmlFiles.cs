// MainForm.BtnJoin.ProcessHtmlFiles.cs

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void ProcessHtmlFiles(
            string htmlDir,
            string outputDir,
            int picCount,
            List<string> lsfiles,
            XmlNode objTocRoot,
            XmlDocument objToc,
            XmlDocument searchWords,
            List<string> errorList)
        {
            foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
            {
                string htmlFile = Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html");
                string outFile = Path.Combine(outputDir, selRow["Column4"].ToString() + ".html");

                if (!File.Exists(htmlFile))
                {
                    errorList.Add("「" + htmlFile + "」は存在しません。");
                    continue;
                }

                if (File.Exists(outFile) && selRow["Column4"].ToString().Contains("00000"))
                {
                    continue;
                }

                string selHtml = CopyHtmlAndImages(htmlFile, outFile, htmlDir, outputDir, picCount);
                selHtml = GenerateBreadcrumbsAndToc(selHtml, selRow, objTocRoot, objToc, searchWords);
                selHtml = FixRelativeLinks(selHtml, lsfiles);

                using (var sw = new StreamWriter(outFile, false, Encoding.UTF8))
                {
                    sw.Write(selHtml);
                }
            }
        }

        // HTMLファイルと画像のコピー、画像リンクの修正
        private string CopyHtmlAndImages(string htmlFile, string outFile, string htmlDir, string outputDir, int picCount)
        {
            File.Copy(htmlFile, outFile, true);
            string selHtml;
            using (var sr = new StreamReader(htmlFile))
            {
                selHtml = sr.ReadToEnd();
            }

            string[] coverKINs = { "EdgeTracker_logo50mm.png", "hyousi.png", "MJS_LOGO_255.gif" };
            foreach (string coverKIN in coverKINs)
            {
                string src = Path.Combine(htmlDir, "pict", coverKIN);
                string dst = Path.Combine(outputDir, "pict", coverKIN);
                if (File.Exists(src) && !File.Exists(dst))
                    File.Copy(src, dst);
            }

            if (Regex.IsMatch(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\].+?"""))
            {
                string dirName = "pict";
                string pictDir = Path.Combine(outputDir, dirName);
                if (!Directory.Exists(pictDir))
                {
                    Directory.CreateDirectory(pictDir);
                }

                foreach (Match m in Regex.Matches(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\](.+?)"""))
                {
                    string src = Path.Combine(htmlDir, "pict", m.Groups[1].Value);
                    string dst = Path.Combine(pictDir, Path.GetFileNameWithoutExtension(m.Groups[1].Value) + "_" + picCount.ToString("00") + Path.GetExtension(m.Groups[1].Value));
                    if (!File.Exists(dst))
                    {
                        File.Copy(src, dst);
                    }
                }

                selHtml = Regex.Replace(selHtml, @"(<img(?: [^ />]+)* src="")pict[/\\](.+?)(\.\w+"")", "$1" + dirName + "/$2_" + picCount.ToString("00") + "$3");
            }
            return selHtml;
        }

        // パンくずリスト・目次・全文検索用データ生成
        private string GenerateBreadcrumbsAndToc(string selHtml, DataRow selRow, XmlNode objTocRoot, XmlDocument objToc, XmlDocument searchWords)
        {
            if (!Regex.IsMatch(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>"))
                return selHtml;

            string[] breadcrumbs = Regex.Replace(
                Regex.Match(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>").Groups[1].Value,
                "<.+?>", "").Split(new string[] { " &gt; " }, StringSplitOptions.None);

            Regex r = new Regex(@"<a.*?href=("")(?<href>.*?)(""|').*?>(?<value>.*?)</a>");
            MatchCollection urls2 = r.Matches(selHtml);

            for (int i = 0; i < breadcrumbs.Length; i++)
            {
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
                        if (urls != "" && urls.Contains("http") == false && urls.Contains(".html") == true && urls.Contains("#") == true)
                        {
                            addItem.SetAttribute("href", urls.Replace(".html", "").Replace("./", ""));
                        }
                        else
                        {
                            addItem.SetAttribute("href", selRow["Column4"].ToString());
                        }

                        XmlElement breadcrumbDisplay = objToc.CreateElement("div");
                        string breadcrumb = "";
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
                        string bodyStr = Regex.Replace(
                            Regex.Replace(
                                Regex.Replace(
                                    Regex.Replace(selHtml, "\r?\n", ""),
                                    "^.+<body[^>]*>(.+?)</body>.*$", "$1", RegexOptions.Multiline),
                                @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">.+?</div>", ""),
                            "<.+?>", "");

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
                    }
                }
            }
            return selHtml;
        }

        // 相対リンクの修正
        private string FixRelativeLinks(string selHtml, List<string> lsfiles)
        {
            MatchCollection mc = Regex.Matches(selHtml, @"(?<=<a href="")(?!\./|http)(?:[^""]*?/)+([^""]*?)(?="")", RegexOptions.Singleline);
            foreach (Match m in mc)
            {
                string[] splitText = m.Groups[1].Value.Split('#');
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
            return selHtml;
        }
    }
}
