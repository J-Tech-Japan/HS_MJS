using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 目次ファイルの作成
        public void ExportTocAsJsFiles(XmlDocument objToc, string rootPath, string exportDir, Dictionary<string, string> mergeScript)
        {
            string htmlToc = "";
            string htmlToc1 = "";
            string htmlToc2 = "";
            string htmlToc3 = "";

            foreach (XmlNode toc in objToc.SelectNodes("/result/item"))
            {
                htmlToc = $@"{{""type"":""book"",""name"":""{CleanTitle(((XmlElement)toc).GetAttribute("title"))}"",""key"":""toc1""}}";
                htmlToc1 = "";

                foreach (XmlNode toc1 in toc.SelectNodes("item"))
                {
                    if (htmlToc1 != "") htmlToc1 += ",";
                    htmlToc1 += $@"{{""type"":""{(toc1.SelectNodes("item").Count != 0 ? "book" : "item")}"",""name"":""{CleanTitle(((XmlElement)toc1).GetAttribute("title"))}"",";

                    if (toc1.SelectNodes("item").Count != 0)
                        htmlToc1 += $@"""key"":""toc{toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2}"",";

                    if (!string.IsNullOrEmpty(((XmlElement)toc1).GetAttribute("href")))
                        htmlToc1 += $@"""url"":""{makeHrefWithMerge(mergeScript, ((XmlElement)toc1).GetAttribute("href"))}""";

                    htmlToc1 = htmlToc1.TrimEnd(',') + "}";

                    htmlToc2 = "";
                    foreach (XmlNode toc2 in toc1.SelectNodes("item"))
                    {
                        if (htmlToc2 != "") htmlToc2 += ",";
                        htmlToc2 += $@"{{""type"":""{(toc2.SelectNodes("item").Count != 0 ? "book" : "item")}"",""name"":""{CleanTitle(((XmlElement)toc2).GetAttribute("title"))}"",";

                        if (toc2.SelectNodes("item").Count != 0)
                            htmlToc2 += $@"""key"":""toc{toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3}"",";

                        if (!string.IsNullOrEmpty(((XmlElement)toc2).GetAttribute("href")))
                            htmlToc2 += $@"""url"":""{makeHrefWithMerge(mergeScript, ((XmlElement)toc2).GetAttribute("href"))}""";

                        htmlToc2 = htmlToc2.TrimEnd(',') + "}";

                        htmlToc3 = "";
                        foreach (XmlNode toc3 in toc2.SelectNodes("item"))
                        {
                            if (htmlToc3 != "") htmlToc3 += ",";
                            htmlToc3 += $@"{{""type"":""item"",""name"":""{CleanTitle(((XmlElement)toc3).GetAttribute("title"))}"",""url"":""{makeHrefWithMerge(mergeScript, ((XmlElement)toc3).GetAttribute("href"))}""}}";
                        }

                        if (htmlToc3 != "")
                        {
                            string fileName = $"toc{toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3}.new.js";
                            WriteTocJsFile(rootPath, exportDir, fileName, htmlToc3);
                        }
                    }

                    if (htmlToc2 != "")
                    {
                        string fileName = $"toc{toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2}.new.js";
                        WriteTocJsFile(rootPath, exportDir, fileName, htmlToc2);
                    }
                }

                WriteTocJsFile(rootPath, exportDir, "toc1.new.js", htmlToc1);
            }

            WriteTocJsFile(rootPath, exportDir, "toc.new.js", htmlToc);
        }

        //private void WriteTocJsFile(string rootPath, string exportDir, string fileName, string tocContent)
        //{
        //    string cleaned = Regex.Replace(tocContent, "(　[Ø²]|[Ø²]　)", "");
        //    string path = Path.Combine(rootPath, exportDir, "whxdata", fileName);
        //    using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
        //    {
        //        sw.WriteLine("(function() {");
        //        sw.WriteLine("var toc =  [" + cleaned + "];");
        //        sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
        //        sw.WriteLine("})();");
        //    }
        //}

        //// タイトル取得時に除去
        //private string CleanTitle(string title)
        //{
        //    return Regex.Replace(title, "(　[Ø²]|[Ø²]　)", "");
        //}

        // タイトル取得時に除去
        private string CleanTitle(string title)
        {
            // removeSymbolsの各記号の前後に全角スペースがある場合も含めて除去
            string pattern = string.Join("|", removeSymbols.Select(s => $"(　?{Regex.Escape(s.ToString())}　?|　?{Regex.Escape(s.ToString())})"));
            return Regex.Replace(title, pattern, "");
        }

        private void WriteTocJsFile(string rootPath, string exportDir, string fileName, string tocContent)
        {
            // removeSymbolsの各記号の前後に全角スペースがある場合も含めて除去
            string pattern = string.Join("|", removeSymbols.Select(s => $"(　?{Regex.Escape(s.ToString())}　?|　?{Regex.Escape(s.ToString())})"));
            string cleaned = Regex.Replace(tocContent, pattern, "");
            string path = Path.Combine(rootPath, exportDir, "whxdata", fileName);
            using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                sw.WriteLine("(function() {");
                sw.WriteLine("var toc =  [" + cleaned + "];");
                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                sw.WriteLine("})();");
            }
        }
    }
}