using System.Collections.Generic;
using System.IO;
using System.Text;
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
                htmlToc = @"{""type"":""book"",""name"":""" + ((XmlElement)toc).GetAttribute("title") + @""",""key"":""toc1""}";

                foreach (XmlNode toc1 in toc.SelectNodes("item"))
                {
                    if (htmlToc1 != "")
                    {
                        htmlToc1 = htmlToc1 + ",";
                    }

                    htmlToc1 = htmlToc1 + @"{""type"":""" + (toc1.SelectNodes("item").Count != 0 ? "book" : "item") + @""",""name"":""" + ((XmlElement)toc1).GetAttribute("title") + @"""";

                    if (toc1.SelectNodes("item").Count != 0)
                    {
                        htmlToc1 += @",""key"":""toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + @"""";
                    }

                    if (((XmlElement)toc1).GetAttribute("href") != "")
                    {
                        htmlToc1 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc1).GetAttribute("href")) + @"""";
                    }

                    htmlToc1 += "}";

                    foreach (XmlNode toc2 in toc1.SelectNodes("item"))
                    {
                        if (htmlToc2 != "")
                        {
                            htmlToc2 = htmlToc2 + ",";
                        }

                        htmlToc2 += @"{""type"":""" + (toc2.SelectNodes("item").Count != 0 ? "book" : "item") + @""",""name"":""" + ((XmlElement)toc2).GetAttribute("title") + @"""";

                        if (toc2.SelectNodes("item").Count != 0)
                        {
                            htmlToc2 += @",""key"":""toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + @"""";
                        }
                        if (((XmlElement)toc2).GetAttribute("href") != "")
                        {
                            htmlToc2 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc2).GetAttribute("href")) + @"""";
                        }

                        htmlToc2 += "}";

                        foreach (XmlNode toc3 in toc2.SelectNodes("item"))
                        {
                            if (htmlToc3 != "")
                            {
                                htmlToc3 += ",";
                            }

                            htmlToc3 += @"{""type"":""item"",""name"":""" + ((XmlElement)toc3).GetAttribute("title") + @""",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc3).GetAttribute("href")) + @"""}";
                        }

                        if (htmlToc3 != "")
                        {
                            using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + ".new.js", false, Encoding.UTF8))
                            {
                                sw.WriteLine("(function() {");
                                sw.WriteLine("var toc =  [" + htmlToc3 + "];");
                                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                                sw.WriteLine("})();");
                            }
                            htmlToc3 = "";
                        }
                    }

                    if (htmlToc2 != "")
                    {
                        using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + ".new.js", false, Encoding.UTF8))
                        {
                            sw.WriteLine("(function() {");
                            sw.WriteLine("var toc =  [" + htmlToc2 + "];");
                            sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                            sw.WriteLine("})();");
                        }
                        htmlToc2 = "";
                    }
                }

                using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc1.new.js", false, Encoding.UTF8))
                {
                    sw.WriteLine("(function() {");
                    sw.WriteLine("var toc =  [" + htmlToc1 + "];");
                    sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                    sw.WriteLine("})();");
                }
            }

            using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc.new.js", false, Encoding.UTF8))
            {
                sw.WriteLine("(function() {");
                sw.WriteLine("var toc =  [" + htmlToc + "];");
                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                sw.WriteLine("})();");
            }
        }
    }
}
