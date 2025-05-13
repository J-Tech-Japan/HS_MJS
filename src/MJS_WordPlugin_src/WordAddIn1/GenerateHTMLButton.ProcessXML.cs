using Microsoft.Office.Interop.Word;

using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        public void ProcessXmlDocuments(XmlDocument objXml, string docTitle, out XmlDocument objToc, out XmlDocument objBody)
        {
            // img ノードの height と width 属性を削除
            foreach (XmlElement imgNode in objXml.SelectNodes("//img"))
            {
                imgNode.RemoveAttribute("height");
                imgNode.RemoveAttribute("width");
            }

            // 不要なページ区切りを削除
            foreach (XmlElement pageBreak in objXml.SelectNodes("//span[(translate(., ' &#10;&#13;&#9;', '') = '') and (count(*) = 1) and boolean(br[@style = 'page-break-before:always'])]"))
            {
                pageBreak.ParentNode.RemoveChild(pageBreak);
            }
            foreach (XmlElement pageBreak in objXml.SelectNodes("//br[translate(@style, ' &#10;&#13;&#9;', '') = 'page-break-before:always']"))
            {
                pageBreak.ParentNode.RemoveChild(pageBreak);
            }

            // コメントを削除
            foreach (XmlElement comment in objXml.SelectNodes("//*[boolean(./*/@class[starts-with(., 'msocom')])]"))
            {
                comment.ParentNode.RemoveChild(comment);
            }

            // リンクのテキストを正規化
            foreach (XmlElement link in objXml.SelectNodes("//a[boolean(@href)]"))
            {
                if (link.InnerText.Contains("http")) continue;
                link.InnerText = Regex.Replace(link.InnerText, @"^(.*?)(?=[\s　](\d+\.\d+|[^\s|　]*?章))", "");
                link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:第[\d０-９]+章)*[\s　]+", "");
                link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:\d+\.)*\d+[\s　]+", "");
            }

            // 見出しで箇条書きタグを削除
            foreach (XmlElement toc in objXml.SelectNodes("//a[starts-with(@name, '_Toc')]"))
            {
                foreach (XmlElement childSpan in toc.SelectNodes(".//span[contains(@style, 'Wingdings')]"))
                    childSpan.ParentNode.RemoveChild(childSpan);
                foreach (XmlElement brotherSpan in toc.ParentNode.SelectNodes(".//span[contains(@style, 'Wingdings')]"))
                    brotherSpan.ParentNode.RemoveChild(brotherSpan);
            }

            // objToc と objBody の初期化
            objToc = new XmlDocument();
            objToc.LoadXml(@"<result><item title=""" + docTitle + @"""></item></result>");

            objBody = new XmlDocument();
            objBody.LoadXml("<result></result>");
        }

        // HTML セクションを解析し、目次 (objToc) と本文 (objBody) を構築する
        public void GenerateTocAndBody(
            XmlDocument objXml,
            XmlDocument objBody,
            XmlDocument objToc,
            string chapterSplitClass,
            Dictionary<string, string> styleName,
            string docid,
            string bookInfoDef,
            ref XmlNode objBodyCurrent,
            ref XmlNode objTocCurrent,
            loader load)
        {
            string lv1styleName = "";
            string lv2styleName = "";
            string lv3styleName = "";

            int lv1count = 0;
            int lv2count = 0;
            int lv3count = 0;

            bool chapterSplit = false;

            foreach (XmlElement sectionNode in objXml.SelectNodes("/html/body/div"))
            {
                objBodyCurrent = (XmlElement)objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));

                if (chapterSplit)
                {
                    chapterSplit = false;
                }

                if (sectionNode.SelectSingleNode(chapterSplitClass) != null)
                {
                    ((XmlElement)objBodyCurrent).SetAttribute("style", "width:714px");
                    lv1styleName = chapterSplitClass;
                    chapterSplit = true;
                }

                bool breakFlg = false;

                foreach (XmlNode childs in sectionNode.SelectNodes("*"))
                {
                    string thisStyleName = "";

                    if (childs.SelectSingleNode("@class") == null)
                    {
                        if (styleName.ContainsKey(childs.Name))
                        {
                            thisStyleName = styleName[childs.Name];
                        }
                    }
                    else
                    {
                        if (styleName.ContainsKey(childs.Name + "." + childs.SelectSingleNode("@class").InnerText))
                        {
                            thisStyleName = styleName[childs.Name + "." + childs.SelectSingleNode("@class").InnerText];
                        }
                    }

                    if ((thisStyleName == "") && (childs.SelectSingleNode("*[@class != '']") != null))
                    {
                        if (styleName.ContainsKey(childs.SelectSingleNode("*[@class != '']").Name + "." + childs.SelectSingleNode("*[@class != '']/@class").InnerText))
                        {
                            thisStyleName = styleName[childs.SelectSingleNode("*[@class != '']").Name + "." + childs.SelectSingleNode("*[@class != '']/@class").InnerText];
                        }
                    }
                    else if ((thisStyleName == "") && (childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']") != null))
                    {
                        if (styleName.ContainsKey(childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name))
                        {
                            thisStyleName = styleName[childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name];
                        }
                    }

                    if (childs.SelectSingleNode(".//text()[1]") != null)
                    {
                        if (Regex.IsMatch(childs.SelectSingleNode(".//text()[1]").InnerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                        {
                            breakFlg = true;
                            break;
                        }

                        if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[\d０-９](?:[^・用]+|)$") && Regex.IsMatch(childs.SelectSingleNode(".//text()[1]").InnerText, @"^(?:\d+\.)*\d+[\s　]+"))
                        {
                            childs.SelectSingleNode(".//text()[1]").InnerText = Regex.Replace(childs.SelectSingleNode(".//text()[1]").InnerText, @"^(?:\d+\.)*\d+[\s　]+", "");
                        }
                    }

                    string setid = "";
                    if (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1２2３3](?:[^・用]+|)$"))
                    {
                        if (childs.SelectSingleNode(".//a[starts-with(@name, '" + docid + bookInfoDef + "')]") != null)
                        {
                            //aaa
                            setid = ((XmlElement)childs.SelectSingleNode(".//a[starts-with(@name, '" + docid + bookInfoDef + "')]")).GetAttribute("name");
                        }
                        else
                        {
                            load.Visible = false;
                            MessageBox.Show(childs.InnerText + ":書誌情報ブックマークの設定が行われていません。");
                            load.Visible = true;
                        }
                    }


                    if (Regex.IsMatch(thisStyleName, "目[　 ]*次"))
                    {
                    }
                    else if (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル"))
                    {
                        lv1count++;
                        lv2styleName = "";
                        lv2count = 0;
                        lv3styleName = "";
                        lv2count = 0;

                        objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");
                        objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                        ((XmlElement)objTocCurrent).SetAttribute("title", Regex.Replace(childs.InnerText, @"^第[\d０-９]+章[　\s]*", ""));

                        ((XmlElement)objBodyCurrent).SetAttribute("id", setid);
                    }
                    else if (Regex.IsMatch(thisStyleName, "章[　 ]*扉"))
                    {
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$"))
                    {
                        if (!Regex.IsMatch(childs.InnerText, @"目\s*次\s*$"))
                        {
                            if ((lv1styleName == "") || (lv1styleName == thisStyleName) || Regex.IsMatch(lv1styleName, @"(見出し|Heading)\s*[２2]"))
                            {
                                lv1count++;
                                lv2styleName = "";
                                lv2count = 0;
                                lv3styleName = "";
                                lv3count = 0;

                                objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");

                                objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                ((XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                ((XmlElement)objTocCurrent).SetAttribute("href", setid);

                                lv1styleName = thisStyleName;
                            }
                            else
                            {
                                lv2count++;
                                lv3styleName = "";
                                lv3count = 0;

                                if ((objTocCurrent == null) || (objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]") == null))
                                {
                                }
                                else
                                {
                                    objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]");

                                    objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                    ((XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                    ((XmlElement)objTocCurrent).SetAttribute("href", setid);
                                }
                                lv2styleName = thisStyleName;

                            }
                            objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                            ((XmlElement)objBodyCurrent).SetAttribute("id", setid);

                            ((XmlElement)objBodyCurrent).AppendChild(objBody.CreateElement("p"));
                            ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading1");


                            foreach (XmlNode childItem in childs.ChildNodes)
                            {
                                InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                            }
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[２2](?![・用])"))
                    {
                        if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                        {
                            lv1count++;
                            lv2styleName = "";
                            lv2count = 0;
                            lv3styleName = "";
                            lv3count = 0;

                            objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");
                            objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                            ((XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                            ((XmlElement)objTocCurrent).SetAttribute("href", setid);
                        }
                        else
                        {
                            if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                            {
                                lv2count++;
                                lv3styleName = "";
                                lv3count = 0;

                                objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]");
                            }
                            else
                            {
                                lv3count++;

                                objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 2]");
                            }

                            objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                            ((XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                            ((XmlElement)objTocCurrent).SetAttribute("href", setid);
                        }

                        objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                        ((XmlElement)objBodyCurrent).SetAttribute("id", setid);

                        objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading1 NoPageBreak");

                        foreach (XmlNode childItem in childs.ChildNodes)
                        {
                            InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                        }

                        if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                        {
                            lv1styleName = thisStyleName;
                        }
                        else if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                        {
                            lv2styleName = thisStyleName;
                        }
                        else
                        {
                            lv3styleName = thisStyleName;
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                    {
                        objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading3");
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("id", Regex.Replace(setid, "^.*?♯(.*?)$", "$1"));

                        foreach (XmlNode childItem in childs.ChildNodes)
                        {
                            InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[４4](?![・用])"))
                    {
                        objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading4");
                        foreach (XmlNode childItem in childs.ChildNodes)
                        {
                            InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[５5]"))
                    {
                        objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading5");
                        foreach (XmlNode childItem in childs.ChildNodes)
                        {
                            InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                        }
                    }
                    else
                    {
                        if (objBodyCurrent.Name == "result")
                        {
                            objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                        }
                        InnerNode(styleName, objBodyCurrent, childs);
                    }
                }

                if (breakFlg) break;
            }
        }

        // 目次ファイルの作成
        public void GenerateTocFiles(XmlDocument objToc, string rootPath, string exportDir, Dictionary<string, string> mergeScript)
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

        public void CleanUpXmlNodes(XmlDocument objBody)
        {
            // lang 属性や name 属性を削除し、不要なノードを整理
            foreach (XmlElement langSpan in objBody.SelectNodes(".//span[boolean(@lang)]|.//a"))
            {
                langSpan.RemoveAttribute("lang");

                if (langSpan.Name == "a")
                {
                    langSpan.RemoveAttribute("name");
                }

                if (langSpan.Attributes.Count == 0)
                {
                    while (langSpan.ChildNodes.Count != 0)
                    {
                        langSpan.ParentNode.InsertBefore(langSpan.ChildNodes[0], langSpan);
                    }
                    langSpan.ParentNode.RemoveChild(langSpan);
                }
            }

            // 不要な <div> や <br> タグを削除
            while (objBody.SelectSingleNode("/result/div//*[((name() = 'div') or (name() = 'br')) and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ") != null)
            {
                XmlNode lineBreak = objBody.SelectSingleNode("/result/div//*[((name() = 'div') or (name() = 'br')) and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ");
                lineBreak.ParentNode.RemoveChild(lineBreak);
            }

            // 不要な <p> タグを削除
            while (objBody.SelectSingleNode("/result/div//*[not(img)][(name() = 'p') and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ") != null)
            {
                XmlNode lineBreak = objBody.SelectSingleNode("/result/div//*[not(img)][(name() = 'p') and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ");
                lineBreak.ParentNode.RemoveChild(lineBreak);
            }
        }

    }
}
