using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // XML から目次 (objToc) と本文 (objBody) を構築する
        public void BuildTocBodyFromXml(
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
                            setid = ((XmlElement)childs.SelectSingleNode(".//a[starts-with(@name, '" + docid + bookInfoDef + "')]")).GetAttribute("name");
                        }
                        else
                        {
                            load.Visible = false;
                            //MessageBox.Show(childs.InnerText + ":書誌情報ブックマークの設定が行われていません。");
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

                        // 先頭の「・」や「・　」を除去
                        if (childs.InnerText.StartsWith("・") || childs.InnerText.StartsWith("･"))
                        {
                            var firstTextNode = childs.SelectSingleNode(".//text()[1]");
                            if (firstTextNode != null)
                            {
                                firstTextNode.InnerText = Regex.Replace(firstTextNode.InnerText, @"^[・･]\s*", "");
                            }
                        }

                        foreach (XmlNode childItem in childs.ChildNodes)
                        {
                            InnerNode(styleName, objBodyCurrent.LastChild, childItem);
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[５5]"))
                    {
                        objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                        ((XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading5");

                        // 先頭の「・」や「･」を除去
                        if (childs.InnerText.StartsWith("・") || childs.InnerText.StartsWith("･"))
                        {
                            var firstTextNode = childs.SelectSingleNode(".//text()[1]");
                            if (firstTextNode != null)
                            {
                                firstTextNode.InnerText = Regex.Replace(firstTextNode.InnerText, @"^[・･]\s*", "");
                            }
                        }

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
    }
}
