using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private (XmlDocument objXml, XmlDocument objToc, XmlDocument objBody) LoadAndProcessXml(string htmlStr, string docTitle)
        {
            XmlDocument objXml = new XmlDocument();
            objXml.LoadXml(htmlStr);
            ProcessXmlDocuments(objXml, docTitle, out XmlDocument objToc, out XmlDocument objBody);
            return (objXml, objToc, objBody);
        }

        public void ProcessXmlDocuments(XmlDocument objXml, string docTitle, out XmlDocument objToc, out XmlDocument objBody)
        {
            RemoveImgAttributes(objXml);
            RemovePageBreaks(objXml);
            RemoveComments(objXml);
            NormalizeLinkText(objXml);
            CleanTocHeadings(objXml);
            InitializeTocAndBody(docTitle, out objToc, out objBody);
        }

        private void RemoveImgAttributes(XmlDocument objXml)
        {
            foreach (XmlElement imgNode in objXml.SelectNodes("//img"))
            {
                imgNode.RemoveAttribute("height");
                imgNode.RemoveAttribute("width");
            }
        }

        // ページ区切りを削除する
        private void RemovePageBreaks(XmlDocument objXml)
        {
            foreach (XmlElement pageBreak in objXml.SelectNodes("//span[(translate(., ' \u0020\u000a\u000d\u0009', '') = '') and (count(*) = 1) and boolean(br[@style = 'page-break-before:always'])]"))
            {
                pageBreak.ParentNode.RemoveChild(pageBreak);
            }
            foreach (XmlElement pageBreak in objXml.SelectNodes("//br[translate(@style, ' \u0020\u000a\u000d\u0009', '') = 'page-break-before:always']"))
            {
                pageBreak.ParentNode.RemoveChild(pageBreak);
            }
        }

        // コメントノードを削除する
        private void RemoveComments(XmlDocument objXml)
        {
            foreach (XmlElement comment in objXml.SelectNodes("//*[boolean(./*/@class[starts-with(., 'msocom')])]"))
            {
                comment.ParentNode.RemoveChild(comment);
            }
        }

        // リンクテキストを正規化する
        private void NormalizeLinkText(XmlDocument objXml)
        {
            foreach (XmlElement link in objXml.SelectNodes("//a[boolean(@href)]"))
            {
                if (link.InnerText.Contains("http")) continue;
                link.InnerText = Regex.Replace(link.InnerText, @"^(.*?)(?=[\s　](\d+\.\d+|[^\s|　]*?章))", "");
                link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:第[\d０-９]+章)*[\s　]+", "");
                link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:\d+\.)*\d+[\s　]+", "");
            }
        }

        // 目次の見出しを整理する
        private void CleanTocHeadings(XmlDocument objXml)
        {
            foreach (XmlElement toc in objXml.SelectNodes("//a[starts-with(@name, '_Toc')]"))
            {
                foreach (XmlElement childSpan in toc.SelectNodes(".//span[contains(@style, 'Wingdings')]"))
                    childSpan.ParentNode.RemoveChild(childSpan);
                foreach (XmlElement brotherSpan in toc.ParentNode.SelectNodes(".//span[contains(@style, 'Wingdings')]"))
                    brotherSpan.ParentNode.RemoveChild(brotherSpan);
                if (!string.IsNullOrEmpty(toc.InnerText))
                {
                    toc.InnerText = Regex.Replace(toc.InnerText, @"^[・･]\s*", "");
                }
            }
        }

        // 目次と本文の初期化
        private void InitializeTocAndBody(string docTitle, out XmlDocument objToc, out XmlDocument objBody)
        {
            objToc = new XmlDocument();
            objToc.LoadXml("<result><item title=\"" + docTitle + "\"></item></result>");
            objBody = new XmlDocument();
            objBody.LoadXml("<result></result>");
        }

        // 不要な XML ノードを整理する
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
