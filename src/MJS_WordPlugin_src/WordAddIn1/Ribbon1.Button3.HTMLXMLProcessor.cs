using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        // HTML 文字列を処理する   
        public string ProcessHtmlString(string htmlStr, bool isTmpDot)
        {
            htmlStr = Regex.Replace(htmlStr, "\r\n", " ");
            htmlStr = Regex.Replace(htmlStr, "<meta[^>]*?>", "");
            htmlStr = Regex.Replace(htmlStr, "(<(?:input|br|img)[^>]*)>", "$1/>");
            htmlStr = Regex.Replace(htmlStr, "<span [^>]+>(?:&nbsp;)+ </span>", "　");
            htmlStr = Regex.Replace(htmlStr, "&nbsp;", ((char)160).ToString());
            htmlStr = Regex.Replace(htmlStr, "&copy;", ((char)169).ToString());
            while (Regex.IsMatch(htmlStr, @"(src\s*=\s*""[^""]*?)\\([^""]*?"")"))
                htmlStr = Regex.Replace(htmlStr, @"(src\s*=\s*""[^""]*?)\\([^""]*?"")", "$1/$2");

            while (Regex.IsMatch(htmlStr, @"(<[A-z]+[^>]* [A-z-]+=)([^""'][^ />]*)"))
            {
                htmlStr = Regex.Replace(htmlStr, @"(<[A-z]+[^>]* [A-z-]+=)([^""'][^ />]*)", @"$1""$2""");
            }

            if (isTmpDot)
            {
                htmlStr = Regex.Replace(htmlStr, @"src=""tmp\.files/", @"src=""pict/");
            }
            else
            {
                htmlStr = Regex.Replace(htmlStr, @"src=""tmp_files/", @"src=""pict/");
            }
            htmlStr = Regex.Replace(htmlStr, @"<a name=""_Toc\d+?""></a>", "");
            htmlStr = Regex.Replace(htmlStr, @"<span lang=""[^""]+?"">([^<]+?)</span>", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(<hr(?: [^/>]*)?)(>)", "$1/$2");
            htmlStr = Regex.Replace(htmlStr, @"z-index:-?\d{3,};", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(?<=<[^>]+?) style=['""]?[^'"" ]+['""]?( (?:[^>]*)style=['""]?[^'"" >/]+['""]?)", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(<p[^>]*?(?<!/)>)([^<]*)(</(?!p))", @"$1$2</p>$3");
            htmlStr = htmlStr.Replace("MJS--", "MJSTT");

            return htmlStr;
        }

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

            // objToc と objBody の初期化
            objToc = new XmlDocument();
            objToc.LoadXml(@"<result><item title=""" + docTitle + @"""></item></result>");

            objBody = new XmlDocument();
            objBody.LoadXml("<result></result>");
        }
    }
}
