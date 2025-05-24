using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private string ReadAndProcessHtml(string tmpHtmlPath, bool isTmpDot)
        {
            string htmlStr;
            using (StreamReader sr = new StreamReader(tmpHtmlPath, Encoding.UTF8))
            {
                htmlStr = sr.ReadToEnd();
            }
            return ProcessHtmlString(htmlStr, isTmpDot);
        }

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

            htmlStr = isTmpDot
                ? Regex.Replace(htmlStr, @"src=""tmp\.files/", @"src=""pict/")
                : Regex.Replace(htmlStr, @"src=""tmp_files/", @"src=""pict/");

            htmlStr = Regex.Replace(htmlStr, @"<a name=""_Toc\d+?""></a>", "");
            htmlStr = Regex.Replace(htmlStr, @"<span lang=""[^""]+?"">([^<]+?)</span>", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(<hr(?: [^/>]*)?)(>)", "$1/$2");
            htmlStr = Regex.Replace(htmlStr, @"z-index:-?\d{3,};", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(?<=<[^>]+?) style=['""]?[^'"" ]+['""]?( (?:[^>]*)style=['""]?[^'"" >/]+['""]?)", "$1");
            htmlStr = Regex.Replace(htmlStr, @"(<p[^>]*?(?<!/)>)([^<]*)(</(?!p))", @"$1$2</p>$3");
            htmlStr = htmlStr.Replace("MJS--", "MJSTT");

            return htmlStr;
        }
    }

}