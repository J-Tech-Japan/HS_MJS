// GenerateHTMLButton.HtmlTemplate1.cs

using System.Collections.Generic;
using System.IO;
using System.Text;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private static string BuildHtmlTemplate1(Dictionary<string, string[]> title4Collection, Dictionary<string, string> mergeScript, string rootPath, string exportDir)
        {
            // exportDirに展開されたhtmlTemplate1Base.htmlのパスを取得
            string htmlTemplate1BasePath = Path.Combine(rootPath, exportDir, "htmlTemplate1Base.html");

            string htmlTemplate1 = "";

            // htmlTemplate1Base.htmlファイルが存在する場合はそれを読み込み
            if (File.Exists(htmlTemplate1BasePath))
            {
                htmlTemplate1 = File.ReadAllText(htmlTemplate1BasePath, Encoding.UTF8);

                // 改行文字を\nに統一
                htmlTemplate1 = Utils.NormalizeLineEndings(htmlTemplate1);

                // プレースホルダーを実際の値に置換
                htmlTemplate1 = htmlTemplate1.Replace("{{TOPIC_BREADCRUMBS_PLACEHOLDER}}", @" <meta name=""topic-breadcrumbs"" content="""" />");
                htmlTemplate1 = htmlTemplate1.Replace("{{BODY_STYLE_PLACEHOLDER}}", @"style=""text-justify-trim: punctuation;""");
                htmlTemplate1 = htmlTemplate1.Replace("{{TOPIC_ID_PLACEHOLDER}}", "gTopicId = \"♪\";");
                htmlTemplate1 = htmlTemplate1.Replace("{{REFPAGE_DATA_PLACEHOLDER}}", $"refPage = {{{BuildRefPageData(title4Collection)}}};");
                htmlTemplate1 = htmlTemplate1.Replace("{{MERGEPAGE_DATA_PLACEHOLDER}}", $"mergePage = {{{BuildMergePageData(mergeScript)}}};");
            }

            return htmlTemplate1;
        }

        private static string BuildRefPageData(Dictionary<string, string[]> title4Collection)
        {
            var refPageData = new StringBuilder();
            foreach (var item in title4Collection)
            {
                refPageData.Append($"{item.Key}:['{item.Value[0]}','{item.Value[1]}'],");
            }
            return refPageData.ToString();
        }
        
        private static string BuildMergePageData(Dictionary<string, string> mergeScript)
        {
            var mergePageData = new StringBuilder();
            foreach (var item in mergeScript)
            {
                mergePageData.Append($"{item.Value.Split(new char[] { '♯' })[0]}:'{item.Key.Split(new char[] { '♯' })[0]}',");
            }
            return mergePageData.ToString();
        }
    }
}
