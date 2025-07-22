using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 指定テキストと <div class="search_title">...</div> の中身が一致した場合、
        // <div class="search_title">...</div> タグ全体と、
        // 直後の <div class="displayText">...</div><div class="search_word">...</div> も削除
        private void RemoveSearchBlockByTitle(string searchTitleText, string rootPath, string exportDir)
        {
            string searchJsPath = Path.Combine(rootPath, exportDir, "search.js");
            if (!File.Exists(searchJsPath)) return;

            string content = File.ReadAllText(searchJsPath, Encoding.UTF8);

            string pattern = @"<div\s+class=""search_title"">(.*?)</div>\s*<div\s+class=""displayText"">(.*?)</div>\s*<div\s+class=""search_word"">(.*?)</div>";
            var regex = new Regex(pattern, RegexOptions.Singleline);
            var matches = regex.Matches(content);

            foreach (Match match in matches)
            {
                string titleInner = match.Groups[1].Value.Trim();
                if (titleInner == searchTitleText)
                {
                    // 一致したブロック全体を削除
                    content = content.Replace(match.Value, "");
                }
            }

            File.WriteAllText(searchJsPath, content, Encoding.UTF8);
        }

        // 指定されたスタイル名の見出しを取得
        private List<string> GetHeadingsByStyles(List<string> styleNames)
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            List<string> headings = new List<string>();

            foreach (Paragraph para in activeDocument.Paragraphs)
            {
                string styleName = para.get_Style().NameLocal;
                if (styleNames.Contains(styleName))
                {
                    string text = para.Range.Text.Trim();
                    if (!string.IsNullOrEmpty(text))
                    {
                        headings.Add(text);
                    }
                }
            }

            return headings;
        }
    }
}
