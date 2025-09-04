using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void RemoveSearchBlockByTitle(string searchTitleText, string rootPath, string exportDir)
        {
            string searchJsPath = Path.Combine(rootPath, exportDir, "search.js");
            if (!File.Exists(searchJsPath)) return;

            string content = File.ReadAllText(searchJsPath, Encoding.UTF8);

            // 改行も含めてマッチするように修正
            string pattern = @"<div\s+class=""search_title"">([\s\S]*?)</div>\s*<div\s+class=""displayText"">([\s\S]*?)</div>\s*<div\s+class=""search_word"">([\s\S]*?)</div>";

            var regex = new Regex(pattern, RegexOptions.Singleline);
            var matches = regex.Matches(content);

            foreach (Match match in matches)
            {
                // HTMLエスケープを解除してから正規化
                string titleInner = System.Net.WebUtility.HtmlDecode(match.Groups[1].Value.Trim())
                    .Replace("\r", "").Replace("\n", "").Replace("　", " ")
                    .Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                    .Replace("&quot;", "\"").Replace("&#39;", "'")
                    .Normalize();

                string searchTitleNormalized = System.Net.WebUtility.HtmlDecode(searchTitleText.Trim())
                    .Replace("\r", "").Replace("\n", "").Replace("　", " ")
                    .Replace("&amp;", "&").Replace("&lt;", "<").Replace("&gt;", ">")
                    .Replace("&quot;", "\"").Replace("&#39;", "'")
                    .Normalize();

                // 大文字小文字を無視して比較
                if (string.Equals(titleInner, searchTitleNormalized, StringComparison.OrdinalIgnoreCase))
                {
                    content = content.Replace(match.Value, "");
                }
            }

            File.WriteAllText(searchJsPath, content, Encoding.UTF8);
        }
    }
}
