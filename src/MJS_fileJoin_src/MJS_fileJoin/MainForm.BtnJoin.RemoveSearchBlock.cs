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
                // 改行・空白・全角半角を除去して比較
                string titleInner = match.Groups[1].Value.Trim()
                    .Replace("\r", "").Replace("\n", "").Replace("　", " ").Normalize();

                string searchTitleNormalized = searchTitleText.Trim()
                    .Replace("\r", "").Replace("\n", "").Replace("　", " ").Normalize();

                if (titleInner == searchTitleNormalized)
                {
                    content = content.Replace(match.Value, "");
                }
            }

            File.WriteAllText(searchJsPath, content, Encoding.UTF8);
        }

        /// <summary>
        /// 統合時に新たに追加されたタイトルを表示する（デバッグ用）
        /// </summary>
        /// <param name="newTitles">新たに追加されたタイトルのリスト</param>
        private void DisplayNewTitles(List<string> newTitles)
        {
            if (newTitles.Count == 0)
            {
                MessageBox.Show("統合時に新たに追加されたタイトルはありません。", "新規タイトル", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== 統合時に新たに追加されたタイトル ===");
            sb.AppendLine();
            sb.AppendLine($"【新規追加タイトル数】: {newTitles.Count}件");
            sb.AppendLine();

            for (int i = 0; i < newTitles.Count; i++)
            {
                sb.AppendLine($"【{i + 1:D3}】 {newTitles[i]}");
            }

            MessageBox.Show(sb.ToString(), "新規タイトル", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 結合元の各フォルダのsearch.jsファイルから<div class="search_title">の内容を抽出して合わせたリストを返す
        /// </summary>
        /// <returns>結合元フォルダのsearch_titleの内容を合わせたリスト</returns>
        private List<string> ExtractSourceSearchTitles()
        {
            var allSourceTitles = new List<string>();

            try
            {
                // 結合元の各フォルダのsearch.jsファイルの内容を抽出
                foreach (string htmlDir in lbHtmlList.Items)
                {
                    string sourceSearchJsPath = Path.Combine(htmlDir, "search.js");
                    var sourceTitles = ExtractSearchTitlesFromFile(sourceSearchJsPath);
                    allSourceTitles.AddRange(sourceTitles);
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合はログに記録（必要に応じて）
                // 現在は何もしない
            }

            return allSourceTitles;
        }

        /// <summary>
        /// 指定されたsearch.jsファイルから<div class="search_title">の内容を抽出してリストで返す
        /// </summary>
        /// <param name="searchJsFilePath">search.jsファイルのパス</param>
        /// <returns>抽出されたsearch_titleの内容のリスト</returns>
        private List<string> ExtractSearchTitlesFromFile(string searchJsFilePath)
        {
            var titles = new List<string>();

            try
            {
                if (!File.Exists(searchJsFilePath))
                {
                    return titles;
                }

                string fileContent = File.ReadAllText(searchJsFilePath, Encoding.UTF8);

                // search.jsファイルからsearchWordsの変数内容を抽出
                var searchWordsMatch = Regex.Match(fileContent, @"var searchWords = \$\('(.*?)'\);", RegexOptions.Singleline);
                if (!searchWordsMatch.Success)
                {
                    return titles;
                }

                string searchWordsContent = searchWordsMatch.Groups[1].Value;

                // <div class="search_title">の内容をすべて抽出
                var titleMatches = Regex.Matches(searchWordsContent, @"<div\s+class=['""]search_title['""]>(.*?)</div>", RegexOptions.Singleline);

                foreach (Match match in titleMatches)
                {
                    string titleContent = match.Groups[1].Value;
                    // HTMLエンティティをデコード
                    titleContent = titleContent.Replace("&amp;", "&")
                                             .Replace("&lt;", "<")
                                             .Replace("&gt;", ">")
                                             .Replace("&quot;", "\"")
                                             .Replace("&apos;", "'");

                    titles.Add(titleContent);
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合は空のリストを返す
            }

            return titles;
        }
    }
}
