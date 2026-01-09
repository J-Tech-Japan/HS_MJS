using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MJS_fileJoin
{
    partial class MainForm
    {
        // 結合元フォルダのsearch.jsを読み込み、searchWordsを書き換えるメソッド
        private static string GetSearchJsWithReplacedWords(List<string> htmlDirs, string newSearchWordsXml)
        {
            // 結合元フォルダから最初に見つかったsearch.jsを使用
            foreach (string htmlDir in htmlDirs)
            {
                string searchJsPath = Path.Combine(htmlDir, "search.js");
                
                if (File.Exists(searchJsPath))
                {
                    try
                    {
                        string searchJsContent = File.ReadAllText(searchJsPath, Encoding.UTF8);
                        
                        // var searchWords = $('...'); の部分を新しい内容で置き換える
                        // パターン: var searchWords = $('...');\n から最初の行の終わりまで
                        string pattern = @"var\s+searchWords\s*=\s*\$\('.*?'\);?";
                        string replacement = $"var searchWords = $('{newSearchWordsXml}');";
                        
                        string result = Regex.Replace(searchJsContent, pattern, replacement, RegexOptions.Singleline);
                        
                        return result;
                    }
                    catch (Exception ex)
                    {
                        // 読み込みエラーの場合は次のフォルダを試す
                        System.Diagnostics.Trace.WriteLine($"search.js読み込みエラー ({searchJsPath}): {ex.Message}");
                    }
                }
            }

            // どの結合元フォルダにもsearch.jsが見つからない場合は、
            // 最小限のsearch.jsを生成して返す
            System.Diagnostics.Trace.WriteLine("search.jsが見つかりませんでした。最小限の内容を生成します。");
            return $"var searchWords = $('{newSearchWordsXml}');";
        }
    }
}