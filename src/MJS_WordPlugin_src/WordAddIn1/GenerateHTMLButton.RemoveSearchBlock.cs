// GenerateHTMLButton.RemoveSearchBlock.cs

using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
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

        // 指定されたスタイル名の見出しを取得
        //private List<string> GetHeadingsByStyles(List<string> styleNames)
        //{
        //    var application = Globals.ThisAddIn.Application;
        //    var activeDocument = application.ActiveDocument;
        //    List<string> headings = new List<string>();

        //    foreach (Paragraph para in activeDocument.Paragraphs)
        //    {
        //        string styleName = para.get_Style().NameLocal;
        //        if (styleNames.Contains(styleName))
        //        {
        //            string text = para.Range.Text.Trim();
        //            if (!string.IsNullOrEmpty(text))
        //            {
        //                headings.Add(text);
        //            }
        //        }
        //    }

        //    return headings;
        //}

        // 指定されたスタイル名の見出し内にコメントがついている場合
        // 見出しのテキストと、見出しの配下の見出しテキストを取得
        private List<string> GetHeadingsWithComment(List<string> styleNames, string commentText)
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            List<string> headings = new List<string>();

            int? currentLevel = null;
            bool inBlock = false;

            foreach (Paragraph para in activeDocument.Paragraphs)
            {
                string styleName = para.get_Style().NameLocal;
                string headingText = para.Range.Text.Trim();
                int paraLevel = (int)para.OutlineLevel;

                if (string.IsNullOrEmpty(headingText)) continue;

                if (styleNames.Contains(styleName))
                {
                    bool hasComment = false;
                    foreach (Comment comment in para.Range.Comments)
                    {
                        if (comment.Range.Text.Contains(commentText))
                        {
                            hasComment = true;
                            break;
                        }
                    }

                    if (hasComment)
                    {
                        // コメント付き見出しを見つけたらブロック開始
                        headings.Add(headingText);
                        currentLevel = paraLevel;
                        inBlock = true;
                        continue;
                    }
                }

                if (inBlock && currentLevel.HasValue)
                {
                    if (paraLevel > currentLevel)
                    {
                        // 配下の見出しを追加
                        headings.Add(headingText);
                    }
                    else if (paraLevel <= currentLevel)
                    {
                        // 同レベルまたは上位レベルの見出しが来たらブロック終了
                        inBlock = false;
                        currentLevel = null;
                    }
                }
            }

            return headings;
        }

        // 以下の条件を共に満たす場合に、見出しのタイトルとその配下の見出しを取得
        //  ・見出しスタイルが「MJS_見出し 1（項番なし）」または「MJS_見出し 2（項番なし）」である
        //  ・見出しのタイトルが「はじめに」または「マニュアル内の記号・表記について」である
        private List<string> GetSpecificHeadingsWithSubheadings()
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            List<string> result = new List<string>();

            // 対象スタイル・タイトル
            var targetStyles = new HashSet<string> { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）" };
            var targetTitles = new HashSet<string> { "はじめに", "マニュアル内の記号・表記について" };

            int? currentLevel = null;
            bool inBlock = false;

            foreach (Paragraph para in activeDocument.Paragraphs)
            {
                string styleName = para.get_Style().NameLocal;
                string text = para.Range.Text.Trim();
                int paraLevel = (int)para.OutlineLevel;

                if (string.IsNullOrEmpty(text)) continue;

                if (targetStyles.Contains(styleName) && targetTitles.Contains(text))
                {
                    // 対象見出しを見つけたらブロック開始
                    result.Add(text);
                    currentLevel = paraLevel;
                    inBlock = true;
                    continue;
                }

                if (inBlock && currentLevel.HasValue)
                {
                    if (paraLevel > currentLevel)
                    {
                        // 配下の見出しを追加
                        result.Add(text);
                    }
                    else if (paraLevel <= currentLevel)
                    {
                        // 同レベルまたは上位レベルの見出しが来たらブロック終了
                        inBlock = false;
                        currentLevel = null;
                    }
                }
            }

            return result;
        }

        // アウトラインレベルと見出しテキストをメッセージボックスで表示（動作確認用）
        private void ShowHeadingsWithOutlineLevels()
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            var sb = new StringBuilder();

            foreach (Paragraph para in activeDocument.Paragraphs)
            {
                int outlineLevel = (int)para.OutlineLevel;
                string text = para.Range.Text.Trim();

                if (outlineLevel >= 1 && outlineLevel <= 9 && !string.IsNullOrEmpty(text))
                {
                    sb.AppendLine($"レベル{outlineLevel}: {text}");
                }
            }

            string result = sb.Length > 0 ? sb.ToString() : "見出しが見つかりませんでした。";
            MessageBox.Show(result, "見出しのアウトラインレベル一覧", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
