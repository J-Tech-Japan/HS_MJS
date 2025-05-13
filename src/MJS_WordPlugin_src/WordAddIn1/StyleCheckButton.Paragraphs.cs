using Microsoft.Office.Interop.Word;

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;

using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 段落を処理するメソッド
        private void ProcessParagraphs(Word.Document activeDoc, List<string> styleList, Stopwatch sw, ref int pro, ref bool isProcessing, ref bool isProcessHalted, ref bool flag)
        {
            foreach (Paragraph paragraph in activeDoc.Paragraphs)
            {
                // プログレスバーの更新
                ProgressBar.SetProgressBarValue(++pro);
                TimeSpan ts = sw.Elapsed;
                ProgressBar.ProgressTime(ts);

                // プログレスバーが破棄されている場合の処理
                if (ProgressBar.mInstance.IsDisposed)
                {
                    DeleteInvalidComments(activeDoc.Comments);
                    isProcessHalted = true;
                    break;
                }

                try
                {
                    var pStyle = paragraph.Range.ParagraphStyle();
                    // 空の段落をスキップ
                    if (IsEmptyParagraph(paragraph, activeDoc.Styles[-1].NameLocal))
                    {
                        continue;
                    }

                    // 特定のスタイルをスキップ
                    if (IsExcludedStyle(pStyle))
                    {
                        continue;
                    }

                    // 使用できないスタイルを検出
                    if (!styleList.Contains(pStyle))
                    {
                        AddComment(paragraph.Range, $"【{pStyle}】: 使用できない書式です。");
                        flag = true;
                    }

                    // 手順番号リセット用スタイルのチェック
                    if (pStyle == "MJS_見出し-手順")
                    {
                        isProcessing = true;
                    }
                    else if (isProcessing && pStyle == "MJS_手順番号リセット用")
                    {
                        isProcessing = false;
                    }
                    else if (isProcessing && pStyle == "MJS_手順文")
                    {
                        isProcessing = false;
                        AddComment(paragraph.Range, "上の段落に【MJS_手順番号リセット用】スタイルを挿入してください。");
                        flag = true;
                    }
                }
                catch (Exception ex)
                {
                    Debug.WriteLine($"段落処理中に例外が発生しました: {ex.Message}");
                }
            }
        }

        // ヘルパーメソッド: 空段落を判定
        private bool IsEmptyParagraph(Word.Paragraph paragraph, string emptyStyleName)
        {
            return paragraph.Range.ParagraphStyle() == emptyStyleName &&
                   string.IsNullOrEmpty(paragraph.Range.Text.Trim().Replace("\u0007", ""));
        }

        // ヘルパーメソッド: 除外対象のスタイルの判定
        private bool IsExcludedStyle(string styleName)
        {
            var excludedStyles = new List<string>
                {
                    "見出し 7",
                    "章扉-見出し1",
                    "章扉-目次1",
                    "奥付",
                    "索引"
                };

            return excludedStyles.Any(style => styleName.Contains(style));
        }

        // ヘルパーメソッド: コメントを追加
        private void AddComment(Word.Range range, string commentText)
        {
            range.Comments.Add(range, commentText);
        }
    }
}
