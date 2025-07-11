using System.Collections.Generic;
using System.Text.RegularExpressions;
using MJS_fileJoin;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // REF形式のハイパーリンクをHYPERLINK形式に変換
        // 【注意！】スタイルに関係なくすべてのHYPERLINKフィールドをREFに変換
        //public static void ConvertH4HyperlinkToRefAll(Word.Document doc)
        //{
        //    int count = 0;
        //    foreach (Word.Field field in doc.Fields)
        //    {
        //        string code = field.Code.Text;
        //        int pos = code.IndexOf("HYPERLINK _Ref");
        //        if (pos >= 0)
        //        {
        //            // "_Ref"の直後からIDを抽出
        //            string refId = code.Substring(pos + "HYPERLINK ".Length);
        //            // 余分な部分を除去（スペースや改行など）
        //            refId = refId.Trim().Split(' ')[0];
        //            // フィールドコードをREF形式に書き換え
        //            field.Code.Text = $" REF {refId} \\h ";
        //            field.Update();
        //            count++;
        //        }
        //    }
        //}

        // 特定のスタイル名を持つ「HYPERLINK _Ref...」形式のフィールドを「REF ... \h」形式に変換
        public static void ConvertHyperlinkToRef(Word.Document doc, List<string> targetStyleNames)
        {
            int count = 0;
            foreach (Word.Field field in doc.Fields)
            {
                string code = field.Code.Text;
                int pos = code.IndexOf("HYPERLINK _Ref");
                if (pos >= 0)
                {
                    // スタイル名を取得
                    Word.Range rng = field.Result;
                    string styleName = rng.get_Style() is Word.Style style ? style.NameLocal : rng.get_Style().ToString();
                    // 指定リストに含まれていなければスキップ
                    if (!targetStyleNames.Contains(styleName)) continue;

                    // "_Ref"の直後からIDを抽出
                    string refId = code.Substring(pos + "HYPERLINK ".Length);
                    // 余分な部分を除去（スペースや改行など）
                    refId = refId.Trim().Split(' ')[0];
                    // フィールドコードをREF形式に書き換え
                    field.Code.Text = $" REF {refId} \\h ";
                    field.Update();
                    count++;
                }
            }
        }

        // ハイパーリンクの更新
        private void UpdateHyperlinks(Word.Document objDocLast, MainForm fm)
        {
            fm.label10.Text = "ハイパーリンク更新中...";
            List<string> bookmarkNames = GetBookmarkNames(objDocLast);

            fm.progressBar1.Value = 0;
            fm.progressBar1.Maximum = objDocLast.Fields.Count;

            foreach (Word.Field wf in objDocLast.Fields)
            {
                fm.progressBar1.Increment(1);
                if (wf.Type == Word.WdFieldType.wdFieldHyperlink)
                {
                    if (wf.Code.Text.Contains("\"http")) continue;
                    string text = ExtractHyperlinkText(wf.Code.Text);
                    if (text == null) continue;

                    string[] subtext = text.Split('\\');
                    text = subtext[subtext.Length - 1];
                    subtext = text.Split('/');
                    text = subtext[subtext.Length - 1];
                    string normalized = text.Replace(".html", "").Replace("#", "♯").Trim();
                    if (bookmarkNames.Contains(normalized))
                    {
                        wf.Code.Text = @"HYPERLINK \l """ + normalized + @"""";
                        wf.Update();
                    }
                    else
                    {
                        wf.Unlink();
                    }
                }
            }

            UpdateHyperlinkDisplayText(objDocLast);
        }

        // ヘルパーメソッド：ドキュメント内のブックマーク名を取得
        private List<string> GetBookmarkNames(Word.Document objDocLast)
        {
            List<string> names = new List<string>();
            foreach (Word.Bookmark wb in objDocLast.Bookmarks)
                names.Add(wb.Name);
            return names;
        }

        // ヘルパーメソッド：ハイパーリンクのテキストを抽出
        private string ExtractHyperlinkText(string codeText)
        {
            if (!codeText.Contains(@"\l"))
            {
                return Regex.Replace(codeText, @".*?""([^""]*?)"".*?", "$1");
            }
            else
            {
                if (!Regex.IsMatch(codeText, @".*?""[^""]*?"".*?""[^""]*?"".*?")) return null;
                return Regex.Replace(codeText, @".*?""([^""]*?)"".*?""([^""]*?)"".*?", "$1#$2");
            }
        }

        // ヘルパーメソッド：ハイパーリンクの表示テキストを更新
        private void UpdateHyperlinkDisplayText(Word.Document objDocLast)
        {
            foreach (Word.Hyperlink wh in objDocLast.Hyperlinks)
            {
                if (Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}") ||
                    Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}♯\w{3}\d{5}"))
                    wh.TextToDisplay = Regex.Replace(wh.TextToDisplay, @".*(\d+\.)+\d+[\s　\t]", "");
            }
        }
    }
}
