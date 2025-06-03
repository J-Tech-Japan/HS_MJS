using System.Collections.Generic;
using MJS_fileJoin;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
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

        private List<string> GetBookmarkNames(Word.Document objDocLast)
        {
            List<string> names = new List<string>();
            foreach (Word.Bookmark wb in objDocLast.Bookmarks)
                names.Add(wb.Name);
            return names;
        }

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
