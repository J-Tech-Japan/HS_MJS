using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;
using System.Linq;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // 指定スタイルの章扉・章節項番号を後方から検索し、2つ目以降を削除（リファクタリング版）
        private void RemoveLastSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCntLast, ref bool last)
        {
            fm.label10.Text = "章扉章節項番号修正中...";
            fm.progressBar1.Maximum = objDocLast.Sections.Count;
            fm.progressBar1.Value = 1;

            string[] lastItems = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次" };
            foreach (string styleName in lastItems)
            {
                // スタイルが存在しない場合はスキップ
                if (!objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                {
                    continue;
                }
                object styleObject = styleName;
                int i = objDocLast.Sections.Count;
                while (i > chapCntLast)
                {
                    Word.Range wr = objDocLast.Sections[i].Range;
                    wr.Find.ClearFormatting();
                    wr.Find.set_Style(ref styleObject);
                    //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    bool found = wr.Find.Execute();
                    if (found && wr.Find.Found)
                    {
                        if (last)
                        {
                            last = false;
                        }
                        else
                        {
                            wr.Delete();
                            // セクション削除後はインデックスを1つ戻す
                            i--;
                            chapCntLast--;
                        }
                    }
                    else
                    {
                        // 見つからなければ次のセクションへ
                        i--;
                    }
                    // 進捗バーを更新
                    if (fm.progressBar1.Value < fm.progressBar1.Maximum)
                        fm.progressBar1.Value++;
                }
            }
        }

        // 指定スタイルの重複セクションを削除（リファクタリング版）
        private void RemoveDuplicateSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCnt, ref int chapCntLast)
        {
            fm.label10.Text = "重複箇所削除中...";
            fm.progressBar1.Maximum = 11;
            fm.progressBar1.Value = 1;

            string[] lsStyleName = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次", "奥付タイトル", "索引見出し" };
            foreach (string styleName in lsStyleName)
            {
                // スタイルが存在しない場合はスキップ
                if (!objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                {
                    continue;
                }
                object styleObject = styleName;
                int i = chapCnt + 1;
                while (i <= chapCntLast)
                {
                    Word.Range wr = objDocLast.Sections[i].Range;
                    wr.Find.ClearFormatting();
                    wr.Find.set_Style(ref styleObject);
                    //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    bool found = wr.Find.Execute();
                    if (found && wr.Find.Found)
                    {
                        wr.Delete();
                        chapCntLast--;
                        // セクション削除後は同じインデックスで再チェック
                    }
                    else
                    {
                        i++;
                    }
                    // 進捗バーを更新
                    if (fm.progressBar1.Value < fm.progressBar1.Maximum)
                        fm.progressBar1.Value++;
                }
                fm.progressBar1.Increment(1);
            }
        }

        // 指定スタイルのセクションを後方から検索し、2つ目以降を削除（リファクタリング版）
        private void RemoveDuplicateIndexSections(Word.Document doc, string styleName)
        {
            // スタイルが存在しない場合はスキップ
            if (!doc.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
            {
                return;
            }
            object styleObject = styleName;
            bool foundFirst = false;
            int sectionCount = doc.Sections.Count;
            for (int i = sectionCount; i > 0; i--)
            {
                Word.Range wr = doc.Sections[i].Range;
                wr.Find.ClearFormatting();
                wr.Find.set_Style(ref styleObject);
                wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                bool found = wr.Find.Execute();
                if (found && wr.Find.Found)
                {
                    if (foundFirst)
                    {
                        wr.Delete();
                        i--; // セクション削除後はインデックスを1つ戻す
                    }
                    else
                    {
                        foundFirst = true;
                    }
                }
            }
        }
    }
}
