using System.Collections.Generic;
using System.Diagnostics;
using MJS_fileJoin;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        private void RemoveSectionsInRangeByStyle(
            Word.Document objDocLast,
            string[] lsStyleName,
            int chapCnt,
            ref int chapCntLast,
            MainForm form)
        {
            form.label10.Text = "重複箇所削除中...";
            form.progressBar1.Maximum = 11;
            form.progressBar1.Value = 1;

            // 有効なスタイル名だけを抽出
            var validStyleNames = GetValidStyleNames(objDocLast, lsStyleName);

            // 新しい配列でループ
            foreach (string styleName in validStyleNames)
            {
                object styleObject = styleName;
                int i = chapCnt + 1;
                int deletedCount = 0;
                while (i <= chapCntLast)
                {
                    if (i > objDocLast.Sections.Count)
                        break;

                    Word.Range wr = objDocLast.Sections[i].Range;
                    wr.Find.ClearFormatting();
                    wr.Find.set_Style(ref styleObject);
                    wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    wr.Find.Execute();
                    
                    if (wr.Find.Found)
                    {
                        Trace.WriteLine($"[RemoveSectionsInRangeByStyle] セクション削除: スタイル='{styleName}', セクション番号={i}");
                        objDocLast.Sections[i].Range.Delete();
                        chapCntLast--;
                        deletedCount++;
                        // iはインクリメントしない（削除により次のセクションが現在位置に来るため）
                    }
                    else
                    {
                        i++;
                    }
                }
                if (deletedCount > 0)
                {
                    Trace.WriteLine($"[RemoveSectionsInRangeByStyle] スタイル '{styleName}' で {deletedCount} セクション削除");
                }
                form.progressBar1.Increment(1);
            }
        }

        // 指定したスタイル名が見つかったらlastフラグをtrueにして進捗バーを進める
        private void SetLastFlagIfStyleFound(
            Word.Document objDocLast,
            string[] styleNames,
            ref bool last,
            int chapCntLast,
            MainForm form)
        {
            // 有効なスタイル名だけを抽出
            var validStyleNames = GetValidStyleNames(objDocLast, styleNames);

            foreach (string styleName in validStyleNames)
            {
                object styleObject = styleName;
                int allChap = objDocLast.Sections.Count;
                for (int i = allChap; i > chapCntLast; i--)
                {
                    Word.Range wr = objDocLast.Sections[i].Range;
                    wr.Find.ClearFormatting();
                    wr.Find.set_Style(ref styleObject);
                    wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                    wr.Find.Execute();
                    if (wr.Find.Found)
                    {
                        last = true;
                        break;
                    }
                }
                form.progressBar1.Increment(1);
            }
        }

        // 末尾からchapCntLastより大きいセクションを後方走査
        // 指定スタイルで見つかったらlastフラグに応じて削除
        // 例外時はbreak
        private void RemoveSectionsFromEndByStyleWithLastFlag(
            Word.Document objDocLast,
            string[] styleNames,
            ref int chapCntLast,
            ref bool last,
            MainForm form)
        {
            form.label10.Text = "章扉章節項番号修正中...";
            form.progressBar1.Maximum = objDocLast.Sections.Count;
            form.progressBar1.Value = 1;

            // 有効なスタイル名だけを抽出
            var validStyleNames = GetValidStyleNames(objDocLast, styleNames);

            foreach (string styleName in validStyleNames)
            {
                object styleObject = styleName;
                int i = objDocLast.Sections.Count;
                int deletedCount = 0;
                
                while (i > chapCntLast)
                {
                    try
                    {
                        if (i > objDocLast.Sections.Count)
                        {
                            i = objDocLast.Sections.Count;
                            continue;
                        }

                        Word.Range wr = objDocLast.Sections[i].Range;
                        wr.Find.ClearFormatting();
                        wr.Find.set_Style(ref styleObject);
                        wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        wr.Find.Execute();
                        
                        if (wr.Find.Found)
                        {
                            if (last)
                            {
                                Trace.WriteLine($"[RemoveSectionsFromEndByStyleWithLastFlag] セクション保持（lastフラグ）: スタイル='{styleName}', セクション番号={i}");
                                last = false;
                                i--;
                            }
                            else
                            {
                                Trace.WriteLine($"[RemoveSectionsFromEndByStyleWithLastFlag] セクション削除: スタイル='{styleName}', セクション番号={i}");
                                objDocLast.Sections[i].Range.Delete();
                                chapCntLast--;
                                deletedCount++;
                                // iはデクリメントしない（削除により次のセクションが現在位置に来るため）
                            }
                        }
                        else
                        {
                            i--;
                        }
                    }
                    catch
                    {
                        break;
                    }
                }
                if (deletedCount > 0)
                {
                    Trace.WriteLine($"[RemoveSectionsFromEndByStyleWithLastFlag] スタイル '{styleName}' で {deletedCount} セクション削除");
                }
            }
        }

        // 指定したスタイル名のセクションを後方から1つだけ残して削除
        private void RemoveSectionsByStyleKeepLast(Word.Document doc, string styleName, MainForm form)
        {
            form.label10.Text = "索引更新中...";
            bool found = false;
            int i = doc.Sections.Count;
            int deletedCount = 0;
            object styleObject = styleName;
            
            while (i > 0)
            {
                if (i > doc.Sections.Count)
                {
                    i = doc.Sections.Count;
                    continue;
                }

                Word.Range wr = doc.Sections[i].Range;
                wr.Find.ClearFormatting();
                wr.Find.set_Style(ref styleObject);
                wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                wr.Find.Execute();
                
                if (wr.Find.Found)
                {
                    if (found)
                    {
                        Trace.WriteLine($"[RemoveSectionsByStyleKeepLast] セクション削除: スタイル='{styleName}', セクション番号={i}");
                        doc.Sections[i].Range.Delete();
                        deletedCount++;
                    }
                    else
                    {
                        Trace.WriteLine($"[RemoveSectionsByStyleKeepLast] セクション保持（最後の1つ）: スタイル='{styleName}', セクション番号={i}");
                        found = true;
                        i--;
                    }
                }
                else
                {
                    i--;
                }
            }
            if (deletedCount > 0)
            {
                Trace.WriteLine($"[RemoveSectionsByStyleKeepLast] スタイル '{styleName}' で {deletedCount} セクション削除（1つ保持）");
            }
        }

        // ヘルパーメソッド：有効なスタイル名だけを抽出
        private List<string> GetValidStyleNames(Word.Document doc, IEnumerable<string> styleNames)
        {
            var validStyleNames = new List<string>();
            foreach (string styleName in styleNames)
            {
                foreach (Word.Style style in doc.Styles)
                {
                    if (style.NameLocal == styleName)
                    {
                        validStyleNames.Add(styleName);
                        break;
                    }
                }
            }
            return validStyleNames;
        }
    }
}
