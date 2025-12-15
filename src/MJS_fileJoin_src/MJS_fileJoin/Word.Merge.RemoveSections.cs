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
            // 有効なスタイル名だけを抽出
            var validStyleNames = GetValidStyleNames(objDocLast, lsStyleName);
            
            using (var progress = Utils.BeginProgress(form, "重複箇所削除中...", validStyleNames.Count))
            {
                int styleIndex = 0;
                
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
                    
                    styleIndex++;
                    progress.SetValue(styleIndex);
                }
                
                progress.Complete();
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
            
            using (var progress = Utils.BeginProgress(form, "索引見出し検索中...", validStyleNames.Count))
            {
                int styleIndex = 0;
                
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
                    
                    styleIndex++;
                    progress.SetValue(styleIndex);
                }
                
                progress.Complete();
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
            // 有効なスタイル名だけを抽出
            var validStyleNames = GetValidStyleNames(objDocLast, styleNames);
            
            using (var progress = Utils.BeginProgress(form, "章扉章節項番号修正中...", objDocLast.Sections.Count))
            {
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
                                    // セクションは削除されないため、次のセクションに移動
                                    i--;
                                }
                                else
                                {
                                    Trace.WriteLine($"[RemoveSectionsFromEndByStyleWithLastFlag] セクション削除: スタイル='{styleName}', セクション番号={i}");
                                    
                                    objDocLast.Sections[i].Range.Delete();
                                    deletedCount++;
                                    // iはデクリメントしない（削除により次のセクションが現在位置に来るため）
                                }
                            }
                            else
                            {
                                i--;
                            }
                            
                            // プログレスバー更新（10セクションごと）
                            if ((objDocLast.Sections.Count - i) % 10 == 0)
                            {
                                progress.SetValue(objDocLast.Sections.Count - i);
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
                
                progress.Complete();
            }
        }

        // 指定したスタイル名のセクションを後方から1つだけ残して削除
        private void RemoveSectionsByStyleKeepLast(Word.Document doc, string styleName, MainForm form)
        {
            bool found = false;
            int i = doc.Sections.Count;
            int deletedCount = 0;
            object styleObject = styleName;
            
            using (var progress = Utils.BeginProgress(form, "索引セクション削除中...", doc.Sections.Count))
            {
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
                    
                    // プログレスバー更新（10セクションごと）
                    if ((doc.Sections.Count - i) % 10 == 0)
                    {
                        progress.SetValue(doc.Sections.Count - i);
                    }
                }
                if (deletedCount > 0)
                {
                    Trace.WriteLine($"[RemoveSectionsByStyleKeepLast] スタイル '{styleName}' で {deletedCount} セクション削除（1つ保持）");
                }
                
                progress.Complete();
            }
        }

        // ヘルパーメソッド：有効なスタイル名だけを抽出
        private List<string> GetValidStyleNames(Word.Document doc, IEnumerable<string> styleNames)
        {
            var validStyleNames = new List<string>();
            foreach (string styleName in styleNames)
            {
                // "MJS_マニュアルタイトル"の場合は部分一致で検索
                if (styleName == "MJS_マニュアルタイトル")
                {
                    string searchPattern = "マニュアルタイトル";
                    foreach (Word.Style style in doc.Styles)
                    {
                        if (style.NameLocal.Contains(searchPattern))
                        {
                            // 重複を避けるため、まだリストに含まれていない場合のみ追加
                            if (!validStyleNames.Contains(style.NameLocal))
                            {
                                validStyleNames.Add(style.NameLocal);
                                Trace.WriteLine($"[GetValidStyleNames] 部分一致で検出: '{style.NameLocal}' (検索パターン: '{searchPattern}')");
                            }
                        }
                    }
                }
                else
                {
                    // 通常の完全一致検索
                    foreach (Word.Style style in doc.Styles)
                    {
                        if (style.NameLocal == styleName)
                        {
                            validStyleNames.Add(styleName);
                            break;
                        }
                    }
                }
            }
            return validStyleNames;
        }
    }
}
