using System;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        /// <summary>
        /// 結合後のセクションのページ番号を通し番号として設定し直す
        /// </summary>
        private void ResetPageNumbering(Word.Document document, MainForm form)
        {
            form.label10.Text = "ページ番号を通し番号に設定中...";
            
            if (document.Sections.Count <= 1)
            {
                Trace.WriteLine("  セクションが1つ以下のため、処理をスキップ");
                return;
            }

            form.progressBar1.Maximum = document.Sections.Count;
            form.progressBar1.Value = 1;

            int processedCount = 0;
            int linkedCount = 0;
            int resetCount = 0;
            int errorCount = 0;

            try
            {
                // 2番目のセクション以降を処理（最初のセクションはそのまま）
                for (int i = 2; i <= document.Sections.Count; i++)
                {
                    try
                    {
                        Word.Section section = document.Sections[i];
                        
                        // アプローチ1: LinkToPreviousを設定してヘッダー/フッターを前のセクションとリンク
                        bool linked = LinkHeaderFootersToPrevious(section);
                        if (linked) linkedCount++;
                        
                        // アプローチ2: PageNumbers設定で明示的に再開始を無効化
                        bool reset = ResetPageNumbersInSection(section);
                        if (reset) resetCount++;
                        
                        processedCount++;
                    }
                    catch (Exception ex)
                    {
                        errorCount++;
                        Trace.WriteLine($"  セクション {i} のページ番号設定でエラー: {ex.Message}");
                    }
                    
                    form.progressBar1.Increment(1);
                }

                Trace.WriteLine($"  処理完了: {processedCount}セクション (LinkToPrevious: {linkedCount}, PageNumbers設定: {resetCount}, エラー: {errorCount})");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"  ページ番号設定で致命的エラー: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// セクションのヘッダー/フッターを前のセクションとリンク
        /// </summary>
        private bool LinkHeaderFootersToPrevious(Word.Section section)
        {
            bool anyLinked = false;
            
            try
            {
                // フッターをリンク
                anyLinked |= SetLinkToPrevious(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                anyLinked |= SetLinkToPrevious(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                anyLinked |= SetLinkToPrevious(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
                
                // ヘッダーをリンク
                anyLinked |= SetLinkToPrevious(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                anyLinked |= SetLinkToPrevious(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                anyLinked |= SetLinkToPrevious(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
            }
            catch
            {
                // 無視
            }
            
            return anyLinked;
        }

        /// <summary>
        /// 個別のHeaderFooterでLinkToPreviousを設定
        /// </summary>
        private bool SetLinkToPrevious(Word.HeaderFooter headerFooter)
        {
            try
            {
                if (!headerFooter.LinkToPrevious)
                {
                    headerFooter.LinkToPrevious = true;
                    return true;
                }
            }
            catch
            {
                // ヘッダー/フッターが存在しない場合は無視
            }
            
            return false;
        }

        /// <summary>
        /// セクションのPageNumbers設定で再開始を無効化
        /// </summary>
        private bool ResetPageNumbersInSection(Word.Section section)
        {
            bool anyReset = false;
            
            try
            {
                // フッターのPageNumbersをリセット
                anyReset |= ResetPageNumbersInHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                anyReset |= ResetPageNumbersInHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                anyReset |= ResetPageNumbersInHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
                
                // ヘッダーのPageNumbersをリセット
                anyReset |= ResetPageNumbersInHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                anyReset |= ResetPageNumbersInHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                anyReset |= ResetPageNumbersInHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
            }
            catch
            {
                // 無視
            }
            
            return anyReset;
        }

        /// <summary>
        /// 個別のHeaderFooterでページ番号の再開始を無効化
        /// </summary>
        private bool ResetPageNumbersInHeaderFooter(Word.HeaderFooter headerFooter)
        {
            try
            {
                // PageNumbersの存在をチェック（Countだけでなく、実際に設定を試みる）
                var pageNumbers = headerFooter.PageNumbers;
                
                // ページ番号の再開始を無効化（前のセクションから継続）
                pageNumbers.RestartNumberingAtSection = false;
                pageNumbers.StartingNumber = 0; // 0 = 継続
                
                return true;
            }
            catch
            {
                // PageNumbersが存在しない場合やアクセスできない場合は無視
            }
            
            return false;
        }
    }
}
