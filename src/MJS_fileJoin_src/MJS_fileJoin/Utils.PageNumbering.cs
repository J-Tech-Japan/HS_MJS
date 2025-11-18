// Utils.PageNumbering.cs

using System;
using System.Diagnostics;
using Word = Microsoft.Office.Interop.Word;

namespace MJS_fileJoin
{
    internal partial class Utils
    {
        /// <summary>
        /// セクションのページ番号を通し番号として設定します
        /// ヘッダーとフッターは保持されます（空の場合のみ前のセクションとリンク）
        /// </summary>
        public static void ResetPageNumbering(Word.Document document, MainForm form)
        {
            form.label10.Text = "ページ番号通し番号に設定中...";
            
            if (document.Sections.Count <= 1)
            {
                Trace.WriteLine("  セクション数が1以下のため、処理をスキップ");
                return;
            }

            form.progressBar1.Maximum = document.Sections.Count;
            form.progressBar1.Value = 1;

            int processedCount = 0;
            int resetCount = 0;
            int preservedCount = 0;
            int linkedCount = 0;
            int errorCount = 0;

            try
            {
                // 2番目のセクション以降を処理（最初のセクションはそのまま）
                for (int i = 2; i <= document.Sections.Count; i++)
                {
                    try
                    {
                        Word.Section section = document.Sections[i];
                        
                        // 空のヘッダー・フッターのみ前のセクションとリンク
                        int linked = LinkEmptyHeaderFootersToPrevious(section);
                        linkedCount += linked;
                        
                        // 非空のヘッダー・フッターは保持される
                        int preserved = PreserveNonEmptyHeaderFooters(section);
                        preservedCount += preserved;
                        
                        // ページ番号の再開始を無効化（通し番号にする）
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

                Trace.WriteLine($"  処理完了: {processedCount}セクション (空リンク: {linkedCount}, 保持: {preservedCount}, PageNumbers設定: {resetCount}, エラー: {errorCount})");
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"  ページ番号設定で致命的エラー: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 空のヘッダー・フッターのみ前のセクションとリンク
        /// </summary>
        private static int LinkEmptyHeaderFootersToPrevious(Word.Section section)
        {
            int linkedCount = 0;
            
            try
            {
                // フッターの処理
                linkedCount += LinkEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                linkedCount += LinkEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                linkedCount += LinkEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
                
                // ヘッダーの処理
                linkedCount += LinkEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                linkedCount += LinkEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                linkedCount += LinkEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"    LinkEmptyHeaderFootersToPrevious エラー: {ex.Message}");
            }
            
            return linkedCount;
        }

        /// <summary>
        /// 空のHeaderFooterのみLinkToPreviousを設定
        /// </summary>
        private static int LinkEmptyHeaderFooter(Word.HeaderFooter headerFooter)
        {
            try
            {
                // 既にリンクされている場合はスキップ
                if (headerFooter.LinkToPrevious)
                {
                    return 0;
                }
                
                // 範囲が存在し、空の場合のみリンク
                var range = headerFooter.Range;
                if (range != null && string.IsNullOrWhiteSpace(range.Text.Trim()))
                {
                    headerFooter.LinkToPrevious = true;
                    return 1;
                }
            }
            catch
            {
                // ヘッダー/フッターが存在しない場合は無視
            }
            
            return 0;
        }

        /// <summary>
        /// 非空のヘッダー・フッターを保持（LinkToPreviousをfalseに）
        /// </summary>
        private static int PreserveNonEmptyHeaderFooters(Word.Section section)
        {
            int preservedCount = 0;
            
            try
            {
                // フッターの処理
                preservedCount += PreserveNonEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                preservedCount += PreserveNonEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                preservedCount += PreserveNonEmptyHeaderFooter(section.Footers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
                
                // ヘッダーの処理
                preservedCount += PreserveNonEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterPrimary]);
                preservedCount += PreserveNonEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage]);
                preservedCount += PreserveNonEmptyHeaderFooter(section.Headers[Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages]);
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"    PreserveNonEmptyHeaderFooters エラー: {ex.Message}");
            }
            
            return preservedCount;
        }

        /// <summary>
        /// 非空のHeaderFooterのLinkToPreviousをfalseに設定
        /// </summary>
        private static int PreserveNonEmptyHeaderFooter(Word.HeaderFooter headerFooter)
        {
            try
            {
                // 範囲が存在し、非空の場合はリンクを解除
                var range = headerFooter.Range;
                if (range != null && !string.IsNullOrWhiteSpace(range.Text.Trim()))
                {
                    if (headerFooter.LinkToPrevious)
                    {
                        headerFooter.LinkToPrevious = false;
                        return 1;
                    }
                }
            }
            catch
            {
                // ヘッダー/フッターが存在しない場合は無視
            }
            
            return 0;
        }

        /// <summary>
        /// セクションのPageNumbers設定で再開始を無効化
        /// </summary>
        private static bool ResetPageNumbersInSection(Word.Section section)
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
        private static bool ResetPageNumbersInHeaderFooter(Word.HeaderFooter headerFooter)
        {
            try
            {
                // PageNumbersの存在をチェック（Countが使えないが、実際に設定は可能）
                var pageNumbers = headerFooter.PageNumbers;
                
                // ページ番号の再開始を無効化（前のセクションから継続）
                pageNumbers.RestartNumberingAtSection = false;
                pageNumbers.StartingNumber = 0; // 0 = 継続
                
                return true;
            }
            catch
            {
                // PageNumbersが存在しない場合や、アクセスできない場合は無視
            }
            
            return false;
        }
    }
}
