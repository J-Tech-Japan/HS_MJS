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
            if (document.Sections.Count <= 1)
            {
                Trace.WriteLine("  セクション数が1以下のため、処理をスキップ");
                return;
            }

            using (var progress = BeginProgress(form, "ページ番号通し番号に設定中...", document.Sections.Count))
            {
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
                            
                            // ヘッダー・フッターの処理を統合（1回のループで完了）
                            ProcessHeaderFooters(section, ref linkedCount, ref preservedCount, ref resetCount);
                            
                            processedCount++;
                            
                            // UI更新頻度を調整（10セクションごと、または最後）
                            if (i % 10 == 0 || i == document.Sections.Count)
                            {
                                progress.SetValue(i);
                            }
                        }
                        catch (Exception ex)
                        {
                            errorCount++;
                            Trace.WriteLine($"  セクション {i} のページ番号設定でエラー: {ex.Message}");
                        }
                    }

                    progress.Complete();
                    Trace.WriteLine($"  処理完了: {processedCount}セクション (空リンク: {linkedCount}, 保持: {preservedCount}, PageNumbers設定: {resetCount}, エラー: {errorCount})");
                }
                catch (Exception ex)
                {
                    Trace.WriteLine($"  ページ番号設定で致命的エラー: {ex.Message}");
                    throw;
                }
            }
        }

        /// <summary>
        /// ヘッダー・フッターの処理を統合（リンク、保持、ページ番号設定を1回で実行）
        /// </summary>
        private static void ProcessHeaderFooters(
            Word.Section section, 
            ref int linkedCount, 
            ref int preservedCount, 
            ref int resetCount)
        {
            // 処理対象のインデックス
            var indices = new[]
            {
                Word.WdHeaderFooterIndex.wdHeaderFooterPrimary,
                Word.WdHeaderFooterIndex.wdHeaderFooterFirstPage,
                Word.WdHeaderFooterIndex.wdHeaderFooterEvenPages
            };

            try
            {
                // フッターの処理
                foreach (var index in indices)
                {
                    ProcessSingleHeaderFooter(section.Footers[index], ref linkedCount, ref preservedCount, ref resetCount);
                }
                
                // ヘッダーの処理
                foreach (var index in indices)
                {
                    ProcessSingleHeaderFooter(section.Headers[index], ref linkedCount, ref preservedCount, ref resetCount);
                }
            }
            catch (Exception ex)
            {
                Trace.WriteLine($"    ProcessHeaderFooters エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 単一のHeaderFooterを処理（空判定、リンク/保持、ページ番号設定を統合）
        /// </summary>
        private static void ProcessSingleHeaderFooter(
            Word.HeaderFooter headerFooter, 
            ref int linkedCount, 
            ref int preservedCount, 
            ref int resetCount)
        {
            try
            {
                // 既にリンクされているかチェック
                bool isLinked = headerFooter.LinkToPrevious;
                
                // 範囲取得
                var range = headerFooter.Range;
                if (range != null)
                {
                    string text = range.Text?.Trim() ?? string.Empty;
                    bool isEmpty = string.IsNullOrWhiteSpace(text);

                    // 空の場合：リンク設定
                    if (isEmpty && !isLinked)
                    {
                        headerFooter.LinkToPrevious = true;
                        linkedCount++;
                    }
                    // 非空の場合：リンク解除して保持
                    else if (!isEmpty && isLinked)
                    {
                        headerFooter.LinkToPrevious = false;
                        preservedCount++;
                    }
                }

                // ページ番号の再開始を無効化（通し番号化）
                var pageNumbers = headerFooter.PageNumbers;
                if (pageNumbers != null)
                {
                    pageNumbers.RestartNumberingAtSection = false;
                    pageNumbers.StartingNumber = 0; // 0 = 継続
                    resetCount++;
                }
            }
            catch
            {
                // ヘッダー/フッターが存在しない場合や、アクセスできない場合は無視
            }
        }
    }
}
