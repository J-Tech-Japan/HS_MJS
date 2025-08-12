// Utils.RemoveAllImageMarkers.cs

using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {

        /// <summary>
        /// WordドキュメントからExtractImagesAndCanvasFromWordWithTextで設定された全ての画像マーカーを削除する
        /// </summary>
        /// <param name="document">マーカーを削除する対象のWordドキュメント</param>
        /// <returns>削除されたマーカーの数</returns>
        public static int RemoveAllImageMarkers(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            int removedCount = 0;

            try
            {
                // ドキュメント全体の範囲を取得
                var fullRange = document.Range();

                // 検索条件を設定
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;
                find.MatchSoundsLike = false;
                find.MatchAllWordForms = false;

                // マーカーを順次検索して削除
                while (find.Execute())
                {
                    try
                    {
                        // マーカーテキストの範囲を取得
                        var markerRange = fullRange.Duplicate;

                        // マーカーの前後の改行も含めて削除範囲を拡張
                        ExtendRangeToIncludeAssociatedLineBreaks(markerRange);

                        // マーカーとその前後の改行を削除
                        markerRange.Delete();
                        removedCount++;

                        // 削除後、検索範囲をリセット
                        fullRange.SetRange(0, document.Range().End);
                        find.ClearFormatting();
                        find.Text = @"\[IMAGEMARKER:*\]";
                        find.MatchWildcards = true;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"個別マーカー削除エラー: {ex.Message}");
                        // 個別のマーカー削除でエラーが発生しても処理を継続
                        break;
                    }
                }

                return removedCount;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"画像マーカー削除中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// マーカー範囲を拡張して、関連する改行も含めるように調整
        /// </summary>
        /// <param name="markerRange">マーカーの範囲</param>
        private static void ExtendRangeToIncludeAssociatedLineBreaks(Word.Range markerRange)
        {
            try
            {
                // マーカーの前の文字をチェック
                if (markerRange.Start > 0)
                {
                    var beforeRange = markerRange.Document.Range(markerRange.Start - 1, markerRange.Start);
                    if (beforeRange.Text == "\r" || beforeRange.Text == "\n")
                    {
                        // 前の改行も削除範囲に含める
                        markerRange.SetRange(markerRange.Start - 1, markerRange.End);
                    }
                }

                // マーカーの後の文字をチェック
                if (markerRange.End < markerRange.Document.Range().End)
                {
                    var afterRange = markerRange.Document.Range(markerRange.End, markerRange.End + 1);
                    if (afterRange.Text == "\r" || afterRange.Text == "\n")
                    {
                        // 後の改行も削除範囲に含める
                        markerRange.SetRange(markerRange.Start, markerRange.End + 1);

                        // 連続する改行もチェック
                        if (markerRange.End < markerRange.Document.Range().End)
                        {
                            var nextAfterRange = markerRange.Document.Range(markerRange.End, markerRange.End + 1);
                            if (nextAfterRange.Text == "\r" || nextAfterRange.Text == "\n")
                            {
                                markerRange.SetRange(markerRange.Start, markerRange.End + 1);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"範囲拡張エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 特定の画像マーカーのみを削除する
        /// </summary>
        /// <param name="document">マーカーを削除する対象のWordドキュメント</param>
        /// <param name="markerText">削除したいマーカーのテキスト（拡張子なしのファイル名）</param>
        /// <returns>削除されたマーカーが見つかったかどうか</returns>
        public static bool RemoveSpecificImageMarker(Word.Document document, string markerText)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(markerText))
                throw new ArgumentException("マーカーテキストが指定されていません。", nameof(markerText));

            try
            {
                // ドキュメント全体の範囲を取得
                var fullRange = document.Range();

                // 検索条件を設定
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = $"[IMAGEMARKER:{markerText}]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = false;

                // マーカーを検索
                if (find.Execute())
                {
                    // マーカーテキストの範囲を取得
                    var markerRange = fullRange.Duplicate;

                    // マーカーの前後の改行も含めて削除範囲を拡張
                    ExtendRangeToIncludeAssociatedLineBreaks(markerRange);

                    // マーカーとその前後の改行を削除
                    markerRange.Delete();

                    return true;
                }

                return false;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"画像マーカー削除中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// ドキュメント内の画像マーカーの数を取得する
        /// </summary>
        /// <param name="document">対象のWordドキュメント</param>
        /// <returns>見つかったマーカーの数</returns>
        public static int CountImageMarkers(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            int count = 0;

            try
            {
                // ドキュメント全体の範囲を取得
                var fullRange = document.Range();

                // 検索条件を設定
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;

                // マーカーを順次検索してカウント
                while (find.Execute())
                {
                    count++;

                    // 次の検索のために範囲を調整
                    fullRange.SetRange(fullRange.End, document.Range().End);
                    if (fullRange.Start >= fullRange.End)
                        break;
                }

                return count;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"画像マーカーカウントエラー: {ex.Message}");
                return 0;
            }
        }

        /// <summary>
        /// ドキュメント内の全ての画像マーカーの一覧を取得する
        /// </summary>
        /// <param name="document">対象のWordドキュメント</param>
        /// <returns>見つかったマーカーテキストのリスト</returns>
        public static List<string> GetImageMarkersList(Word.Document document)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            var markers = new List<string>();

            try
            {
                // ドキュメント全体の範囲を取得
                var fullRange = document.Range();

                // 検索条件を設定
                var find = fullRange.Find;
                find.ClearFormatting();
                find.Text = @"\[IMAGEMARKER:*\]";
                find.Forward = true;
                find.Wrap = Word.WdFindWrap.wdFindStop;
                find.Format = false;
                find.MatchCase = false;
                find.MatchWholeWord = false;
                find.MatchWildcards = true;

                // マーカーを順次検索して一覧に追加
                while (find.Execute())
                {
                    try
                    {
                        string markerText = fullRange.Text;
                        markers.Add(markerText);

                        // 次の検索のために範囲を調整
                        fullRange.SetRange(fullRange.End, document.Range().End);
                        if (fullRange.Start >= fullRange.End)
                            break;
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"個別マーカー取得エラー: {ex.Message}");
                        break;
                    }
                }

                return markers;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"画像マーカー一覧取得エラー: {ex.Message}");
                return new List<string>();
            }
        }
    }
}