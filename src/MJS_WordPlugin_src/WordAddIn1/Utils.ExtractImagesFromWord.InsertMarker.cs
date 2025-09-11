using System;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// インライン図形の次の行にマーカーテキストを挿入
        /// </summary>
        private static void InsertMarker(Word.Range range, string filePath)
        {
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);

                // 図形を含む段落を取得
                var paragraph = range.Paragraphs[1];

                // 段落の末尾に移動
                var insertRange = range.Document.Range(paragraph.Range.End - 1, paragraph.Range.End - 1);

                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";

                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = range.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;

                // マーカーの後に改行を追加
                var afterMarkerRange = range.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"マーカー挿入エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 指定した位置の次の行にマーカーを挿入（フローティング図形用）
        /// </summary>
        private static void InsertMarkerAtPosition(Word.Range anchor, string filePath)
        {
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);

                // アンカー位置を含む段落を取得
                var anchorParagraph = anchor.Paragraphs[1];

                // 段落の末尾に移動
                var insertRange = anchor.Document.Range(anchorParagraph.Range.End - 1, anchorParagraph.Range.End - 1);

                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";

                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = anchor.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;

                // マーカーの後に改行を追加
                var afterMarkerRange = anchor.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"マーカー挿入エラー: {ex.Message}");
            }
        }
    }
}
