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
            System.Diagnostics.Trace.WriteLine($"[InsertMarker] 開始: filePath={filePath}");
            
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] マーカーテキスト生成: {markerText}");

                // 図形を含む段落を取得
                var paragraph = range.Paragraphs[1];
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] 段落取得完了: Start={paragraph.Range.Start}, End={paragraph.Range.End}");

                // 段落の末尾に移動
                var insertRange = range.Document.Range(paragraph.Range.End - 1, paragraph.Range.End - 1);
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] 挿入位置設定: Start={insertRange.Start}, End={insertRange.End}");

                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";
                System.Diagnostics.Trace.WriteLine("[InsertMarker] 改行挿入完了");

                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = range.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] マーカー挿入完了: {marker}");

                // マーカーの後に改行を追加
                var afterMarkerRange = range.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
                System.Diagnostics.Trace.WriteLine("[InsertMarker] 後続改行挿入完了");
                
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] 処理完了: {marker}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] エラー発生: {ex.Message}");
                System.Diagnostics.Trace.WriteLine($"[InsertMarker] スタックトレース: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// 指定した位置の次の行にマーカーを挿入（フローティング図形用）
        /// </summary>
        private static void InsertMarkerAtPosition(Word.Range anchor, string filePath)
        {
            System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] 開始: filePath={filePath}");
            
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] マーカーテキスト生成: {markerText}");

                // アンカー位置を含む段落を取得
                var anchorParagraph = anchor.Paragraphs[1];
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] アンカー段落取得完了: Start={anchorParagraph.Range.Start}, End={anchorParagraph.Range.End}");

                // 段落の末尾に移動
                var insertRange = anchor.Document.Range(anchorParagraph.Range.End - 1, anchorParagraph.Range.End - 1);
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] 挿入位置設定: Start={insertRange.Start}, End={insertRange.End}");

                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";
                System.Diagnostics.Trace.WriteLine("[InsertMarkerAtPosition] 改行挿入完了");

                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = anchor.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] マーカー挿入完了: {marker}");

                // マーカーの後に改行を追加
                var afterMarkerRange = anchor.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
                System.Diagnostics.Trace.WriteLine("[InsertMarkerAtPosition] 後続改行挿入完了");
                
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] 処理完了: {marker}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] エラー発生: {ex.Message}");
                System.Diagnostics.Trace.WriteLine($"[InsertMarkerAtPosition] スタックトレース: {ex.StackTrace}");
            }
        }
    }
}
