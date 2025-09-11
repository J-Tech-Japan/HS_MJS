using System;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 指定されたスタイル名が特定のMJSスタイルかどうかを判定
        /// </summary>
        /// <param name="styleName">判定対象のスタイル名</param>
        /// <param name="forceExtract">強制抽出フラグ（出力）</param>
        /// <param name="forceSkip">強制スキップフラグ（出力）</param>
        /// <param name="includeMjsTableImages">MJS_画像（表内）スタイルを含めるかどうか</param>
        private static void CheckMjsStyleConditions(string styleName, out bool forceExtract, out bool forceSkip, bool includeMjsTableImages = true)
        {
            forceExtract = false;
            forceSkip = false;

            if (string.IsNullOrEmpty(styleName))
                return;

            // MJS_画像（表内）の特別処理
            if (styleName.Contains("MJS_画像（表内）"))
            {
                if (includeMjsTableImages)
                {
                    forceExtract = true;
                    System.Diagnostics.Debug.WriteLine($"スタイル '{styleName}' により強制抽出対象に設定（MJS表内画像許可）");
                }
                else
                {
                    forceSkip = true;
                    System.Diagnostics.Debug.WriteLine($"スタイル '{styleName}' により強制スキップ対象に設定（MJS表内画像除外）");
                }
                return;
            }

            // その他の強制抽出対象のスタイル（サイズに関わりなく必ず抽出）
            if (styleName.Contains("MJS_画像（手順内）") ||
                styleName.Contains("MJS_画像（本文内）") ||
                styleName.Contains("MJS_画像（コラム内）"))
            {
                forceExtract = true;
                System.Diagnostics.Debug.WriteLine($"スタイル '{styleName}' により強制抽出対象に設定");
                return;
            }

            // 強制スキップ対象のスタイル（サイズに関わりなく抽出しない）
            if (styleName.Contains("MJS_処理フロー") || styleName.Contains("MJS_表内-項目_センタリング"))
            {
                forceSkip = true;
                System.Diagnostics.Debug.WriteLine($"スタイル '{styleName}' により強制スキップ対象に設定");
                return;
            }
        }

        /// <summary>
        /// インライン図形を含む段落のスタイルを取得
        /// </summary>
        /// <param name="inlineShape">インライン図形</param>
        /// <returns>段落のスタイル名</returns>
        private static string GetInlineShapeParagraphStyle(Word.InlineShape inlineShape)
        {
            try
            {
                var paragraph = inlineShape.Range.Paragraphs[1];
                return paragraph.get_Style().NameLocal;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"インライン図形の段落スタイル取得エラー: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// フローティング図形を含む段落のスタイルを取得
        /// </summary>
        /// <param name="shape">フローティング図形</param>
        /// <returns>段落のスタイル名</returns>
        private static string GetShapeAnchorParagraphStyle(Word.Shape shape)
        {
            try
            {
                if (shape.Anchor != null)
                {
                    var paragraph = shape.Anchor.Paragraphs[1];
                    return paragraph.get_Style().NameLocal;
                }
                return string.Empty;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"フローティング図形のアンカー段落スタイル取得エラー: {ex.Message}");
                return string.Empty;
            }
        }

        /// <summary>
        /// 指定されたRangeが表紙（第1セクション）にあるかどうかを判定
        /// </summary>
        /// <param name="range">判定対象のRange</param>
        /// <param name="skipCoverMarkers">表紙マーカーをスキップするかどうか</param>
        /// <returns>表紙にある場合はtrue</returns>
        private static bool IsInCoverSection(Word.Range range, bool skipCoverMarkers)
        {
            if (!skipCoverMarkers)
                return false;

            try
            {
                // Rangeが属するセクション番号を取得
                int sectionNumber = range.Information[Word.WdInformation.wdActiveEndSectionNumber];
                return sectionNumber == 1;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"セクション番号の取得でエラー: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 指定されたShapeが表紙（第1セクション）にあるかどうかを判定
        /// </summary>
        /// <param name="shape">判定対象のShape</param>
        /// <param name="skipCoverMarkers">表紙マーカーをスキップするかどうか</param>
        /// <returns>表紙にある場合はtrue</returns>
        private static bool IsShapeInCoverSection(Word.Shape shape, bool skipCoverMarkers)
        {
            if (!skipCoverMarkers)
                return false;

            try
            {
                // Shapeを選択してセクション番号を取得
                shape.Select();
                var selection = Globals.ThisAddIn.Application.Selection;
                int sectionNumber = selection.Information[Word.WdInformation.wdActiveEndSectionNumber];
                return sectionNumber == 1;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"図形のセクション番号取得でエラー: {ex.Message}");
                return false;
            }
        }
    }
}
