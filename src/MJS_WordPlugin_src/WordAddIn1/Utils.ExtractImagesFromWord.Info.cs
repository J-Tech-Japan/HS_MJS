using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 抽出結果の統計情報を取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>統計情報の文字列</returns>
        public static string GetExtractionStatisticsWithText(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return "抽出された画像はありません。";

            var statistics = new System.Text.StringBuilder();
            statistics.AppendLine($"抽出された画像数: {extractedImages.Count}");
            statistics.AppendLine();

            var groupedByType = extractedImages
                .GroupBy(img => img.ImageType)
                .OrderBy(g => g.Key);

            foreach (var group in groupedByType)
            {
                statistics.AppendLine($"{group.Key}: {group.Count()}個");
            }

            // 元サイズ統計を追加
            if (extractedImages.Any(img => img.OriginalWidth > 0 && img.OriginalHeight > 0))
            {
                statistics.AppendLine();
                statistics.AppendLine("元画像サイズ統計:");
                var avgWidth = extractedImages.Where(img => img.OriginalWidth > 0).Average(img => img.OriginalWidth);
                var avgHeight = extractedImages.Where(img => img.OriginalHeight > 0).Average(img => img.OriginalHeight);
                statistics.AppendLine($"平均サイズ: {avgWidth:F1} x {avgHeight:F1} points");

                var maxWidth = extractedImages.Where(img => img.OriginalWidth > 0).Max(img => img.OriginalWidth);
                var maxHeight = extractedImages.Where(img => img.OriginalHeight > 0).Max(img => img.OriginalHeight);
                statistics.AppendLine($"最大サイズ: {maxWidth:F1} x {maxHeight:F1} points");

                var minWidth = extractedImages.Where(img => img.OriginalWidth > 0).Min(img => img.OriginalWidth);
                var minHeight = extractedImages.Where(img => img.OriginalHeight > 0).Min(img => img.OriginalHeight);
                statistics.AppendLine($"最小サイズ: {minWidth:F1} x {minHeight:F1} points");
            }

            return statistics.ToString();
        }

        /// <summary>
        /// 抽出結果をテキストファイルに出力
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        public static void ExportImageInfoToTextFile(
            List<ExtractedImageInfo> extractedImages, string outputPath)
        {
            try
            {
                using (var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("抽出された画像の一覧");
                    writer.WriteLine("==================");
                    writer.WriteLine();

                    foreach (var image in extractedImages.OrderBy(img => img.Position))
                    {
                        writer.WriteLine($"ファイル名: {Path.GetFileName(image.FilePath)}");
                        writer.WriteLine($"種別: {image.ImageType}");
                        writer.WriteLine($"位置: {image.Position}");
                        if (image.OriginalWidth > 0 && image.OriginalHeight > 0)
                        {
                            writer.WriteLine($"元サイズ: {image.OriginalWidth:F1} x {image.OriginalHeight:F1} points");
                        }
                        writer.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"テキストファイル出力エラー: {ex.Message}");
            }
        }
    }
}
