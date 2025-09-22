using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 画像サイズ比較情報を格納するクラス
        /// </summary>
        public class ImageSizeComparison
        {
            public string FileName { get; set; }
            public string ImageType { get; set; }
            public int Position { get; set; }
            public float OriginalPointsWidth { get; set; }
            public float OriginalPointsHeight { get; set; }
            public int OriginalPixelsWidth { get; set; }
            public int OriginalPixelsHeight { get; set; }
            public int PngPixelsWidth { get; set; }
            public int PngPixelsHeight { get; set; }
            public double WidthRatio { get; set; }
            public double HeightRatio { get; set; }
            public double SizeRatio { get; set; }
            public string SizeChange { get; set; }
        }

        /// <summary>
        /// ポイントをピクセルに変換（96DPI基準）
        /// </summary>
        /// <param name="points">ポイント値</param>
        /// <returns>ピクセル値</returns>
        private static int ConvertPointsToPixels(float points)
        {
            // 1ポイント = 1.33333ピクセル（96DPI環境）
            return (int)Math.Round(points * 96f / 72f);
        }

        /// <summary>
        /// 元画像のサイズと抽出したPNG画像のサイズをピクセルで比較する一覧を取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>サイズ比較情報のリスト</returns>
        public static List<ImageSizeComparison> GetImageSizeComparisonList(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return new List<ImageSizeComparison>();

            var comparisons = new List<ImageSizeComparison>();

            foreach (var image in extractedImages.OrderBy(img => img.Position))
            {
                // 元画像サイズをピクセルに変換
                int originalPixelWidth = ConvertPointsToPixels(image.OriginalWidth);
                int originalPixelHeight = ConvertPointsToPixels(image.OriginalHeight);

                // 比率計算（PNG画像サイズが有効な場合のみ）
                double widthRatio = 0;
                double heightRatio = 0;
                double sizeRatio = 0;
                string sizeChange = "不明";

                if (image.PngPixelWidth > 0 && image.PngPixelHeight > 0 && originalPixelWidth > 0 && originalPixelHeight > 0)
                {
                    widthRatio = (double)image.PngPixelWidth / originalPixelWidth;
                    heightRatio = (double)image.PngPixelHeight / originalPixelHeight;
                    sizeRatio = (double)(image.PngPixelWidth * image.PngPixelHeight) / (originalPixelWidth * originalPixelHeight);

                    // サイズ変化の判定
                    if (Math.Abs(sizeRatio - 1.0) < 0.05) // 5%以内の差
                    {
                        sizeChange = "ほぼ同じ";
                    }
                    else if (sizeRatio > 1.0)
                    {
                        sizeChange = $"拡大 ({sizeRatio:F2}倍)";
                    }
                    else
                    {
                        sizeChange = $"縮小 ({sizeRatio:F2}倍)";
                    }
                }

                var comparison = new ImageSizeComparison
                {
                    FileName = Path.GetFileName(image.FilePath),
                    ImageType = image.ImageType,
                    Position = image.Position,
                    OriginalPointsWidth = image.OriginalWidth,
                    OriginalPointsHeight = image.OriginalHeight,
                    OriginalPixelsWidth = originalPixelWidth,
                    OriginalPixelsHeight = originalPixelHeight,
                    PngPixelsWidth = image.PngPixelWidth,
                    PngPixelsHeight = image.PngPixelHeight,
                    WidthRatio = widthRatio,
                    HeightRatio = heightRatio,
                    SizeRatio = sizeRatio,
                    SizeChange = sizeChange
                };

                comparisons.Add(comparison);
            }

            return comparisons;
        }

        /// <summary>
        /// 画像サイズ比較の統計情報を取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>比較統計情報の文字列</returns>
        public static string GetImageSizeComparisonStatistics(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return "抽出された画像はありません。";

            var comparisons = GetImageSizeComparisonList(extractedImages);
            var validComparisons = comparisons.Where(c => c.SizeRatio > 0).ToList();

            if (validComparisons.Count == 0)
                return "サイズ比較可能な画像がありません。";

            var statistics = new System.Text.StringBuilder();
            statistics.AppendLine("画像サイズ比較統計（元画像 vs PNG画像）");
            statistics.AppendLine("==========================================");
            statistics.AppendLine();

            // 拡大・縮小・同じサイズの分類
            var enlarged = validComparisons.Where(c => c.SizeRatio > 1.05).ToList();
            var reduced = validComparisons.Where(c => c.SizeRatio < 0.95).ToList();
            var similar = validComparisons.Where(c => c.SizeRatio >= 0.95 && c.SizeRatio <= 1.05).ToList();

            statistics.AppendLine($"総画像数: {comparisons.Count}");
            statistics.AppendLine($"比較可能: {validComparisons.Count}");
            statistics.AppendLine($"拡大された画像: {enlarged.Count}");
            statistics.AppendLine($"縮小された画像: {reduced.Count}");
            statistics.AppendLine($"ほぼ同サイズ: {similar.Count}");
            statistics.AppendLine();

            // サイズ比率の統計
            if (validComparisons.Any())
            {
                var avgRatio = validComparisons.Average(c => c.SizeRatio);
                var maxRatio = validComparisons.Max(c => c.SizeRatio);
                var minRatio = validComparisons.Min(c => c.SizeRatio);

                statistics.AppendLine("サイズ比率統計:");
                statistics.AppendLine($"平均比率: {avgRatio:F3}倍");
                statistics.AppendLine($"最大比率: {maxRatio:F3}倍");
                statistics.AppendLine($"最小比率: {minRatio:F3}倍");
                statistics.AppendLine();

                // 幅と高さの個別統計を追加
                var avgWidthRatio = validComparisons.Average(c => c.WidthRatio);
                var maxWidthRatio = validComparisons.Max(c => c.WidthRatio);
                var minWidthRatio = validComparisons.Min(c => c.WidthRatio);

                var avgHeightRatio = validComparisons.Average(c => c.HeightRatio);
                var maxHeightRatio = validComparisons.Max(c => c.HeightRatio);
                var minHeightRatio = validComparisons.Min(c => c.HeightRatio);

                statistics.AppendLine("幅比率統計:");
                statistics.AppendLine($"平均幅比率: {avgWidthRatio:F3}倍");
                statistics.AppendLine($"最大幅比率: {maxWidthRatio:F3}倍");
                statistics.AppendLine($"最小幅比率: {minWidthRatio:F3}倍");
                statistics.AppendLine();

                statistics.AppendLine("高さ比率統計:");
                statistics.AppendLine($"平均高さ比率: {avgHeightRatio:F3}倍");
                statistics.AppendLine($"最大高さ比率: {maxHeightRatio:F3}倍");
                statistics.AppendLine($"最小高さ比率: {minHeightRatio:F3}倍");
                statistics.AppendLine();
            }

            // 最も変化の大きい画像（上位5件）
            var mostChanged = validComparisons
                .OrderByDescending(c => Math.Abs(c.SizeRatio - 1.0))
                .Take(5)
                .ToList();

            if (mostChanged.Any())
            {
                statistics.AppendLine("最もサイズ変化の大きい画像（上位5件）:");
                foreach (var item in mostChanged)
                {
                    statistics.AppendLine($"  {item.FileName}: {item.SizeChange}");
                }
                statistics.AppendLine();
            }

            return statistics.ToString();
        }

        /// <summary>
        /// 幅と高さの比較に特化した詳細統計を取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>幅・高さ比較統計の文字列</returns>
        public static string GetWidthHeightComparisonStatistics(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return "抽出された画像はありません。";

            var comparisons = GetImageSizeComparisonList(extractedImages);
            var validComparisons = comparisons.Where(c => c.WidthRatio > 0 && c.HeightRatio > 0).ToList();

            if (validComparisons.Count == 0)
                return "比較可能な画像がありません。";

            var statistics = new System.Text.StringBuilder();
            statistics.AppendLine("画像サイズ比較統計（幅・高さ個別分析）");
            statistics.AppendLine("=====================================");
            statistics.AppendLine();

            // 幅の変化分類
            var widthEnlarged = validComparisons.Where(c => c.WidthRatio > 1.05).ToList();
            var widthReduced = validComparisons.Where(c => c.WidthRatio < 0.95).ToList();
            var widthSimilar = validComparisons.Where(c => c.WidthRatio >= 0.95 && c.WidthRatio <= 1.05).ToList();

            statistics.AppendLine("【幅の変化分類】");
            statistics.AppendLine($"拡大された画像: {widthEnlarged.Count}個");
            statistics.AppendLine($"縮小された画像: {widthReduced.Count}個");
            statistics.AppendLine($"ほぼ同じ幅: {widthSimilar.Count}個");
            statistics.AppendLine();

            // 高さの変化分類
            var heightEnlarged = validComparisons.Where(c => c.HeightRatio > 1.05).ToList();
            var heightReduced = validComparisons.Where(c => c.HeightRatio < 0.95).ToList();
            var heightSimilar = validComparisons.Where(c => c.HeightRatio >= 0.95 && c.HeightRatio <= 1.05).ToList();

            statistics.AppendLine("【高さの変化分類】");
            statistics.AppendLine($"拡大された画像: {heightEnlarged.Count}個");
            statistics.AppendLine($"縮小された画像: {heightReduced.Count}個");
            statistics.AppendLine($"ほぼ同じ高さ: {heightSimilar.Count}個");
            statistics.AppendLine();

            // 幅・高さの変化パターン分析
            var bothEnlarged = validComparisons.Where(c => c.WidthRatio > 1.05 && c.HeightRatio > 1.05).ToList();
            var bothReduced = validComparisons.Where(c => c.WidthRatio < 0.95 && c.HeightRatio < 0.95).ToList();
            var bothSimilar = validComparisons.Where(c => c.WidthRatio >= 0.95 && c.WidthRatio <= 1.05 && c.HeightRatio >= 0.95 && c.HeightRatio <= 1.05).ToList();
            var mixedChange = validComparisons.Where(c => 
                !(c.WidthRatio > 1.05 && c.HeightRatio > 1.05) && 
                !(c.WidthRatio < 0.95 && c.HeightRatio < 0.95) && 
                !(c.WidthRatio >= 0.95 && c.WidthRatio <= 1.05 && c.HeightRatio >= 0.95 && c.HeightRatio <= 1.05)).ToList();

            statistics.AppendLine("【変化パターン分析】");
            statistics.AppendLine($"幅・高さ両方拡大: {bothEnlarged.Count}個");
            statistics.AppendLine($"幅・高さ両方縮小: {bothReduced.Count}個");
            statistics.AppendLine($"幅・高さほぼ維持: {bothSimilar.Count}個");
            statistics.AppendLine($"幅・高さで異なる変化: {mixedChange.Count}個");
            statistics.AppendLine();

            // 最も幅が変化した画像（上位5件）
            var mostWidthChanged = validComparisons
                .OrderByDescending(c => Math.Abs(c.WidthRatio - 1.0))
                .Take(5)
                .ToList();

            if (mostWidthChanged.Any())
            {
                statistics.AppendLine("最も幅が変化した画像（上位5件）:");
                foreach (var item in mostWidthChanged)
                {
                    string widthChange = item.WidthRatio > 1.0 ? $"拡大 ({item.WidthRatio:F2}倍)" : $"縮小 ({item.WidthRatio:F2}倍)";
                    if (Math.Abs(item.WidthRatio - 1.0) < 0.05) widthChange = "ほぼ同じ";
                    statistics.AppendLine($"  {item.FileName}: {widthChange}");
                }
                statistics.AppendLine();
            }

            // 最も高さが変化した画像（上位5件）
            var mostHeightChanged = validComparisons
                .OrderByDescending(c => Math.Abs(c.HeightRatio - 1.0))
                .Take(5)
                .ToList();

            if (mostHeightChanged.Any())
            {
                statistics.AppendLine("最も高さが変化した画像（上位5件）:");
                foreach (var item in mostHeightChanged)
                {
                    string heightChange = item.HeightRatio > 1.0 ? $"拡大 ({item.HeightRatio:F2}倍)" : $"縮小 ({item.HeightRatio:F2}倍)";
                    if (Math.Abs(item.HeightRatio - 1.0) < 0.05) heightChange = "ほぼ同じ";
                    statistics.AppendLine($"  {item.FileName}: {heightChange}");
                }
                statistics.AppendLine();
            }

            return statistics.ToString();
        }

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

            // PNG画像ピクセルサイズ統計を追加
            if (extractedImages.Any(img => img.PngPixelWidth > 0 && img.PngPixelHeight > 0))
            {
                statistics.AppendLine();
                statistics.AppendLine("PNG画像ピクセルサイズ統計:");
                var avgPngWidth = extractedImages.Where(img => img.PngPixelWidth > 0).Average(img => img.PngPixelWidth);
                var avgPngHeight = extractedImages.Where(img => img.PngPixelHeight > 0).Average(img => img.PngPixelHeight);
                statistics.AppendLine($"平均サイズ: {avgPngWidth:F1} x {avgPngHeight:F1} pixels");

                var maxPngWidth = extractedImages.Where(img => img.PngPixelWidth > 0).Max(img => img.PngPixelWidth);
                var maxPngHeight = extractedImages.Where(img => img.PngPixelHeight > 0).Max(img => img.PngPixelHeight);
                statistics.AppendLine($"最大サイズ: {maxPngWidth} x {maxPngHeight} pixels");

                var minPngWidth = extractedImages.Where(img => img.PngPixelWidth > 0).Min(img => img.PngPixelWidth);
                var minPngHeight = extractedImages.Where(img => img.PngPixelHeight > 0).Min(img => img.PngPixelHeight);
                statistics.AppendLine($"最小サイズ: {minPngWidth} x {minPngHeight} pixels");
            }

            // サイズ比較統計を追加
            var comparisonStats = GetImageSizeComparisonStatistics(extractedImages);
            if (!comparisonStats.StartsWith("抽出された画像はありません") && 
                !comparisonStats.StartsWith("サイズ比較可能な画像がありません"))
            {
                statistics.AppendLine();
                statistics.AppendLine(comparisonStats);
            }

            // 幅・高さ個別比較統計を追加
            var widthHeightStats = GetWidthHeightComparisonStatistics(extractedImages);
            if (!widthHeightStats.StartsWith("抽出された画像はありません") && 
                !widthHeightStats.StartsWith("比較可能な画像がありません"))
            {
                statistics.AppendLine();
                statistics.AppendLine(widthHeightStats);
            }

            return statistics.ToString();
        }

        /// <summary>
        /// すべての画像の幅と高さを元画像と比較した完全な一覧をCSV形式で取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>すべての画像の幅・高さ比較一覧のCSV文字列</returns>
        public static string GetCompleteWidthHeightComparisonListAsCsv(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return "抽出された画像はありません。";

            var comparisons = GetImageSizeComparisonList(extractedImages);
            var result = new System.Text.StringBuilder();

            // CSVヘッダー
            result.AppendLine("位置,ファイル名,種別,元サイズ(points),元サイズ(pixels),PNGサイズ(pixels),幅比率,高さ比率,幅変化,高さ変化,変化パターン");

            foreach (var comparison in comparisons.OrderBy(c => c.Position))
            {
                string originalSizePoints = $"{comparison.OriginalPointsWidth:F1}x{comparison.OriginalPointsHeight:F1}";
                string originalSizePixels = $"{comparison.OriginalPixelsWidth}x{comparison.OriginalPixelsHeight}";
                string pngSizePixels = $"{comparison.PngPixelsWidth}x{comparison.PngPixelsHeight}";

                string widthRatioText = comparison.WidthRatio > 0 ? $"{comparison.WidthRatio:F3}" : "";
                string heightRatioText = comparison.HeightRatio > 0 ? $"{comparison.HeightRatio:F3}" : "";

                string widthChange = "";
                string heightChange = "";
                string changePattern = "";

                if (comparison.WidthRatio > 0 && comparison.HeightRatio > 0)
                {
                    // 幅の変化判定
                    if (Math.Abs(comparison.WidthRatio - 1.0) < 0.05)
                    {
                        widthChange = "維持";
                    }
                    else if (comparison.WidthRatio > 1.0)
                    {
                        widthChange = $"拡大({comparison.WidthRatio:F2}倍)";
                    }
                    else
                    {
                        widthChange = $"縮小({comparison.WidthRatio:F2}倍)";
                    }

                    // 高さの変化判定
                    if (Math.Abs(comparison.HeightRatio - 1.0) < 0.05)
                    {
                        heightChange = "維持";
                    }
                    else if (comparison.HeightRatio > 1.0)
                    {
                        heightChange = $"拡大({comparison.HeightRatio:F2}倍)";
                    }
                    else
                    {
                        heightChange = $"縮小({comparison.HeightRatio:F2}倍)";
                    }

                    // 変化パターンの判定
                    if (Math.Abs(comparison.WidthRatio - 1.0) < 0.05 && Math.Abs(comparison.HeightRatio - 1.0) < 0.05)
                    {
                        changePattern = "維持";
                    }
                    else if (comparison.WidthRatio > 1.05 && comparison.HeightRatio > 1.05)
                    {
                        changePattern = "全体拡大";
                    }
                    else if (comparison.WidthRatio < 0.95 && comparison.HeightRatio < 0.95)
                    {
                        changePattern = "全体縮小";
                    }
                    else if (Math.Abs(comparison.WidthRatio - comparison.HeightRatio) > 0.1)
                    {
                        changePattern = "アスペクト比変化";
                    }
                    else
                    {
                        changePattern = "複合変化";
                    }
                }

                // CSVではカンマが含まれる可能性があるフィールドをダブルクォートで囲む
                string fileNameCsv = $"\"{comparison.FileName}\"";
                string imageTypeCsv = $"\"{comparison.ImageType}\"";
                string widthChangeCsv = $"\"{widthChange}\"";
                string heightChangeCsv = $"\"{heightChange}\"";
                string changePatternCsv = $"\"{changePattern}\"";

                result.AppendLine($"{comparison.Position},{fileNameCsv},{imageTypeCsv},\"{originalSizePoints}\",\"{originalSizePixels}\",\"{pngSizePixels}\",{widthRatioText},{heightRatioText},{widthChangeCsv},{heightChangeCsv},{changePatternCsv}");
            }

            return result.ToString();
        }

        /// <summary>
        /// すべての画像の幅と高さ比較一覧をCSVファイルに出力
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        public static void ExportCompleteWidthHeightComparisonListToCsvFile(
            List<ExtractedImageInfo> extractedImages, string outputPath)
        {
            try
            {
                using (var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
                {
                    var csvText = GetCompleteWidthHeightComparisonListAsCsv(extractedImages);
                    writer.Write(csvText);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"幅・高さ比較CSVファイル出力エラー: {ex.Message}");
            }
        }
    }
}
