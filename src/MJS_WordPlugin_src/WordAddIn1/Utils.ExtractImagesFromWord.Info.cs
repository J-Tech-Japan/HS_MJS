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
            result.AppendLine("位置,ファイル名,種別,元サイズ(points),元サイズ(pixels),PNGサイズ(pixels),幅比率,高さ比率");

            foreach (var comparison in comparisons.OrderBy(c => c.Position))
            {
                string originalSizePoints = $"{comparison.OriginalPointsWidth:F1}x{comparison.OriginalPointsHeight:F1}";
                string originalSizePixels = $"{comparison.OriginalPixelsWidth}x{comparison.OriginalPixelsHeight}";
                string pngSizePixels = $"{comparison.PngPixelsWidth}x{comparison.PngPixelsHeight}";

                string widthRatioText = comparison.WidthRatio > 0 ? $"{comparison.WidthRatio:F3}" : "";
                string heightRatioText = comparison.HeightRatio > 0 ? $"{comparison.HeightRatio:F3}" : "";

                // CSVではカンマが含まれる可能性があるフィールドをダブルクォートで囲む
                string fileNameCsv = $"\"{comparison.FileName}\"";
                string imageTypeCsv = $"\"{comparison.ImageType}\"";

                result.AppendLine($"{comparison.Position},{fileNameCsv},{imageTypeCsv},\"{originalSizePoints}\",\"{originalSizePixels}\",\"{pngSizePixels}\",{widthRatioText},{heightRatioText}");
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
