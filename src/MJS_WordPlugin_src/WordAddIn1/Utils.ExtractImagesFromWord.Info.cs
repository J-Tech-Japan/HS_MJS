// Utils.ExtractImagesFromWord.Info.cs

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
            public long OriginalTotalPixels { get; set; }
            public int PngPixelsWidth { get; set; }
            public int PngPixelsHeight { get; set; }
            public long PngTotalPixels { get; set; }
            public long PngFileSize { get; set; }
            public double BitsPerPixel { get; set; }
            public double CompressionRatio { get; set; }
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
                long originalTotalPixels = (long)originalPixelWidth * originalPixelHeight;

                // PNG画像の総ピクセル数を計算
                long pngTotalPixels = (long)image.PngPixelWidth * image.PngPixelHeight;

                // PNGファイルサイズを取得
                long pngFileSize = 0;
                try
                {
                    if (!string.IsNullOrEmpty(image.FilePath) && File.Exists(image.FilePath))
                    {
                        var fileInfo = new FileInfo(image.FilePath);
                        pngFileSize = fileInfo.Length;
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"ファイルサイズ取得エラー ({image.FilePath}): {ex.Message}");
                }

                // Bits/Pixel を計算
                // bpp = (ファイルサイズ × 8) / 総ピクセル数
                double bitsPerPixel = 0;
                if (pngTotalPixels > 0 && pngFileSize > 0)
                {
                    bitsPerPixel = ((double)pngFileSize * 8) / pngTotalPixels;
                }

                // 圧縮率を計算
                // PNG 24bit RGB (無圧縮想定: 24 bits/pixel) に対する実際のbppの比率
                // 値が小さいほど高圧縮、大きいほど低圧縮
                double compressionRatio = 0;
                if (bitsPerPixel > 0)
                {
                    compressionRatio = bitsPerPixel / 24.0;
                }

                // 比率計算（PNG画像サイズが有効な場合のみ）
                double widthRatio = 0;
                double heightRatio = 0;
                double sizeRatio = 0;
                string sizeChange = "不明";

                var comparison = new ImageSizeComparison
                {
                    FileName = Path.GetFileName(image.FilePath),
                    ImageType = image.ImageType,
                    Position = image.Position,
                    OriginalPointsWidth = image.OriginalWidth,
                    OriginalPointsHeight = image.OriginalHeight,
                    OriginalPixelsWidth = originalPixelWidth,
                    OriginalPixelsHeight = originalPixelHeight,
                    OriginalTotalPixels = originalTotalPixels,
                    PngPixelsWidth = image.PngPixelWidth,
                    PngPixelsHeight = image.PngPixelHeight,
                    PngTotalPixels = pngTotalPixels,
                    PngFileSize = pngFileSize,
                    BitsPerPixel = bitsPerPixel,
                    CompressionRatio = compressionRatio,
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
            result.AppendLine("位置,ファイル名,種別,元サイズ(px),元総ピクセル数,出力サイズ(px),出力総ピクセル数,出力ファイルサイズ(bytes),BPP,圧縮率,幅比率,高さ比率");

            foreach (var comparison in comparisons.OrderBy(c => c.Position))
            {
                string originalSizePixels = $"{comparison.OriginalPixelsWidth}x{comparison.OriginalPixelsHeight}";
                string pngSizePixels = $"{comparison.PngPixelsWidth}x{comparison.PngPixelsHeight}";

                string bitsPerPixelText = comparison.BitsPerPixel > 0 ? $"{comparison.BitsPerPixel:F2}" : "";
                string compressionRatioText = comparison.CompressionRatio > 0 ? $"{comparison.CompressionRatio:F3}" : "";
                string widthRatioText = comparison.WidthRatio > 0 ? $"{comparison.WidthRatio:F3}" : "";
                string heightRatioText = comparison.HeightRatio > 0 ? $"{comparison.HeightRatio:F3}" : "";

                // CSVではカンマが含まれる可能性があるフィールドをダブルクォートで囲む
                string fileNameCsv = $"\"{comparison.FileName}\"";
                string imageTypeCsv = $"\"{comparison.ImageType}\"";

                result.AppendLine($"{comparison.Position},{fileNameCsv},{imageTypeCsv},\"{originalSizePixels}\",{comparison.OriginalTotalPixels},\"{pngSizePixels}\",{comparison.PngTotalPixels},{comparison.PngFileSize},{bitsPerPixelText},{compressionRatioText},{widthRatioText},{heightRatioText}");
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
