// Utils.ExtractImagesFromWord.cs

using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {
        // マジックナンバーの定数化
        private const int DefaultMinContentSizePixels = 250;
        private const int DefaultMaxOutputSizePixels = 1024;
        private const float DefaultMinOriginalSizePoints = 50.0f;
        private const int AlphaThresholdForTransparency = 0;
        private const int MinPixelSize = 1;

        /// <summary>
        /// 画像抽出オプション
        /// </summary>
        public class ImageExtractionOptions
        {
            public bool IncludeInlineShapes { get; set; } = true;
            public bool IncludeShapes { get; set; } = true;
            public bool IncludeFreeforms { get; set; } = true;
            public bool AddMarkers { get; set; } = true;
            public bool SkipCoverMarkers { get; set; } = true;
            public float MinOriginalWidth { get; set; } = DefaultMinOriginalSizePoints;
            public float MinOriginalHeight { get; set; } = DefaultMinOriginalSizePoints;
            public bool IncludeMjsTableImages { get; set; } = true;
            public int MaxOutputWidth { get; set; } = DefaultMaxOutputSizePixels;
            public int MaxOutputHeight { get; set; } = DefaultMaxOutputSizePixels;
        }

        /// <summary>
        /// 抽出された画像の情報
        /// </summary>
        public class ExtractedImageInfo
        {
            public string FilePath { get; set; }
            public string ImageType { get; set; }
            public int Position { get; set; }
            public float OriginalWidth { get; set; }
            public float OriginalHeight { get; set; }
            
            /// <summary>
            /// 抽出されたPNG画像の実際の幅（ピクセル単位）
            /// </summary>
            public int PngPixelWidth { get; set; }
            
            /// <summary>
            /// 抽出されたPNG画像の実際の高さ（ピクセル単位）
            /// </summary>
            public int PngPixelHeight { get; set; }

            public ExtractedImageInfo(
                string filePath, 
                string imageType, 
                int position,
                float originalWidth = 0,
                float originalHeight = 0,
                int pngPixelWidth = 0,
                int pngPixelHeight = 0)
            {
                FilePath = filePath;
                ImageType = imageType;
                Position = position;
                OriginalWidth = originalWidth;
                OriginalHeight = originalHeight;
                PngPixelWidth = pngPixelWidth;
                PngPixelHeight = pngPixelHeight;
            }
        }

        /// <summary>
        /// WordドキュメントからEnhMetaFileBitsを使用して画像とキャンバスを抽出する
        /// </summary>
        /// <param name="document">抽出対象のWordドキュメント</param>
        /// <param name="outputDirectory">画像の保存先ディレクトリ</param>
        /// <param name="options">画像抽出オプション（nullの場合はデフォルト値を使用）</param>
        /// <returns>抽出された画像情報のリスト</returns>
        public static List<ExtractedImageInfo> ExtractImagesFromWord(
            Word.Document document, 
            string outputDirectory,
            ImageExtractionOptions options = null)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(outputDirectory))
                throw new ArgumentException("出力ディレクトリが指定されていません。", nameof(outputDirectory));

            // オプションがnullの場合はデフォルトインスタンスを作成
            options = options ?? new ImageExtractionOptions();

            // 出力ディレクトリが存在しない場合は作成
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            // 環境情報をログ出力
            LogInfo($"Word Version: {Globals.ThisAddIn.Application.Version}");
            LogInfo($"Document Compatibility Mode: {document.CompatibilityMode}");
            LogInfo($"Sections Count: {document.Sections.Count}");
            LogInfo($"InlineShapes Count: {document.InlineShapes.Count}");
            LogInfo($"Shapes Count: {document.Shapes.Count}");

            var extractedImages = new List<ExtractedImageInfo>();
            int imageCounter = 1;

            try
            {
                // インライン図形の抽出
                if (options.IncludeInlineShapes)
                {
                    ExtractInlineShapes(document, outputDirectory, ref imageCounter, extractedImages, options);
                }

                // フローティング図形の抽出
                if (options.IncludeShapes)
                {
                    ExtractFloatingShapes(document, outputDirectory, ref imageCounter, extractedImages, options);
                }

                return extractedImages;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"画像抽出中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// ログ情報を出力
        /// </summary>
        private static void LogInfo(string message)
        {
            System.Diagnostics.Trace.WriteLine($"[ExtractImages] {message}");
        }

        /// <summary>
        /// ログエラーを出力
        /// </summary>
        private static void LogError(string message, Exception ex = null)
        {
            System.Diagnostics.Trace.WriteLine($"[ExtractImages ERROR] {message}" + 
                (ex != null ? $": {ex.Message}" : ""));
        }
    }
}
