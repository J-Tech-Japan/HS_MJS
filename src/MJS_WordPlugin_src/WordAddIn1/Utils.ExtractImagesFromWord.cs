// Utils.ExtractImagesFromWord.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
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
        /// 旧シグネチャとの互換性のためのオーバーロード
        /// </summary>
        [Obsolete("ImageExtractionOptionsを使用するオーバーロードを使用してください")]
        public static List<ExtractedImageInfo> ExtractImagesFromWord(
            Word.Document document, 
            string outputDirectory,
            bool includeInlineShapes = true,
            bool includeShapes = true,
            bool includeFreeforms = true,
            bool addMarkers = true,
            bool skipCoverMarkers = true,
            float minOriginalWidth = DefaultMinOriginalSizePoints,
            float minOriginalHeight = DefaultMinOriginalSizePoints,
            bool includeMjsTableImages = true,
            int maxOutputWidth = DefaultMaxOutputSizePixels,
            int maxOutputHeight = DefaultMaxOutputSizePixels)
        {
            var options = new ImageExtractionOptions
            {
                IncludeInlineShapes = includeInlineShapes,
                IncludeShapes = includeShapes,
                IncludeFreeforms = includeFreeforms,
                AddMarkers = addMarkers,
                SkipCoverMarkers = skipCoverMarkers,
                MinOriginalWidth = minOriginalWidth,
                MinOriginalHeight = minOriginalHeight,
                IncludeMjsTableImages = includeMjsTableImages,
                MaxOutputWidth = maxOutputWidth,
                MaxOutputHeight = maxOutputHeight
            };

            return ExtractImagesFromWord(document, outputDirectory, options);
        }
        
        /// <summary>
        /// インライン図形からEnhMetaFileBitsを使用して画像を抽出
        /// </summary>
        private static void ExtractInlineShapes(
            Word.Document document, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages,
            ImageExtractionOptions options)
        {
            foreach (Word.InlineShape inlineShape in document.InlineShapes)
            {
                try
                {
                    // 段落のスタイルを取得
                    string paragraphStyle = GetInlineShapeParagraphStyle(inlineShape);
                    
                    // MJSスタイルによる条件チェック
                    CheckMjsStyleConditions(paragraphStyle, out bool forceExtract, out bool forceSkip, options.IncludeMjsTableImages);
                    
                    // 強制スキップ対象の場合
                    if (forceSkip)
                    {
                        LogInfo($"インライン図形をスキップ: スタイル '{paragraphStyle}' により強制スキップ");
                        continue;
                    }

                    // 元画像サイズでのフィルタリング（強制抽出の場合はスキップ）
                    float originalWidth = inlineShape.Width;
                    float originalHeight = inlineShape.Height;
                    
                    if (!forceExtract && (originalWidth < options.MinOriginalWidth || originalHeight < options.MinOriginalHeight))
                    {
                        LogInfo($"インライン図形をスキップ: 元サイズが小さすぎます ({originalWidth:F1}x{originalHeight:F1} points)");
                        continue;
                    }

                    // EnhMetaFileBitsを取得
                    byte[] metaFileData = (byte[])inlineShape.Range.EnhMetaFileBits;
                    
                    if (metaFileData != null && metaFileData.Length > 0)
                    {
                        var extractResult = ExtractImageFromMetaFileDataWithSize(
                            metaFileData, 
                            outputDirectory, 
                            $"inline_image_{imageCounter}", 
                            inlineShape.Type.ToString(),
                            forceExtract,
                            options.MaxOutputWidth,
                            options.MaxOutputHeight);
                        
                        if (extractResult != null)
                        {
                            var imageInfo = new ExtractedImageInfo(
                                extractResult.FilePath, 
                                $"インライン図形_{inlineShape.Type}", 
                                inlineShape.Range.Start,
                                originalWidth,
                                originalHeight,
                                extractResult.PixelWidth,
                                extractResult.PixelHeight
                            );
                            extractedImages.Add(imageInfo);

                            // マーカーを追加（表紙の画像は除外）
                            if (options.AddMarkers && !IsInCoverSection(inlineShape.Range, options.SkipCoverMarkers))
                            {
                                InsertMarker(inlineShape.Range, extractResult.FilePath);
                            }

                            imageCounter++;
                        }
                    }
                }
                catch (Exception ex)
                {
                    // 個別の図形でエラーが発生しても処理を継続
                    LogError($"インライン図形の抽出でエラー", ex);
                }
            }
        }

        /// <summary>
        /// フローティング図形からEnhMetaFileBitsを使用して画像を抽出
        /// </summary>
        private static void ExtractFloatingShapes(
            Word.Document document, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages,
            ImageExtractionOptions options)
        {
            foreach (Word.Shape shape in document.Shapes)
            {
                try
                {
                    // フリーフォーム図形の場合
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
                    {
                        if (options.IncludeFreeforms)
                        {
                            ExtractSingleShape(shape, outputDirectory, ref imageCounter, extractedImages, "freeform", options);
                        }
                    }
                    // 通常の図形の場合
                    else
                    {
                        ExtractSingleShape(shape, outputDirectory, ref imageCounter, extractedImages, "shape", options);
                    }
                }
                catch (Exception ex)
                {
                    // 個別の図形でエラーが発生しても処理を継続
                    LogError($"フローティング図形の抽出でエラー", ex);
                }
            }
        }

        /// <summary>
        /// 単一の図形から画像を抽出
        /// </summary>
        private static void ExtractSingleShape(
            Word.Shape shape, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages, 
            string prefix,
            ImageExtractionOptions options)
        {
            try
            {
                // アンカー段落のスタイルを取得
                string anchorParagraphStyle = GetShapeAnchorParagraphStyle(shape);
                
                // MJSスタイルによる条件チェック
                CheckMjsStyleConditions(anchorParagraphStyle, out bool forceExtract, out bool forceSkip, options.IncludeMjsTableImages);
                
                // 強制スキップ対象の場合
                if (forceSkip)
                {
                    LogInfo($"{prefix}図形をスキップ: スタイル '{anchorParagraphStyle}' により強制スキップ");
                    return;
                }

                // 元画像サイズでのフィルタリング（強制抽出の場合はスキップ）
                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                
                if (!forceExtract && (originalWidth < options.MinOriginalWidth || originalHeight < options.MinOriginalHeight))
                {
                    LogInfo($"{prefix}図形をスキップ: 元サイズが小さすぎます ({originalWidth:F1}x{originalHeight:F1} points)");
                    return;
                }

                shape.Select();
                byte[] shapeData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                
                if (shapeData != null && shapeData.Length > 0)
                {
                    var extractResult = ExtractImageFromMetaFileDataWithSize(
                        shapeData, 
                        outputDirectory, 
                        $"{prefix}_{imageCounter}", 
                        shape.Type.ToString(),
                        forceExtract,
                        options.MaxOutputWidth,
                        options.MaxOutputHeight);
                    
                    if (extractResult != null)
                    {
                        var imageInfo = new ExtractedImageInfo(
                            extractResult.FilePath, 
                            $"{prefix}_{shape.Type}", 
                            shape.Anchor?.Start ?? 0,
                            originalWidth,
                            originalHeight,
                            extractResult.PixelWidth,
                            extractResult.PixelHeight
                        );
                        extractedImages.Add(imageInfo);

                        // マーカーを追加（表紙の画像は除外）
                        if (options.AddMarkers)
                        {
                            bool inCoverSection = shape.Anchor != null ? IsShapeInCoverSection(shape, options.SkipCoverMarkers) : false;
                            LogInfo($"[ExtractSingleShape] Anchor: {shape.Anchor != null}, InCover: {inCoverSection}");
                            
                            if (shape.Anchor != null && !inCoverSection)
                            {
                                InsertMarkerAtPosition(shape.Anchor, extractResult.FilePath);
                            }
                            else if (shape.Anchor == null)
                            {
                                LogInfo("[ExtractSingleShape] Anchorが取得できないためマーカーをスキップ");
                            }
                        }

                        imageCounter++;
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"図形の抽出でエラー", ex);
            }
        }
        
        /// <summary>
        /// 画像抽出結果を格納するクラス
        /// </summary>
        private class ImageExtractionResult
        {
            public string FilePath { get; set; }
            public int PixelWidth { get; set; }
            public int PixelHeight { get; set; }
        }
        
        /// <summary>
        /// EnhMetaFileBitsから画像ファイルを作成し、PNG画像のサイズも取得
        /// </summary>
        /// <param name="metaFileData">メタファイルデータ</param>
        /// <param name="outputDirectory">出力ディレクトリ</param>
        /// <param name="baseFileName">ベースファイル名</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="forceExtract">強制抽出フラグ</param>
        /// <param name="maxWidth">最大幅（ピクセル、デフォルト: 1024）</param>
        /// <param name="maxHeight">最大高さ（ピクセル、デフォルト: 1024）</param>
        /// <returns>作成されたファイルのパスとピクセルサイズ、失敗時はnull</returns>
        private static ImageExtractionResult ExtractImageFromMetaFileDataWithSize(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType,
            bool forceExtract = false,
            int maxWidth = DefaultMaxOutputSizePixels,
            int maxHeight = DefaultMaxOutputSizePixels)
        {
            try
            {
                using (var memoryStream = new MemoryStream(metaFileData))
                {
                    using (var metafile = new System.Drawing.Imaging.Metafile(memoryStream))
                    {
                        // メタファイルの実際のコンテンツ境界を取得
                        RectangleF bounds;
                        using (var graphics = Graphics.FromImage(new Bitmap(1, 1)))
                        {
                            GraphicsUnit unit = GraphicsUnit.Pixel;
                            bounds = metafile.GetBounds(ref unit);
                        }

                        // 境界が有効かチェック
                        if (bounds.Width <= 0 || bounds.Height <= 0)
                        {
                            LogInfo("メタファイルの境界が無効です");
                            return null;
                        }

                        LogInfo($"メタファイル境界: X={bounds.X}, Y={bounds.Y}, Width={bounds.Width}, Height={bounds.Height}");

                        // 実際のコンテンツサイズ（余白なし）
                        int contentWidth = (int)Math.Ceiling(bounds.Width);
                        int contentHeight = (int)Math.Ceiling(bounds.Height);

                        // 最小サイズのフィルタリング（強制抽出の場合はスキップ）
                        if (!forceExtract && (contentWidth < DefaultMinContentSizePixels || contentHeight < DefaultMinContentSizePixels))
                            return null;

                        // リサイズが必要かチェック
                        bool needsResize = contentWidth > maxWidth || contentHeight > maxHeight;
                        int finalWidth = contentWidth;
                        int finalHeight = contentHeight;

                        if (needsResize)
                        {
                            // 縦横比を維持してリサイズサイズを計算
                            var newSize = CalculateResizedDimensions(contentWidth, contentHeight, maxWidth, maxHeight);
                            finalWidth = newSize.Width;
                            finalHeight = newSize.Height;
                            
                            LogInfo($"画像をリサイズします: {contentWidth}x{contentHeight} → {finalWidth}x{finalHeight}");
                        }

                        // 余白なしで実際のコンテンツのみを含むビットマップを作成
                        using (var bitmap = new Bitmap(finalWidth, finalHeight, PixelFormat.Format32bppArgb))
                        {
                            using (var graphics = Graphics.FromImage(bitmap))
                            {
                                // 高品質な描画設定
                                graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
                                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                                // 背景を透明に設定
                                graphics.Clear(Color.Transparent);

                                // メタファイルの実際のコンテンツ領域のみを描画
                                // 境界のオフセットを考慮して、コンテンツのみを抽出
                                var destRect = new RectangleF(0, 0, finalWidth, finalHeight);
                                graphics.DrawImage(metafile, destRect, bounds, GraphicsUnit.Pixel);
                            }

                            // 透明ピクセルを除去してコンテンツのみの境界を取得
                            var trimmedBounds = GetTrimmedBounds(bitmap);
                            
                            if (trimmedBounds.Width <= 0 || trimmedBounds.Height <= 0)
                            {
                                LogInfo("トリミング後の境界が無効です");
                                return null;
                            }

                            // トリミングされた画像を作成
                            using (var trimmedBitmap = new Bitmap(trimmedBounds.Width, trimmedBounds.Height, PixelFormat.Format32bppArgb))
                            {
                                using (var graphics = Graphics.FromImage(trimmedBitmap))
                                {
                                    graphics.Clear(Color.White);
                                    graphics.DrawImage(bitmap, 
                                        new Rectangle(0, 0, trimmedBounds.Width, trimmedBounds.Height),
                                        trimmedBounds,
                                        GraphicsUnit.Pixel);
                                }

                                // ファイル名の生成
                                string fileName = $"{baseFileName}_{shapeType}.png";
                                string filePath = Path.Combine(outputDirectory, fileName);

                                // 重複ファイル名の回避
                                int duplicateCounter = 1;
                                while (File.Exists(filePath))
                                {
                                    fileName = $"{baseFileName}_{shapeType}_{duplicateCounter}.png";
                                    filePath = Path.Combine(outputDirectory, fileName);
                                    duplicateCounter++;
                                }

                                // PNG形式で保存
                                trimmedBitmap.Save(filePath, ImageFormat.Png);
                                
                                LogInfo($"画像を保存しました: {filePath} ({trimmedBounds.Width}x{trimmedBounds.Height})");
                                
                                return new ImageExtractionResult
                                {
                                    FilePath = filePath,
                                    PixelWidth = trimmedBounds.Width,
                                    PixelHeight = trimmedBounds.Height
                                };
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                LogError($"メタファイルデータからの画像生成でエラー", ex);
                return null;
            }
        }

        /// <summary>
        /// ビットマップから透明ピクセルを除いた実際のコンテンツ境界を取得
        /// </summary>
        /// <param name="bitmap">ビットマップ</param>
        /// <returns>コンテンツの境界矩形</returns>
        private static Rectangle GetTrimmedBounds(Bitmap bitmap)
        {
            int minX = bitmap.Width;
            int minY = bitmap.Height;
            int maxX = 0;
            int maxY = 0;

            // すべてのピクセルをスキャンして、透明でないピクセルの範囲を取得
            for (int y = 0; y < bitmap.Height; y++)
            {
                for (int x = 0; x < bitmap.Width; x++)
                {
                    Color pixel = bitmap.GetPixel(x, y);
                    // 完全に透明でないピクセルを検出
                    if (pixel.A > AlphaThresholdForTransparency)
                    {
                        if (x < minX) minX = x;
                        if (x > maxX) maxX = x;
                        if (y < minY) minY = y;
                        if (y > maxY) maxY = y;
                    }
                }
            }

            // コンテンツが見つからなかった場合
            if (minX > maxX || minY > maxY)
            {
                return Rectangle.Empty;
            }

            // 境界矩形を返す（幅と高さは+1して含める）
            return new Rectangle(minX, minY, maxX - minX + 1, maxY - minY + 1);
        }

        /// <summary>
        /// EnhMetaFileBitsから画像ファイルを作成（旧版、互換性のため保持）
        /// </summary>
        /// <param name="metaFileData">メタファイルデータ</param>
        /// <param name="outputDirectory">出力ディレクトリ</param>
        /// <param name="baseFileName">ベースファイル名</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="forceExtract">強制抽出フラグ</param>
        /// <param name="maxWidth">最大幅（ピクセル、デフォルト: 1024）</param>
        /// <param name="maxHeight">最大高さ（ピクセル、デフォルト: 1024）</param>
        /// <returns>作成されたファイルのパス、失敗時はnull</returns>
        [Obsolete("ExtractImageFromMetaFileDataWithSizeを使用してください")]
        private static string ExtractImageFromMetaFileData(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType,
            bool forceExtract = false,
            int maxWidth = DefaultMaxOutputSizePixels,
            int maxHeight = DefaultMaxOutputSizePixels)
        {
            var result = ExtractImageFromMetaFileDataWithSize(metaFileData, outputDirectory, baseFileName, shapeType, forceExtract, maxWidth, maxHeight);
            return result?.FilePath;
        }

        /// <summary>
        /// 縦横比を維持してリサイズ後のサイズを計算
        /// </summary>
        /// <param name="originalWidth">元の幅</param>
        /// <param name="originalHeight">元の高さ</param>
        /// <param name="maxWidth">最大幅</param>
        /// <param name="maxHeight">最大高さ</param>
        /// <returns>リサイズ後のサイズ</returns>
        private static Size CalculateResizedDimensions(int originalWidth, int originalHeight, int maxWidth, int maxHeight)
        {
            // 元のサイズが最大サイズ以下の場合はそのまま返す
            if (originalWidth <= maxWidth && originalHeight <= maxHeight)
            {
                return new Size(originalWidth, originalHeight);
            }

            // 縦横比を計算
            double aspectRatio = (double)originalWidth / originalHeight;

            int newWidth, newHeight;

            // 幅が制限を超える場合と高さが制限を超える場合の両方を考慮
            if (originalWidth > maxWidth && originalHeight > maxHeight)
            {
                // 両方が制限を超える場合、より制限が厳しい方に合わせる
                double widthRatio = (double)maxWidth / originalWidth;
                double heightRatio = (double)maxHeight / originalHeight;
                double ratio = Math.Min(widthRatio, heightRatio);

                newWidth = (int)Math.Round(originalWidth * ratio);
                newHeight = (int)Math.Round(originalHeight * ratio);
            }
            else if (originalWidth > maxWidth)
            {
                // 幅のみが制限を超える場合
                newWidth = maxWidth;
                newHeight = (int)Math.Round(maxWidth / aspectRatio);
            }
            else
            {
                // 高さのみが制限を超える場合
                newHeight = maxHeight;
                newWidth = (int)Math.Round(maxHeight * aspectRatio);
            }

            // 最小サイズの保証（1ピクセル以上）
            newWidth = Math.Max(MinPixelSize, newWidth);
            newHeight = Math.Max(MinPixelSize, newHeight);

            return new Size(newWidth, newHeight);
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
