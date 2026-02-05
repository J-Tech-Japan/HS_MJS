// Utils.ExtractImagesFromWord.ImageProcessing.cs

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace WordAddIn1
{
    internal partial class Utils
    {
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
        /// EnhMetaFileBitsから画像ファイルを作成し、PNG画像のサイズを取得
        /// </summary>
        /// <param name="metaFileData">メタファイルデータ</param>
        /// <param name="outputDirectory">出力ディレクトリ</param>
        /// <param name="baseFileName">ベースファイル名</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="forceExtract">強制抽出フラグ</param>
        /// <param name="maxWidth">最大幅（ピクセル、デフォルト: 1024）</param>
        /// <param name="maxHeight">最大高さ（ピクセル、デフォルト: 1024）</param>
        /// <param name="originalWidthPoints">元の画像の幅（ポイント単位、0の場合は使用しない）</param>
        /// <param name="originalHeightPoints">元の画像の高さ（ポイント単位、0の場合は使用しない）</param>
        /// <param name="scaleMultiplier">出力スケール倍率（デフォルト: 1.0）</param>
        /// <returns>作成されたファイルのパスとピクセルサイズ、失敗時null</returns>
        private static ImageExtractionResult ExtractImageFromMetaFileDataWithSize(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType,
            bool forceExtract = false,
            int maxWidth = DefaultMaxOutputSizePixels,
            int maxHeight = DefaultMaxOutputSizePixels,
            float originalWidthPoints = 0,
            float originalHeightPoints = 0,
            float scaleMultiplier = DefaultOutputScaleMultiplier)
        {
            try
            {
                using (var memoryStream = new MemoryStream(metaFileData))
                using (var metafile = new System.Drawing.Imaging.Metafile(memoryStream))
                {
                    // メタファイルの実際のコンテンツ境界を取得
                    var bounds = GetMetafileBounds(metafile);
                    if (bounds.Width <= 0 || bounds.Height <= 0)
                    {
                        LogInfo("メタファイルの境界が無効です");
                        return null;
                    }

                    LogInfo($"メタファイル境界: X={bounds.X}, Y={bounds.Y}, Width={bounds.Width}, Height={bounds.Height}");

                    // 実際のコンテンツサイズ（丸め後）
                    var contentWidth = (int)Math.Ceiling(bounds.Width);
                    var contentHeight = (int)Math.Ceiling(bounds.Height);

                    // 最小サイズのフィルタリング（強制抽出の場合はスキップ）
                    if (!forceExtract && (contentWidth < DefaultMinContentSizePixels || contentHeight < DefaultMinContentSizePixels))
                    {
                        return null;
                    }

                    // ステップ1: 元のサイズでメタファイルを描画してビットマップを作成
                    using (var originalBitmap = new Bitmap(contentWidth, contentHeight, PixelFormat.Format32bppArgb))
                    {
                        RenderMetafileToBitmap(originalBitmap, metafile, bounds, contentWidth, contentHeight);

                        // ステップ2: 透明ピクセルを除去してトリミング
                        var trimmedBounds = GetTrimmedBounds(originalBitmap);
                        if (trimmedBounds.Width <= 0 || trimmedBounds.Height <= 0)
                        {
                            LogInfo("トリミング後の境界が無効です");
                            return null;
                        }

                        // ステップ3: 最終サイズを決定
                        var finalSize = CalculateFinalSize(
                            trimmedBounds.Width, 
                            trimmedBounds.Height,
                            originalWidthPoints,
                            originalHeightPoints,
                            scaleMultiplier,
                            maxWidth,
                            maxHeight,
                            out var resizeInfo);

                        // ステップ4: トリミングされた画像を作成（必要に応じてリサイズ）
                        using (var finalBitmap = new Bitmap(finalSize.Width, finalSize.Height, PixelFormat.Format32bppArgb))
                        {
                            RenderFinalBitmap(finalBitmap, originalBitmap, trimmedBounds, finalSize);

                            // ファイル名の生成と保存
                            var filePath = GenerateUniqueFilePath(outputDirectory, baseFileName, shapeType);
                            finalBitmap.Save(filePath, ImageFormat.Png);
                            
                            // ログ出力
                            LogResizeOperation(contentWidth, contentHeight, trimmedBounds.Width, trimmedBounds.Height, 
                                             finalSize.Width, finalSize.Height, resizeInfo);
                            LogInfo($"画像を保存しました: {filePath} ({finalSize.Width}x{finalSize.Height})");
                            
                            return new ImageExtractionResult
                            {
                                FilePath = filePath,
                                PixelWidth = finalSize.Width,
                                PixelHeight = finalSize.Height
                            };
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
        /// メタファイルの境界を取得
        /// </summary>
        private static RectangleF GetMetafileBounds(System.Drawing.Imaging.Metafile metafile)
        {
            const int temporaryBitmapSize = 1;
            using (var tempBitmap = new Bitmap(temporaryBitmapSize, temporaryBitmapSize))
            using (var graphics = Graphics.FromImage(tempBitmap))
            {
                var unit = GraphicsUnit.Pixel;
                return metafile.GetBounds(ref unit);
            }
        }

        /// <summary>
        /// メタファイルをビットマップに描画
        /// </summary>
        private static void RenderMetafileToBitmap(Bitmap bitmap, System.Drawing.Imaging.Metafile metafile, 
            RectangleF bounds, int width, int height)
        {
            using (var graphics = Graphics.FromImage(bitmap))
            {
                ConfigureHighQualityGraphics(graphics);
                graphics.Clear(Color.Transparent);
                
                var destRect = new RectangleF(0, 0, width, height);
                graphics.DrawImage(metafile, destRect, bounds, GraphicsUnit.Pixel);
            }
        }

        /// <summary>
        /// 最終的なビットマップを描画
        /// </summary>
        private static void RenderFinalBitmap(Bitmap finalBitmap, Bitmap sourceBitmap, 
            Rectangle trimmedBounds, Size finalSize)
        {
            using (var graphics = Graphics.FromImage(finalBitmap))
            {
                ConfigureHighQualityGraphics(graphics);
                graphics.Clear(Color.White);
                
                graphics.DrawImage(sourceBitmap, 
                    new Rectangle(0, 0, finalSize.Width, finalSize.Height),
                    trimmedBounds,
                    GraphicsUnit.Pixel);
            }
        }

        /// <summary>
        /// 高品質な描画設定を適用
        /// </summary>
        private static void ConfigureHighQualityGraphics(Graphics graphics)
        {
            graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
            graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
            graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
            graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
            graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;
        }

        /// <summary>
        /// リサイズ情報を保持する構造体
        /// </summary>
        private struct ResizeInfo
        {
            public bool ScaledUp;
            public bool ResizedToOriginal;
            public bool ResizedToMax;
        }

        /// <summary>
        /// 最終的な画像サイズを計算
        /// </summary>
        private static Size CalculateFinalSize(int trimmedWidth, int trimmedHeight,
            float originalWidthPoints, float originalHeightPoints, float scaleMultiplier,
            int maxWidth, int maxHeight, out ResizeInfo resizeInfo)
        {
            resizeInfo = new ResizeInfo();
            var finalWidth = trimmedWidth;
            var finalHeight = trimmedHeight;

            // 元の画像サイズが指定されている場合、それをピクセルに変換して使用
            if (originalWidthPoints > 0 && originalHeightPoints > 0)
            {
                var targetWidth = ConvertPointsToPixels(originalWidthPoints);
                var targetHeight = ConvertPointsToPixels(originalHeightPoints);
                
                // スケール倍率を適用
                if (scaleMultiplier != 1.0f)
                {
                    targetWidth = (int)Math.Round(targetWidth * scaleMultiplier);
                    targetHeight = (int)Math.Round(targetHeight * scaleMultiplier);
                    resizeInfo.ScaledUp = true;
                    LogInfo($"スケール倍率 {scaleMultiplier:F2}x を適用: 目標サイズ {targetWidth}x{targetHeight}px");
                }
                
                // 最大サイズ制限をチェック
                if (targetWidth > maxWidth || targetHeight > maxHeight)
                {
                    var newSize = CalculateResizedDimensions(targetWidth, targetHeight, maxWidth, maxHeight);
                    finalWidth = newSize.Width;
                    finalHeight = newSize.Height;
                    resizeInfo.ResizedToMax = true;
                    LogInfo($"目標サイズ({targetWidth}x{targetHeight}px)が最大サイズを超えるため、制限内にリサイズします");
                }
                else
                {
                    finalWidth = targetWidth;
                    finalHeight = targetHeight;
                    resizeInfo.ResizedToOriginal = true;
                }
            }
            else
            {
                // 元のサイズが指定されていない場合は、最大サイズ制限のみ適用
                if (trimmedWidth > maxWidth || trimmedHeight > maxHeight)
                {
                    var newSize = CalculateResizedDimensions(trimmedWidth, trimmedHeight, maxWidth, maxHeight);
                    finalWidth = newSize.Width;
                    finalHeight = newSize.Height;
                    resizeInfo.ResizedToMax = true;
                }
            }

            return new Size(finalWidth, finalHeight);
        }

        /// <summary>
        /// 一意のファイルパスを生成
        /// </summary>
        private static string GenerateUniqueFilePath(string outputDirectory, string baseFileName, string shapeType)
        {
            var fileName = $"{baseFileName}_{shapeType}.png";
            var filePath = Path.Combine(outputDirectory, fileName);

            var duplicateCounter = 1;
            while (File.Exists(filePath))
            {
                fileName = $"{baseFileName}_{shapeType}_{duplicateCounter}.png";
                filePath = Path.Combine(outputDirectory, fileName);
                duplicateCounter++;
            }

            return filePath;
        }

        /// <summary>
        /// リサイズ操作のログを出力
        /// </summary>
        private static void LogResizeOperation(int contentWidth, int contentHeight, 
            int trimmedWidth, int trimmedHeight, int finalWidth, int finalHeight, ResizeInfo resizeInfo)
        {
            if (resizeInfo.ScaledUp && resizeInfo.ResizedToOriginal)
            {
                LogInfo($"画像をトリミング・スケール適用しました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → スケール適用後 {finalWidth}x{finalHeight}");
            }
            else if (resizeInfo.ResizedToOriginal)
            {
                LogInfo($"画像をトリミング・元サイズにリサイズしました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → 元サイズ {finalWidth}x{finalHeight}");
            }
            else if (resizeInfo.ResizedToMax)
            {
                LogInfo($"画像をトリミング・最大サイズにリサイズしました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → 最大サイズ {finalWidth}x{finalHeight}");
            }
            else if (trimmedWidth != contentWidth || trimmedHeight != contentHeight)
            {
                LogInfo($"画像をトリミングしました: {contentWidth}x{contentHeight} → {trimmedWidth}x{trimmedHeight}");
            }
        }

        /// <summary>
        /// ビットマップから透明ピクセルを除去した実際のコンテンツ境界を取得（高速版）
        /// </summary>
        /// <param name="bitmap">ビットマップ</param>
        /// <returns>コンテンツの境界矩形</returns>
        private static Rectangle GetTrimmedBounds(Bitmap bitmap)
        {
            var bitmapData = bitmap.LockBits(
                new Rectangle(0, 0, bitmap.Width, bitmap.Height),
                ImageLockMode.ReadOnly,
                PixelFormat.Format32bppArgb);

            try
            {
                var minX = bitmap.Width;
                var minY = bitmap.Height;
                var maxX = 0;
                var maxY = 0;

                unsafe
                {
                    var stride = bitmapData.Stride;
                    var scan0 = (byte*)bitmapData.Scan0;

                    // すべてのピクセルをスキャンして、透明でないピクセル範囲を取得
                    for (var y = 0; y < bitmap.Height; y++)
                    {
                        var row = scan0 + (y * stride);
                        for (var x = 0; x < bitmap.Width; x++)
                        {
                            var alpha = row[x * 4 + 3]; // ARGB形式のアルファチャンネル
                            
                            // 完全に透明でないピクセルを検出
                            if (alpha > AlphaThresholdForTransparency)
                            {
                                if (x < minX) minX = x;
                                if (x > maxX) maxX = x;
                                if (y < minY) minY = y;
                                if (y > maxY) maxY = y;
                            }
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
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
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
            var aspectRatio = (double)originalWidth / originalHeight;

            int newWidth, newHeight;

            // 幅も高さも超過する場合と高さのみ超過する場合の両方を考慮
            if (originalWidth > maxWidth && originalHeight > maxHeight)
            {
                // 両方とも超過する場合、より制限的な方に合わせる
                var widthRatio = (double)maxWidth / originalWidth;
                var heightRatio = (double)maxHeight / originalHeight;
                var ratio = Math.Min(widthRatio, heightRatio);

                newWidth = (int)Math.Round(originalWidth * ratio);
                newHeight = (int)Math.Round(originalHeight * ratio);
            }
            else if (originalWidth > maxWidth)
            {
                // 幅のみが超過する場合
                newWidth = maxWidth;
                newHeight = (int)Math.Round(maxWidth / aspectRatio);
            }
            else
            {
                // 高さのみが超過する場合
                newHeight = maxHeight;
                newWidth = (int)Math.Round(maxHeight * aspectRatio);
            }

            // 最小サイズの保証（1ピクセル以上）
            newWidth = Math.Max(MinPixelSize, newWidth);
            newHeight = Math.Max(MinPixelSize, newHeight);

            return new Size(newWidth, newHeight);
        }
    }
}
