// Utils.ExtractImagesFromWord.ImageProcessing.cs

using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;

namespace WordAddIn1
{
    internal partial class Utils
    {
        // ピクセルデータ処理の最適化用定数
        private const int TemporaryBitmapSize = 1; // メタファイル境界取得用の一時ビットマップサイズ
        private const int BytesPerPixel = 4; // ARGB形式のバイト数
        private const int AlphaChannelOffset = 3; // ARGBフォーマットのアルファチャンネルオフセット
        
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
        /// <param name="disableResize">リサイズを無効化するかどうか（デフォルト: false）</param>
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
            float scaleMultiplier = DefaultOutputScaleMultiplier,
            bool disableResize = false)
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
                            disableResize,
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
            using (var tempBitmap = new Bitmap(TemporaryBitmapSize, TemporaryBitmapSize))
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
        /// <param name="useTransparentBackground">透明背景を使用するか（デフォルト: false）</param>
        private static void RenderFinalBitmap(Bitmap finalBitmap, Bitmap sourceBitmap, 
            Rectangle trimmedBounds, Size finalSize, bool useTransparentBackground = false)
        {
            using (var graphics = Graphics.FromImage(finalBitmap))
            {
                ConfigureHighQualityGraphics(graphics);
                
                // 背景色の選択
                graphics.Clear(useTransparentBackground ? Color.Transparent : Color.White);
                
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
            public bool DisabledResize;
        }

        /// <summary>
        /// 最終的な画像サイズを計算
        /// </summary>
        private static Size CalculateFinalSize(int trimmedWidth, int trimmedHeight,
            float originalWidthPoints, float originalHeightPoints, float scaleMultiplier,
            int maxWidth, int maxHeight, bool disableResize, out ResizeInfo resizeInfo)
        {
            resizeInfo = new ResizeInfo();
            var finalWidth = trimmedWidth;
            var finalHeight = trimmedHeight;

            // リサイズ無効化フラグが設定されている場合は、トリミング後のサイズをそのまま返す
            if (disableResize)
            {
                resizeInfo.DisabledResize = true;
                LogInfo($"リサイズ無効化: トリミング後のサイズをそのまま出力します ({trimmedWidth}x{trimmedHeight}px)");
                return new Size(trimmedWidth, trimmedHeight);
            }

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
            if (resizeInfo.DisabledResize)
            {
                LogInfo($"リサイズ無効化: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} (リサイズなし)");
            }
            else if (resizeInfo.ScaledUp && resizeInfo.ResizedToOriginal)
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
        /// ビットマップから透明ピクセルを除去した実際のコンテンツ境界を取得（最適化版）
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
                var width = bitmap.Width;
                var height = bitmap.Height;
                var stride = bitmapData.Stride;
                var bytes = Math.Abs(stride) * height;
                var rgbValues = new byte[bytes];

                // ピクセルデータを配列にコピー
                System.Runtime.InteropServices.Marshal.Copy(bitmapData.Scan0, rgbValues, 0, bytes);

                // エッジ検出の最適化：上下左右から順にスキャンして早期終了
                
                // 上端を検出
                var top = FindTopEdge(rgbValues, width, height, stride);
                if (top == -1) return Rectangle.Empty;

                // 下端を検出
                var bottom = FindBottomEdge(rgbValues, width, height, stride);
                if (bottom == -1 || bottom < top) return Rectangle.Empty;

                // 左端を検出（上端〜下端の範囲内で）
                var left = FindLeftEdge(rgbValues, width, top, bottom, stride);
                if (left == -1) return Rectangle.Empty;

                // 右端を検出（上端〜下端の範囲内で）
                var right = FindRightEdge(rgbValues, width, top, bottom, stride);
                if (right == -1 || right < left) return Rectangle.Empty;

                // 境界矩形を返す（幅と高さは+1して含める）
                return new Rectangle(left, top, right - left + 1, bottom - top + 1);
            }
            finally
            {
                bitmap.UnlockBits(bitmapData);
            }
        }

        /// <summary>
        /// 上端のエッジを検出（最適化版）
        /// </summary>
        private static int FindTopEdge(byte[] rgbValues, int width, int height, int stride)
        {
            for (var y = 0; y < height; y++)
            {
                var rowStart = y * stride;
                for (var x = 0; x < width; x++)
                {
                    var pixelIndex = rowStart + (x * BytesPerPixel);
                    if (rgbValues[pixelIndex + AlphaChannelOffset] > AlphaThresholdForTransparency)
                    {
                        return y; // 最初の不透明ピクセルが見つかった行
                    }
                }
            }
            return -1; // 不透明ピクセルが見つからない
        }

        /// <summary>
        /// 下端のエッジを検出（最適化版）
        /// </summary>
        private static int FindBottomEdge(byte[] rgbValues, int width, int height, int stride)
        {
            for (var y = height - 1; y >= 0; y--)
            {
                var rowStart = y * stride;
                for (var x = 0; x < width; x++)
                {
                    var pixelIndex = rowStart + (x * BytesPerPixel);
                    if (rgbValues[pixelIndex + AlphaChannelOffset] > AlphaThresholdForTransparency)
                    {
                        return y; // 最初の不透明ピクセルが見つかった行
                    }
                }
            }
            return -1; // 不透明ピクセルが見つからない
        }

        /// <summary>
        /// 左端のエッジを検出（最適化版）
        /// </summary>
        private static int FindLeftEdge(byte[] rgbValues, int width, int top, int bottom, int stride)
        {
            for (var x = 0; x < width; x++)
            {
                for (var y = top; y <= bottom; y++)
                {
                    var pixelIndex = y * stride + (x * BytesPerPixel);
                    if (rgbValues[pixelIndex + AlphaChannelOffset] > AlphaThresholdForTransparency)
                    {
                        return x; // 最初の不透明ピクセルが見つかった列
                    }
                }
            }
            return -1; // 不透明ピクセルが見つからない
        }

        /// <summary>
        /// 右端のエッジを検出（最適化版）
        /// </summary>
        private static int FindRightEdge(byte[] rgbValues, int width, int top, int bottom, int stride)
        {
            for (var x = width - 1; x >= 0; x--)
            {
                for (var y = top; y <= bottom; y++)
                {
                    var pixelIndex = y * stride + (x * BytesPerPixel);
                    if (rgbValues[pixelIndex + AlphaChannelOffset] > AlphaThresholdForTransparency)
                    {
                        return x; // 最初の不透明ピクセルが見つかった列
                    }
                }
            }
            return -1; // 不透明ピクセルが見つからない
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
