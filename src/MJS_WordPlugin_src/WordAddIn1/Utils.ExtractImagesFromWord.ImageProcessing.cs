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

                        // 境界が無効かチェック
                        if (bounds.Width <= 0 || bounds.Height <= 0)
                        {
                            LogInfo("メタファイルの境界が無効です");
                            return null;
                        }

                        LogInfo($"メタファイル境界: X={bounds.X}, Y={bounds.Y}, Width={bounds.Width}, Height={bounds.Height}");

                        // 実際のコンテンツサイズ（丸め後）
                        int contentWidth = (int)Math.Ceiling(bounds.Width);
                        int contentHeight = (int)Math.Ceiling(bounds.Height);

                        // 最小サイズのフィルタリング（強制抽出の場合はスキップ）
                        if (!forceExtract && (contentWidth < DefaultMinContentSizePixels || contentHeight < DefaultMinContentSizePixels))
                            return null;

                        // ステップ1: 元のサイズでメタファイルを描画してビットマップを作成
                        using (var originalBitmap = new Bitmap(contentWidth, contentHeight, PixelFormat.Format32bppArgb))
                        {
                            using (var graphics = Graphics.FromImage(originalBitmap))
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
                                var destRect = new RectangleF(0, 0, contentWidth, contentHeight);
                                graphics.DrawImage(metafile, destRect, bounds, GraphicsUnit.Pixel);
                            }

                            // ステップ2: 透明ピクセルを除去してトリミング
                            var trimmedBounds = GetTrimmedBounds(originalBitmap);
                            
                            if (trimmedBounds.Width <= 0 || trimmedBounds.Height <= 0)
                            {
                                LogInfo("トリミング後の境界が無効です");
                                return null;
                            }

                            // トリミング後のサイズ
                            int trimmedWidth = trimmedBounds.Width;
                            int trimmedHeight = trimmedBounds.Height;

                            // ステップ3: 最終サイズを決定
                            int finalWidth = trimmedWidth;
                            int finalHeight = trimmedHeight;
                            bool resizedToOriginal = false;
                            bool resizedToMax = false;
                            bool scaledUp = false;

                            // 元の画像サイズが指定されている場合、それをピクセルに変換して使用
                            if (originalWidthPoints > 0 && originalHeightPoints > 0)
                            {
                                int targetWidth = ConvertPointsToPixels(originalWidthPoints);
                                int targetHeight = ConvertPointsToPixels(originalHeightPoints);
                                
                                // スケール倍率を適用
                                if (scaleMultiplier != 1.0f)
                                {
                                    targetWidth = (int)Math.Round(targetWidth * scaleMultiplier);
                                    targetHeight = (int)Math.Round(targetHeight * scaleMultiplier);
                                    scaledUp = true;
                                    LogInfo($"スケール倍率 {scaleMultiplier:F2}x を適用: 目標サイズ {targetWidth}x{targetHeight}px");
                                }
                                
                                // 最大サイズ制限をチェック
                                if (targetWidth > maxWidth || targetHeight > maxHeight)
                                {
                                    // 縦横比を維持してリサイズサイズを計算
                                    var newSize = CalculateResizedDimensions(targetWidth, targetHeight, maxWidth, maxHeight);
                                    finalWidth = newSize.Width;
                                    finalHeight = newSize.Height;
                                    resizedToMax = true;
                                    LogInfo($"目標サイズ({targetWidth}x{targetHeight}px)が最大サイズを超えるため、制限内にリサイズします");
                                }
                                else
                                {
                                    finalWidth = targetWidth;
                                    finalHeight = targetHeight;
                                    resizedToOriginal = true;
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
                                    resizedToMax = true;
                                }
                            }

                            // ステップ4: トリミングされた画像を作成（必要に応じてリサイズ）
                            using (var finalBitmap = new Bitmap(finalWidth, finalHeight, PixelFormat.Format32bppArgb))
                            {
                                using (var graphics = Graphics.FromImage(finalBitmap))
                                {
                                    // 高品質な描画設定（リサイズ時に重要）
                                    graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceOver;
                                    graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                                    graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                                    graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                                    graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                                    graphics.Clear(Color.White);
                                    
                                    // トリミングされた部分を最終サイズに描画
                                    graphics.DrawImage(originalBitmap, 
                                        new Rectangle(0, 0, finalWidth, finalHeight),
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
                                finalBitmap.Save(filePath, ImageFormat.Png);
                                
                                // トリミングとリサイズの正確なログ出力
                                if (scaledUp && resizedToOriginal)
                                {
                                    LogInfo($"画像をトリミング・スケール適用しました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → {scaleMultiplier:F2}x倍 {finalWidth}x{finalHeight}");
                                }
                                else if (resizedToOriginal)
                                {
                                    LogInfo($"画像をトリミング・元サイズにリサイズしました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → 元サイズ {finalWidth}x{finalHeight}");
                                }
                                else if (resizedToMax)
                                {
                                    LogInfo($"画像をトリミング・最大サイズにリサイズしました: {contentWidth}x{contentHeight} → トリミング後 {trimmedWidth}x{trimmedHeight} → 最大サイズ {finalWidth}x{finalHeight}");
                                }
                                else if (trimmedWidth != contentWidth || trimmedHeight != contentHeight)
                                {
                                    LogInfo($"画像をトリミングしました: {contentWidth}x{contentHeight} → {trimmedWidth}x{trimmedHeight}");
                                }
                                
                                LogInfo($"画像を保存しました: {filePath} ({finalWidth}x{finalHeight})");
                                
                                return new ImageExtractionResult
                                {
                                    FilePath = filePath,
                                    PixelWidth = finalWidth,
                                    PixelHeight = finalHeight
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
        /// ビットマップから透明ピクセルを除去した実際のコンテンツ境界を取得
        /// </summary>
        /// <param name="bitmap">ビットマップ</param>
        /// <returns>コンテンツの境界矩形</returns>
        private static Rectangle GetTrimmedBounds(Bitmap bitmap)
        {
            int minX = bitmap.Width;
            int minY = bitmap.Height;
            int maxX = 0;
            int maxY = 0;

            // すべてのピクセルをスキャンして、透明でないピクセル範囲を取得
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

            // 幅も高さも超過する場合と高さのみ超過する場合の両方を考慮
            if (originalWidth > maxWidth && originalHeight > maxHeight)
            {
                // 両方とも超過する場合、より制限的な方に合わせる
                double widthRatio = (double)maxWidth / originalWidth;
                double heightRatio = (double)maxHeight / originalHeight;
                double ratio = Math.Min(widthRatio, heightRatio);

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
