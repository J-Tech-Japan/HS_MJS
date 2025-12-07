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
    }
}
