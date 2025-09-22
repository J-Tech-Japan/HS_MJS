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
        /// <param name="includeInlineShapes">インライン図形を含むかどうか</param>
        /// <param name="includeShapes">フローティング図形を含むかどうか</param>
        /// <param name="includeCanvasItems">キャンバス内アイテムを含むかどうか</param>
        /// <param name="includeFreeforms">フリーフォーム図形を含むかどうか</param>
        /// <param name="addMarkers">抽出した画像の後ろにマーカーテキストを追加するかどうか</param>
        /// <param name="skipCoverMarkers">表紙（第1セクション）の画像にマーカーを追加しないかどうか</param>
        /// <param name="minOriginalWidth">元画像の最小幅（ポイント単位）</param>
        /// <param name="minOriginalHeight">元画像の最小高さ（ポイント単位）</param>
        /// <param name="includeMjsTableImages">MJS_画像（表内）スタイルの画像を抽出するかどうか</param>
        /// <param name="maxOutputWidth">出力画像の最大幅（ピクセル単位、デフォルト: 2048）</param>
        /// <param name="maxOutputHeight">出力画像の最大高さ（ピクセル単位、デフォルト: 2048）</param>
        /// <returns>抽出された画像情報のリスト</returns>
        public static List<ExtractedImageInfo> ExtractImagesFromWord(
            Word.Document document, 
            string outputDirectory,
            bool includeInlineShapes = true,
            bool includeShapes = true,
            bool includeCanvasItems = true,
            bool includeFreeforms = true,
            bool addMarkers = true,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            if (document == null)
                throw new ArgumentNullException(nameof(document));

            if (string.IsNullOrEmpty(outputDirectory))
                throw new ArgumentException("出力ディレクトリが指定されていません。", nameof(outputDirectory));

            // 出力ディレクトリが存在しない場合は作成
            if (!Directory.Exists(outputDirectory))
                Directory.CreateDirectory(outputDirectory);

            var extractedImages = new List<ExtractedImageInfo>();
            int imageCounter = 1;

            try
            {
                // インライン図形の抽出
                if (includeInlineShapes)
                {
                    ExtractInlineShapes(
                        document, 
                        outputDirectory, 
                        ref imageCounter, 
                        extractedImages, 
                        addMarkers,
                        skipCoverMarkers,
                        minOriginalWidth,
                        minOriginalHeight,
                        includeMjsTableImages,
                        maxOutputWidth,
                        maxOutputHeight);
                }

                // フローティング図形の抽出
                if (includeShapes)
                {
                    ExtractFloatingShapes(
                        document, 
                        outputDirectory, 
                        ref imageCounter, 
                        extractedImages, 
                        includeCanvasItems, 
                        includeFreeforms, 
                        addMarkers,
                        skipCoverMarkers,
                        minOriginalWidth,
                        minOriginalHeight,
                        includeMjsTableImages,
                        maxOutputWidth,
                        maxOutputHeight);
                }

                return extractedImages;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"画像抽出中にエラーが発生しました: {ex.Message}", ex);
            }
        }
        
        /// <summary>
        /// インライン図形からEnhMetaFileBitsを使用して画像を抽出
        /// </summary>
        private static void ExtractInlineShapes(
            Word.Document document, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages, 
            bool addMarkers = false,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            foreach (Word.InlineShape inlineShape in document.InlineShapes)
            {
                try
                {
                    // 段落のスタイルを取得
                    string paragraphStyle = GetInlineShapeParagraphStyle(inlineShape);
                    
                    // MJSスタイルによる条件チェック
                    CheckMjsStyleConditions(paragraphStyle, out bool forceExtract, out bool forceSkip, includeMjsTableImages);
                    
                    // 強制スキップ対象の場合
                    if (forceSkip)
                    {
                        System.Diagnostics.Debug.WriteLine($"インライン図形をスキップ: スタイル '{paragraphStyle}' により強制スキップ");
                        continue;
                    }

                    // 元画像サイズでのフィルタリング（強制抽出の場合はスキップ）
                    float originalWidth = inlineShape.Width;
                    float originalHeight = inlineShape.Height;
                    
                    if (!forceExtract && (originalWidth < minOriginalWidth || originalHeight < minOriginalHeight))
                    {
                        System.Diagnostics.Debug.WriteLine($"インライン図形をスキップ: 元サイズが小さすぎます ({originalWidth:F1}x{originalHeight:F1} points)");
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
                            maxOutputWidth,
                            maxOutputHeight);
                        
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
                            if (addMarkers && !IsInCoverSection(inlineShape.Range, skipCoverMarkers))
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
                    System.Diagnostics.Debug.WriteLine($"インライン図形の抽出でエラー: {ex.Message}");
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
            bool includeCanvasItems, 
            bool includeFreeforms, 
            bool addMarkers = false,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            foreach (Word.Shape shape in document.Shapes)
            {
                try
                {
                    // キャンバス図形の場合
                    if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                    {
                        ExtractCanvasShape(
                            shape, 
                            outputDirectory, 
                            ref imageCounter, 
                            extractedImages, 
                            includeCanvasItems, 
                            addMarkers,
                            skipCoverMarkers,
                            minOriginalWidth,
                            minOriginalHeight,
                            includeMjsTableImages,
                            maxOutputWidth,
                            maxOutputHeight);
                    }
                    // フリーフォーム図形の場合
                    else if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoFreeform)
                    {
                        if (includeFreeforms)
                        {
                            ExtractSingleShape(
                                shape, 
                                outputDirectory, 
                                ref imageCounter, 
                                extractedImages, 
                                "freeform", 
                                addMarkers,
                                skipCoverMarkers,
                                minOriginalWidth,
                                minOriginalHeight,
                                includeMjsTableImages,
                                maxOutputWidth,
                                maxOutputHeight);
                        }
                    }
                    // 通常の図形の場合
                    else
                    {
                        ExtractSingleShape(
                            shape, 
                            outputDirectory, 
                            ref imageCounter, 
                            extractedImages, 
                            "shape", 
                            addMarkers,
                            skipCoverMarkers,
                            minOriginalWidth,
                            minOriginalHeight,
                            includeMjsTableImages,
                            maxOutputWidth,
                            maxOutputHeight);
                    }
                }
                catch (Exception ex)
                {
                    // 個別の図形でエラーが発生しても処理を継続
                    System.Diagnostics.Debug.WriteLine($"フローティング図形の抽出でエラー: {ex.Message}");
                }
            }
        }

        /// <summary>
        /// キャンバス図形から画像を抽出
        /// </summary>
        private static void ExtractCanvasShape(
            Word.Shape canvas, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages, 
            bool includeCanvasItems, 
            bool addMarkers = false,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            try
            {
                // アンカー段落のスタイルを取得
                string anchorParagraphStyle = GetShapeAnchorParagraphStyle(canvas);
                
                // MJSスタイルによる条件チェック
                CheckMjsStyleConditions(anchorParagraphStyle, out bool forceExtract, out bool forceSkip, includeMjsTableImages);
                
                // 強制スキップ対象の場合
                if (forceSkip)
                {
                    System.Diagnostics.Debug.WriteLine($"キャンバス図形をスキップ: スタイル '{anchorParagraphStyle}' により強制スキップ");
                    return;
                }

                // 元画像サイズでのフィルタリング（強制抽出の場合はスキップ）
                float originalWidth = canvas.Width;
                float originalHeight = canvas.Height;
                
                if (!forceExtract && (originalWidth < minOriginalWidth || originalHeight < minOriginalHeight))
                {
                    System.Diagnostics.Debug.WriteLine($"キャンバス図形をスキップ: 元サイズが小さすぎます ({originalWidth:F1}x{originalHeight:F1} points)");
                }
                else
                {
                    // キャンバス全体を画像として抽出
                    canvas.Select();
                    byte[] canvasData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                    
                    if (canvasData != null && canvasData.Length > 0)
                    {
                        var extractResult = ExtractImageFromMetaFileDataWithSize(
                            canvasData, 
                            outputDirectory, 
                            $"canvas_{imageCounter}", 
                            "Canvas",
                            forceExtract,
                            maxOutputWidth,
                            maxOutputHeight);
                        
                        if (extractResult != null)
                        {
                            var imageInfo = new ExtractedImageInfo(
                                extractResult.FilePath, 
                                "キャンバス", 
                                canvas.Anchor?.Start ?? 0,
                                originalWidth,
                                originalHeight,
                                extractResult.PixelWidth,
                                extractResult.PixelHeight
                            );
                            extractedImages.Add(imageInfo);

                            // マーカーを追加（表紙の画像は除外）
                            if (addMarkers && !IsShapeInCoverSection(canvas, skipCoverMarkers))
                            {
                                if (canvas.Anchor != null)
                                {
                                    InsertMarkerAtPosition(canvas.Anchor, extractResult.FilePath);
                                }
                                else
                                {
                                    // Anchorが利用できない場合、キャンバスが選択された状態で
                                    // 現在の選択範囲を使用してマーカーを挿入
                                    InsertMarkerForSelectedCanvas(extractResult.FilePath);
                                }
                            }

                            imageCounter++;
                        }
                    }
                }

                // キャンバス内のアイテムを個別に抽出
                if (includeCanvasItems && canvas.CanvasItems.Count > 0)
                {
                    ExtractCanvasItems(
                        canvas, 
                        outputDirectory, 
                        ref imageCounter, 
                        extractedImages, 
                        addMarkers,
                        skipCoverMarkers,
                        minOriginalWidth,
                        minOriginalHeight,
                        includeMjsTableImages,
                        maxOutputWidth,
                        maxOutputHeight);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"キャンバス図形の抽出でエラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 選択されたキャンバス用のマーカーを挿入
        /// </summary>
        private static void InsertMarkerForSelectedCanvas(string filePath)
        {
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);
                
                // 現在の選択範囲を取得
                var selection = Globals.ThisAddIn.Application.Selection;
                
                // 選択範囲の末尾に移動
                selection.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                
                // 改行を挿入して新しい行を作成
                selection.TypeText("\r");
                
                // 新しい行に特殊な識別子を挿入
                string marker = $"[IMAGEMARKER:{markerText}]";
                selection.TypeText(marker);
                
                // マーカーの後に改行を追加
                selection.TypeText("\r");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"キャンバス用マーカー挿入エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// キャンバス内のアイテムを抽出
        /// </summary>
        private static void ExtractCanvasItems(
            Word.Shape canvas, 
            string outputDirectory, 
            ref int imageCounter, 
            List<ExtractedImageInfo> extractedImages, 
            bool addMarkers = false,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            foreach (Word.Shape canvasItem in canvas.CanvasItems)
            {
                try
                {
                    ExtractSingleShape(
                        canvasItem, 
                        outputDirectory, 
                        ref imageCounter, 
                        extractedImages, 
                        "canvas_item", 
                        addMarkers,
                        skipCoverMarkers,
                        minOriginalWidth,
                        minOriginalHeight,
                        includeMjsTableImages,
                        maxOutputWidth,
                        maxOutputHeight);
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"キャンバスアイテムの抽出でエラー: {ex.Message}");
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
            string prefix = "shape", 
            bool addMarkers = false,
            bool skipCoverMarkers = true,
            float minOriginalWidth = 50.0f,
            float minOriginalHeight = 50.0f,
            bool includeMjsTableImages = true,
            int maxOutputWidth = 2048,
            int maxOutputHeight = 2048)
        {
            try
            {
                // アンカー段落のスタイルを取得
                string anchorParagraphStyle = GetShapeAnchorParagraphStyle(shape);
                
                // MJSスタイルによる条件チェック
                CheckMjsStyleConditions(anchorParagraphStyle, out bool forceExtract, out bool forceSkip, includeMjsTableImages);
                
                // 強制スキップ対象の場合
                if (forceSkip)
                {
                    System.Diagnostics.Debug.WriteLine($"{prefix}図形をスキップ: スタイル '{anchorParagraphStyle}' により強制スキップ");
                    return;
                }

                // 元画像サイズでのフィルタリング（強制抽出の場合はスキップ）
                float originalWidth = shape.Width;
                float originalHeight = shape.Height;
                
                if (!forceExtract && (originalWidth < minOriginalWidth || originalHeight < minOriginalHeight))
                {
                    System.Diagnostics.Debug.WriteLine($"{prefix}図形をスキップ: 元サイズが小さすぎます ({originalWidth:F1}x{originalHeight:F1} points)");
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
                        maxOutputWidth,
                        maxOutputHeight);
                    
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
                        if (addMarkers && shape.Anchor != null && !IsShapeInCoverSection(shape, skipCoverMarkers))
                        {
                            InsertMarkerAtPosition(shape.Anchor, extractResult.FilePath);
                        }

                        imageCounter++;
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"図形の抽出でエラー: {ex.Message}");
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
        /// <param name="maxWidth">最大幅（ピクセル、デフォルト: 2048）</param>
        /// <param name="maxHeight">最大高さ（ピクセル、デフォルト: 2048）</param>
        /// <returns>作成されたファイルのパスとピクセルサイズ、失敗時はnull</returns>
        private static ImageExtractionResult ExtractImageFromMetaFileDataWithSize(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType,
            bool forceExtract = false,
            int maxWidth = 2048,
            int maxHeight = 2048)
        {
            try
            {
                using (var memoryStream = new MemoryStream(metaFileData))
                {
                    using (var originalImage = Image.FromStream(memoryStream))
                    {
                        // 最小サイズのフィルタリング（強制抽出の場合はスキップ）
                        if (!forceExtract && (originalImage.Width < 250 || originalImage.Height < 250))
                            return null;

                        // リサイズが必要かチェック
                        bool needsResize = originalImage.Width > maxWidth || originalImage.Height > maxHeight;
                        Image finalImage = originalImage;
                        int finalWidth = originalImage.Width;
                        int finalHeight = originalImage.Height;

                        if (needsResize)
                        {
                            // 縦横比を維持してリサイズサイズを計算
                            var newSize = CalculateResizedDimensions(originalImage.Width, originalImage.Height, maxWidth, maxHeight);
                            
                            // リサイズされた画像を作成
                            finalImage = ResizeImageWithQuality(originalImage, newSize.Width, newSize.Height);
                            finalWidth = newSize.Width;
                            finalHeight = newSize.Height;
                            
                            System.Diagnostics.Debug.WriteLine($"画像をリサイズしました: {originalImage.Width}x{originalImage.Height} → {finalWidth}x{finalHeight}");
                        }

                        try
                        {
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
                            finalImage.Save(filePath, ImageFormat.Png);
                            
                            return new ImageExtractionResult
                            {
                                FilePath = filePath,
                                PixelWidth = finalWidth,
                                PixelHeight = finalHeight
                            };
                        }
                        finally
                        {
                            // リサイズした画像のリソースを解放（オリジナル画像と異なる場合のみ）
                            if (needsResize && finalImage != originalImage)
                            {
                                finalImage.Dispose();
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"メタファイルデータからの画像生成でエラー: {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// EnhMetaFileBitsから画像ファイルを作成（旧版、互換性のため保持）
        /// </summary>
        /// <param name="metaFileData">メタファイルデータ</param>
        /// <param name="outputDirectory">出力ディレクトリ</param>
        /// <param name="baseFileName">ベースファイル名</param>
        /// <param name="shapeType">図形タイプ</param>
        /// <param name="forceExtract">強制抽出フラグ</param>
        /// <param name="maxWidth">最大幅（ピクセル、デフォルト: 2048）</param>
        /// <param name="maxHeight">最大高さ（ピクセル、デフォルト: 2048）</param>
        /// <returns>作成されたファイルのパス、失敗時はnull</returns>
        private static string ExtractImageFromMetaFileData(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType,
            bool forceExtract = false,
            int maxWidth = 2048,
            int maxHeight = 2048)
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
            newWidth = Math.Max(1, newWidth);
            newHeight = Math.Max(1, newHeight);

            return new Size(newWidth, newHeight);
        }

        /// <summary>
        /// 高品質でリサイズされた画像を作成
        /// </summary>
        /// <param name="originalImage">元の画像</param>
        /// <param name="newWidth">新しい幅</param>
        /// <param name="newHeight">新しい高さ</param>
        /// <returns>リサイズされた画像</returns>
        private static Image ResizeImageWithQuality(Image originalImage, int newWidth, int newHeight)
        {
            var resizedImage = new Bitmap(newWidth, newHeight);
            
            using (var graphics = Graphics.FromImage(resizedImage))
            {
                // 高品質なリサイズのための設定
                graphics.CompositingMode = System.Drawing.Drawing2D.CompositingMode.SourceCopy;
                graphics.CompositingQuality = System.Drawing.Drawing2D.CompositingQuality.HighQuality;
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.HighQuality;
                graphics.PixelOffsetMode = System.Drawing.Drawing2D.PixelOffsetMode.HighQuality;

                // リサイズ実行
                graphics.DrawImage(originalImage, 0, 0, newWidth, newHeight);
            }

            return resizedImage;
        }
    }
}
