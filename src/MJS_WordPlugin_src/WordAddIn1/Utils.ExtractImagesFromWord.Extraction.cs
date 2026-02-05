// Utils.ExtractImagesFromWord.Extraction.cs

using System;
using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    internal partial class Utils
    {
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
                            options.MaxOutputHeight,
                            originalWidth,
                            originalHeight,
                            options.OutputScaleMultiplier,
                            options.DisableResize);
                        
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
                        options.MaxOutputHeight,
                        originalWidth,
                        originalHeight,
                        options.OutputScaleMultiplier,
                        options.DisableResize);
                    
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
    }
}
