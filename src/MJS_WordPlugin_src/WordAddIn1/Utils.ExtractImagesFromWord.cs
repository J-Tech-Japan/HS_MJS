// Utils.ExtractImagesFromWord.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
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

            public ExtractedImageInfo(
                string filePath, 
                string imageType, 
                int position)
            {
                FilePath = filePath;
                ImageType = imageType;
                Position = position;
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
        /// <returns>抽出された画像情報のリスト</returns>
        public static List<ExtractedImageInfo> ExtractImagesFromWord(
            Word.Document document, 
            string outputDirectory,
            bool includeInlineShapes = true,
            bool includeShapes = true,
            bool includeCanvasItems = true,
            bool includeFreeforms = true,
            bool addMarkers = true)
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
                        addMarkers);
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
                        addMarkers);
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
            bool addMarkers = false)
        {
            foreach (Word.InlineShape inlineShape in document.InlineShapes)
            {
                try
                {
                    // EnhMetaFileBitsを取得
                    byte[] metaFileData = (byte[])inlineShape.Range.EnhMetaFileBits;
                    
                    if (metaFileData != null && metaFileData.Length > 0)
                    {
                        string filePath = ExtractImageFromMetaFileData(
                            metaFileData, 
                            outputDirectory, 
                            $"inline_image_{imageCounter}", 
                            inlineShape.Type.ToString());
                        
                        if (!string.IsNullOrEmpty(filePath))
                        {
                            var imageInfo = new ExtractedImageInfo(
                                filePath, 
                                $"インライン図形_{inlineShape.Type}", 
                                inlineShape.Range.Start
                            );
                            extractedImages.Add(imageInfo);

                            // マーカーを追加
                            if (addMarkers)
                            {
                                InsertMarker(inlineShape.Range, filePath);
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
            bool addMarkers = false)
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
                            addMarkers);
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
                                addMarkers);
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
                            addMarkers);
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
            bool addMarkers = false)
        {
            try
            {
                // キャンバス全体を画像として抽出
                canvas.Select();
                byte[] canvasData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                
                if (canvasData != null && canvasData.Length > 0)
                {
                    string filePath = ExtractImageFromMetaFileData(
                        canvasData, 
                        outputDirectory, 
                        $"canvas_{imageCounter}", 
                        "Canvas");
                    
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        var imageInfo = new ExtractedImageInfo(
                            filePath, 
                            "キャンバス", 
                            canvas.Anchor?.Start ?? 0
                        );
                        extractedImages.Add(imageInfo);

                        // マーカーを追加（キャンバスの場合は異なるアプローチを使用）
                        if (addMarkers)
                        {
                            if (canvas.Anchor != null)
                            {
                                InsertMarkerAtPosition(canvas.Anchor, filePath);
                            }
                            else
                            {
                                // Anchorが利用できない場合、キャンバスが選択された状態で
                                // 現在の選択範囲を使用してマーカーを挿入
                                InsertMarkerForSelectedCanvas(filePath);
                            }
                        }

                        imageCounter++;
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
                        addMarkers);
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
            bool addMarkers = false)
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
                        addMarkers);
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
            bool addMarkers = false)
        {
            try
            {
                shape.Select();
                byte[] shapeData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                
                if (shapeData != null && shapeData.Length > 0)
                {
                    string filePath = ExtractImageFromMetaFileData(
                        shapeData, 
                        outputDirectory, 
                        $"{prefix}_{imageCounter}", 
                        shape.Type.ToString());
                    
                    if (!string.IsNullOrEmpty(filePath))
                    {
                        var imageInfo = new ExtractedImageInfo(
                            filePath, 
                            $"{prefix}_{shape.Type}", 
                            shape.Anchor?.Start ?? 0
                        );
                        extractedImages.Add(imageInfo);

                        // マーカーを追加
                        if (addMarkers && shape.Anchor != null)
                        {
                            InsertMarkerAtPosition(shape.Anchor, filePath);
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
        /// EnhMetaFileBitsから画像ファイルを作成
        /// </summary>
        private static string ExtractImageFromMetaFileData(
            byte[] metaFileData, 
            string outputDirectory, 
            string baseFileName, 
            string shapeType)
        {
            try
            {
                using (var memoryStream = new MemoryStream(metaFileData))
                {
                    using (var image = Image.FromStream(memoryStream))
                    {
                        // 最小サイズのフィルタリング（必要に応じて調整）
                        if (image.Width < 10 || image.Height < 10)
                            return null;

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
                        image.Save(filePath, ImageFormat.Png);
                        
                        return filePath;
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
        /// インライン図形の次の行にマーカーテキストを挿入
        /// </summary>
        private static void InsertMarker(Word.Range range, string filePath)
        {
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);
                
                // 図形を含む段落を取得
                var paragraph = range.Paragraphs[1];
                
                // 段落の末尾に移動
                var insertRange = range.Document.Range(paragraph.Range.End - 1, paragraph.Range.End - 1);
                
                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";
                
                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = range.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;
                
                // マーカーの後に改行を追加
                var afterMarkerRange = range.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"マーカー挿入エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 指定した位置の次の行にマーカーを挿入（フローティング図形用）
        /// </summary>
        private static void InsertMarkerAtPosition(Word.Range anchor, string filePath)
        {
            try
            {
                // ファイル名からファイル名部分のみを取得（拡張子なし）
                string markerText = Path.GetFileNameWithoutExtension(filePath);
                
                // アンカー位置を含む段落を取得
                var anchorParagraph = anchor.Paragraphs[1];
                
                // 段落の末尾に移動
                var insertRange = anchor.Document.Range(anchorParagraph.Range.End - 1, anchorParagraph.Range.End - 1);
                
                // 改行を挿入して新しい行を作成
                insertRange.Text = "\r";
                
                // 新しい行に特殊な識別子を挿入（HTML出力後に置換される）
                var markerRange = anchor.Document.Range(insertRange.End, insertRange.End);
                string marker = $"[IMAGEMARKER:{markerText}]";
                markerRange.Text = marker;
                
                // マーカーの後に改行を追加
                var afterMarkerRange = anchor.Document.Range(markerRange.End, markerRange.End);
                afterMarkerRange.Text = "\r";
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"マーカー挿入エラー: {ex.Message}");
            }
        }

        /// <summary>
        /// 抽出結果の統計情報を取得
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <returns>統計情報の文字列</returns>
        public static string GetExtractionStatisticsWithText(List<ExtractedImageInfo> extractedImages)
        {
            if (extractedImages == null || extractedImages.Count == 0)
                return "抽出された画像はありません。";

            var statistics = new System.Text.StringBuilder();
            statistics.AppendLine($"抽出された画像数: {extractedImages.Count}");
            statistics.AppendLine();

            var groupedByType = extractedImages
                .GroupBy(img => img.ImageType)
                .OrderBy(g => g.Key);

            foreach (var group in groupedByType)
            {
                statistics.AppendLine($"{group.Key}: {group.Count()}個");
            }

            return statistics.ToString();
        }

        /// <summary>
        /// 抽出結果をテキストファイルに出力
        /// </summary>
        /// <param name="extractedImages">抽出された画像情報のリスト</param>
        /// <param name="outputPath">出力ファイルパス</param>
        public static void ExportImageInfoToTextFile(
            List<ExtractedImageInfo> extractedImages, string outputPath)
        {
            try
            {
                using (var writer = new StreamWriter(outputPath, false, System.Text.Encoding.UTF8))
                {
                    writer.WriteLine("抽出された画像の一覧");
                    writer.WriteLine("==================");
                    writer.WriteLine();

                    foreach (var image in extractedImages.OrderBy(img => img.Position))
                    {
                        writer.WriteLine($"ファイル名: {Path.GetFileName(image.FilePath)}");
                        writer.WriteLine($"種別: {image.ImageType}");
                        writer.WriteLine($"位置: {image.Position}");
                        writer.WriteLine();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"テキストファイル出力エラー: {ex.Message}");
            }
        }
    }
}
