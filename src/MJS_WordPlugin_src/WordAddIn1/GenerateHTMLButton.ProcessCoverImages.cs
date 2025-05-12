/*
ProcessCoverImages メソッドでは、Word ドキュメント内の図形や画像を処理し、
特定の条件に基づいて画像を抽出・変換・保存する一連の処理が行われています。

1. 一時ディレクトリの作成
•	EnsureDirectoryExists メソッドを使用して、rootPath 配下に tmpcoverpic ディレクトリを作成します。
•	このディレクトリは、処理中に生成される画像ファイルを一時的に保存するために使用されます。

2. 図形のグループ解除
•	UngroupShapes メソッドを呼び出し、Word ドキュメント内の図形 (Shapes) を再帰的にグループ解除します。
•	グループ化された図形を個別の図形として扱えるようにします。

3. キャンバス内のコンテンツ抽出
•	ExtractCanvasContent メソッドを呼び出し、キャンバス (msoCanvas) 内のテキストや画像を処理します。
•	テキスト抽出: キャンバス内の図形からテキストを抽出し、subTitle に格納します。
•	画像保存: キャンバス内の画像を PNG 形式で保存します。

4. 図形をインライン図形に変換
•	ConvertPicturesToInlineShapes メソッドを呼び出し、msoPicture タイプの図形をインライン図形に変換します。
•	インライン図形にすることで、後続の処理で扱いやすくします。

5. ロゴの抽出
•	ExtractLogos メソッドを呼び出し、特定のスタイル（例: MJS_製品ロゴ（メイン） や MJS_製品ロゴ（サブ））を持つ段落からロゴ画像を抽出します。
•	メインロゴ: 最初のロゴ画像を抽出し、product_logo_main.png として保存。
•	サブロゴ: 最大 3 つのサブロゴを抽出し、それぞれ別のファイルに保存。

6. インライン図形の画像抽出
•	ExtractInlineShapes メソッドを呼び出し、インライン図形から画像を抽出します。
•	特定の条件（例: アスペクト比や高さ）を満たす画像のみを PNG 形式で保存します。

7. 画像のエクスポート
•	ProcessCoverImagesForExport メソッドを呼び出し、抽出した画像をエクスポートします。
•	パターン 1 または 2 の場合: 画像を特定のディレクトリに移動。
•	それ以外の場合: 画像をリサイズして保存。

8. 一時ディレクトリの削除
•	CleanupTemporaryDirectory メソッドを呼び出し、処理が完了した後に一時ディレクトリを削除します。

9. 例外処理
•	処理中に例外が発生した場合、log にエラーメッセージを記録します。

処理の全体像
このメソッドは、以下のようなシナリオで使用されることを想定しています：
1.	Word ドキュメント内の図形や画像を解析。
2.	必要な画像やテキストを抽出。
3.	抽出したデータを特定の形式やディレクトリ構造で保存。
*/

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        public void ProcessCoverImages(
            Word.Document docCopy,
            Word.Application application,
            string rootPath,
            string exportDir,
            ref string subTitle,
            ref int biCount,
            ref List<List<string>> productSubLogoGroups,
            bool isPattern1,
            bool isPattern2,
            StreamWriter log)
        {
            string tempDir = Path.Combine(rootPath, "tmpcoverpic");
            EnsureDirectoryExists(tempDir);

            try
            {
                UngroupShapes(docCopy, application);
                ExtractCanvasContent(docCopy, application, tempDir, ref subTitle, ref biCount);
                ConvertPicturesToInlineShapes(docCopy, application);

                if (isPattern1 || isPattern2)
                {
                    ExtractLogos(docCopy, application, tempDir, ref productSubLogoGroups, log);
                }
                else
                {
                    ExtractInlineShapes(docCopy, tempDir, ref biCount);
                }

                ProcessCoverImagesForExport(tempDir, rootPath, exportDir, isPattern1, isPattern2);
                CleanupTemporaryDirectory(tempDir);
            }
            catch (Exception ex)
            {
                log.WriteLine($"Error in ProcessCoverImages: {ex}");
            }
        }

        private void EnsureDirectoryExists(string path)
        {
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
        }

        private void UngroupShapes(Word.Document docCopy, Word.Application application)
        {
            bool repeatUngroup;
            do
            {
                repeatUngroup = false;
                foreach (Word.Shape shape in docCopy.Shapes)
                {
                    if (IsInFirstSection(shape, application) && shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        shape.Ungroup();
                        repeatUngroup = true;
                    }
                }
            } while (repeatUngroup);
        }

        private void ExtractCanvasContent(
            Word.Document docCopy,
            Word.Application application,
            string tempDir,
            ref string subTitle,
            ref int biCount)
        {
            foreach (Word.Shape shape in docCopy.Shapes)
            {
                if (IsInFirstSection(shape, application) && shape.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                {
                    UngroupCanvasItems(shape);
                    ExtractTextFromCanvas(shape, ref subTitle);

                    if (string.IsNullOrEmpty(subTitle))
                    {
                        SaveCanvasImage(shape, application, tempDir, ref biCount);
                    }
                }
            }
        }

        // 図形や画像がグループ化されている場合、解除する
        private void UngroupCanvasItems(Word.Shape canvas)
        {
            bool checkCanvas;
            do
            {
                checkCanvas = false;
                foreach (Word.Shape item in canvas.CanvasItems)
                {
                    if (item.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        item.Ungroup();
                        checkCanvas = true;
                    }
                }
            } while (checkCanvas);
        }

        // キャンバス内の図形からテキストを抽出
        private void ExtractTextFromCanvas(Word.Shape canvas, ref string subTitle)
        {
            foreach (Word.Shape item in canvas.CanvasItems)
            {
                try
                {
                    string text = item.TextFrame.TextRange.Text;
                    if (!string.IsNullOrEmpty(text) && text != "/" && string.IsNullOrEmpty(subTitle))
                    {
                        subTitle = text;
                        break;
                    }
                }
                catch
                {
                    // Ignore exceptions when accessing text
                }
            }
        }

        // キャンバス内の画像を保存
        private void SaveCanvasImage(Word.Shape canvas, Word.Application application, string tempDir, ref int biCount)
        {
            byte[] imageData = (byte[])application.Selection.EnhMetaFileBits;
            if (imageData != null)
            {
                using (MemoryStream ms = new MemoryStream(imageData))
                {
                    Image image = Image.FromStream(ms);
                    float aspectRatio = (float)image.Width / image.Height;
                    if (aspectRatio > 2.683 || aspectRatio < 2.681)
                    {
                        biCount++;
                        string filePath = Path.Combine(tempDir, $"{biCount}.png");
                        image.Save(filePath, ImageFormat.Png);
                    }
                }
            }
        }

        // 図形をインライン図形に変換
        private void ConvertPicturesToInlineShapes(Word.Document docCopy, Word.Application application)
        {
            foreach (Word.Shape shape in docCopy.Shapes)
            {
                if (IsInFirstSection(shape, application) && shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                {
                    shape.ConvertToInlineShape();
                }
            }
        }

        // ロゴを抽出
        private void ExtractLogos(
            Word.Document docCopy,
            Word.Application application,
            string tempDir,
            ref List<List<string>> productSubLogoGroups,
            StreamWriter log)
        {
            int productSubLogoCount = 0;

            foreach (Word.Paragraph paragraph in docCopy.Sections[1].Range.Paragraphs)
            {
                string styleName = paragraph.get_Style().NameLocal;
                if (styleName == "MJS_製品ロゴ（メイン）")
                {
                    ExtractMainLogo(paragraph, application, tempDir, log);
                }
                else if (styleName == "MJS_製品ロゴ（サブ）" && productSubLogoCount < 3)
                {
                    ExtractSubLogos(paragraph, application, tempDir, ref productSubLogoGroups, ref productSubLogoCount, log);
                }
            }
        }

        // メインロゴを抽出
        private void ExtractMainLogo(Word.Paragraph paragraph, Word.Application application, string tempDir, StreamWriter log)
        {
            try
            {
                foreach (Word.InlineShape inlineShape in paragraph.Range.InlineShapes)
                {
                    inlineShape.Range.Select();
                    Clipboard.Clear();
                    application.Selection.CopyAsPicture();
                    Image image = Clipboard.GetImage();
                    string filePath = Path.Combine(tempDir, "product_logo_main.png");
                    image.Save(filePath, ImageFormat.Png);
                    break; // Only extract the first main logo
                }
            }
            catch (Exception ex)
            {
                log.WriteLine($"Error extracting main logo: {ex}");
            }
        }

        // サブロゴを抽出
        private void ExtractSubLogos(
            Word.Paragraph paragraph,
            Word.Application application,
            string tempDir,
            ref List<List<string>> productSubLogoGroups,
            ref int productSubLogoCount,
            StreamWriter log)
        {
            try
            {
                List<string> subLogoFileNames = new List<string>();

                foreach (Word.InlineShape inlineShape in paragraph.Range.InlineShapes)
                {
                    inlineShape.Range.Select();
                    Clipboard.Clear();
                    application.Selection.CopyAsPicture();
                    Image image = Clipboard.GetImage();

                    productSubLogoCount++;
                    string fileName = $"product_logo_sub{productSubLogoCount}.png";
                    string filePath = Path.Combine(tempDir, fileName);
                    image.Save(filePath, ImageFormat.Png);
                    subLogoFileNames.Add(fileName);

                    Clipboard.Clear();

                    if (productSubLogoCount == 3)
                    {
                        break; // Limit to 3 sub logos
                    }
                }

                productSubLogoGroups.Add(subLogoFileNames);
            }
            catch (Exception ex)
            {
                log.WriteLine($"Error extracting sub logos: {ex}");
            }
        }

        // インライン図形から画像を抽出
        private void ExtractInlineShapes(Word.Document docCopy, string tempDir, ref int biCount)
        {
            foreach (Word.InlineShape inlineShape in docCopy.Sections[1].Range.InlineShapes)
            {
                byte[] imageData = (byte[])inlineShape.Range.EnhMetaFileBits;
                if (imageData != null)
                {
                    using (MemoryStream ms = new MemoryStream(imageData))
                    {
                        Image image = Image.FromStream(ms);
                        float aspectRatio = (float)image.Width / image.Height;
                        if (image.Height < 360 || (aspectRatio > 12.225 && aspectRatio < 12.226) || (aspectRatio > 2.681 && aspectRatio < 2.683))
                        {
                            continue;
                        }

                        biCount++;
                        string filePath = Path.Combine(tempDir, $"{biCount}.png");
                        image.Save(filePath, ImageFormat.Png);
                    }
                }
            }
        }

        // 画像をエクスポート
        private void ProcessCoverImagesForExport(string tempDir, string rootPath, string exportDir, bool isPattern1, bool isPattern2)
        {
            string[] coverPics = Directory.GetFiles(tempDir, "*.png", SearchOption.AllDirectories);
            var imageSizes = coverPics.ToDictionary(
                file => file,
                file => new FileInfo(file).Length);

            var sortedImages = imageSizes.OrderByDescending(kv => kv.Value).ToList();

            if (isPattern1 || isPattern2)
            {
                ExportImagesForPattern(sortedImages, rootPath, exportDir);
            }
            else
            {
                ExportImagesForDefault(sortedImages, rootPath, exportDir);
            }
        }

        // 画像をパターン 1 またはパターン 2 用にエクスポート
        private void ExportImagesForPattern(List<KeyValuePair<string, long>> sortedImages, string rootPath, string exportDir)
        {
            string destDir = Path.Combine(rootPath, exportDir, "template", "images");
            EnsureDirectoryExists(destDir);

            foreach (var image in sortedImages)
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(image.Key));
                if (File.Exists(destFile))
                {
                    File.Delete(destFile);
                }

                File.Move(image.Key, destFile);
            }
        }

        // 画像をデフォルト用にエクスポート
        private void ExportImagesForDefault(List<KeyValuePair<string, long>> sortedImages, string rootPath, string exportDir)
        {
            string destDir = Path.Combine(rootPath, exportDir, "template", "images");
            EnsureDirectoryExists(destDir);

            for (int i = 0; i < sortedImages.Count; i++)
            {
                string destFile;
                if (i == 0 || i + 1 != sortedImages.Count)
                {
                    destFile = Path.Combine(destDir, "cover-4.png");
                }
                else
                {
                    destFile = Path.Combine(destDir, "cover-background.png");
                    ResizeAndSaveImage(sortedImages[i].Key, destFile, 0.2f);
                }

                if (File.Exists(destFile))
                {
                    File.Delete(destFile);
                }

                File.Move(sortedImages[i].Key, destFile);
            }
        }

        // 画像をリサイズして保存
        private void ResizeAndSaveImage(string sourcePath, string destPath, float scale)
        {
            using (Bitmap src = new Bitmap(sourcePath))
            {
                int width = (int)(src.Width * scale);
                int height = (int)(src.Height * scale);
                using (Bitmap dst = new Bitmap(width, height))
                {
                    using (Graphics g = Graphics.FromImage(dst))
                    {
                        g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                        g.DrawImage(src, 0, 0, width, height);
                    }

                    dst.Save(destPath, ImageFormat.Png);
                }
            }
        }

        private void CleanupTemporaryDirectory(string tempDir)
        {
            if (Directory.Exists(tempDir))
            {
                Directory.Delete(tempDir, true);
            }
        }

        private bool IsInFirstSection(Word.Shape shape, Word.Application application)
        {
            shape.Select();
            return application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1;
        }
    }
}

