// ThisAddIn.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // EnhMetaFileBitsを使用した画像・キャンバス抽出（アクティブドキュメントから直接、テキスト情報付き） ***
            ExtractImagesAndCanvasFromActiveDocumentWithText();
        }

        // EnhMetaFileBitsを使用してアクティブWordドキュメントから画像とキャンバスを直接抽出する
        private void ExtractImagesAndCanvasFromActiveDocumentWithText()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                {
                    MessageBox.Show("アクティブドキュメントが見つかりません。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                string docDir = Path.GetDirectoryName(doc.FullName);
                string docName = Path.GetFileNameWithoutExtension(doc.FullName);
                
                // 抽出先フォルダを作成（ドキュメントと同じディレクトリに作成）
                string extractedImagesDir = Path.Combine(docDir, $"{docName}_extracted_images");
                if (!Directory.Exists(extractedImagesDir))
                    Directory.CreateDirectory(extractedImagesDir);

                // 既存の画像ファイルをクリア（必要に応じて）
                if (Directory.Exists(extractedImagesDir))
                {
                    var existingFiles = Directory.GetFiles(extractedImagesDir, "*.png");
                    foreach (var file in existingFiles)
                    {
                        try { File.Delete(file); } catch { }
                    }
                    
                    // 既存のテキストファイルもクリア
                    var existingTextFiles = Directory.GetFiles(extractedImagesDir, "*.txt");
                    foreach (var file in existingTextFiles)
                    {
                        try { File.Delete(file); } catch { }
                    }
                }

                // EnhMetaFileBitsを使用して画像・キャンバスを抽出
                List<Utils.ExtractedImageInfo> extractedImages = Utils.ExtractImagesAndCanvasFromWordWithText(
                    doc, 
                    extractedImagesDir,
                    includeInlineShapes: true,    // インライン図形を抽出
                    includeShapes: true,          // フローティング図形を抽出
                    includeCanvasItems: false     // キャンバス内アイテムは抽出しない
                );

                // 画像情報をファイルに出力
                string imageInfoPath = Path.Combine(extractedImagesDir, $"{docName}_画像情報.txt");
                Utils.ExportImageInfoToTextFile(extractedImages, imageInfoPath);

                // 抽出結果の統計情報を取得
                string statistics = Utils.GetExtractionStatisticsWithText(extractedImages);

                // 結果をユーザーに表示
                string message = $"アクティブドキュメントからのEnhMetaFileBits画像抽出が完了しました。\n\n" +
                               $"ドキュメント: {doc.Name}\n" +
                               $"抽出先: {extractedImagesDir}\n" +
                               $"画像情報ファイル: {Path.GetFileName(imageInfoPath)}\n\n" +
                               $"{statistics}\n\n" +
                               $"抽出フォルダを開きますか？";

                DialogResult result = MessageBox.Show(message, "画像抽出完了", MessageBoxButtons.YesNo, MessageBoxIcon.Information);

                if (result == DialogResult.Yes && Directory.Exists(extractedImagesDir))
                {
                    System.Diagnostics.Process.Start("explorer.exe", extractedImagesDir);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"アクティブドキュメントからの画像抽出中にエラーが発生しました: {ex.Message}\nStackTrace: {ex.StackTrace}", 
                    "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            var doc = this.Application.ActiveDocument;
            try
            {
                // アドイン終了時に全ての画像マーカーを削除
                Utils.RemoveAllImageMarkers(doc);
                
                // 必要に応じてリソースを解放
                if (Globals.ThisAddIn.Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.Application);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"アドイン終了時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        #region VSTO で生成されたコード

        /// デザイナーのサポートに必要なメソッドです。
        /// メソッドの内容をコードエディターで変更しないでください。
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}