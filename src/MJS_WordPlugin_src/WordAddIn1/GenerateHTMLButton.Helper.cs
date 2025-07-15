using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

// リファクタリング完了
namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool PreProcess(Word.Application application, Word.Document activeDocument, loader load)
        {
            application.WindowSelectionChange -= Application_WindowSelectionChange;
            button3.Enabled = false;
            application.DocumentChange -= Application_DocumentChange;
            if (!Regex.IsMatch(activeDocument.Name, FileNamePattern))
            {
                load.Close();
                load.Dispose();
                MessageBox.Show(ErrMsgInvalidFileName, ErrMsgFileNameRule, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        private string GetWebHelpFolderName(Word.Document activeDocument)
        {
            var properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
            return properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                .FirstOrDefault(x => x.Name == "webHelpFolderName")?.Value;
        }

        private void HandleException(Exception ex, StreamWriter log, loader load)
        {
            load.Close();
            load.Dispose();
            StackTrace stackTrace = new StackTrace(ex, true);
            log.WriteLine("[Error] Exception Details:");
            log.WriteLine($"  Source: {ex.Source ?? "Unknown Source"}");
            log.WriteLine($"  TargetSite: {ex.TargetSite}");
            log.WriteLine($"  Message: {ex.Message}");
            log.WriteLine($"  StackTrace: {stackTrace}");
            MessageBox.Show(ErrMsg);
        }

        // 画像をコピーして一時フォルダを削除（元のコード）
        public void CopyAndDeleteTemporaryImages(string tmpFolder, string rootPath, string exportDir, StreamWriter log)
        {
            if (Directory.Exists(tmpFolder))
            {
                try
                {
                    // 一時フォルダ内のすべての画像ファイルをコピー
                    foreach (string pict in Directory.GetFiles(tmpFolder))
                    {
                        File.Copy(pict, Path.Combine(rootPath, exportDir, "pict", Path.GetFileName(pict)));
                    }

                    // 一時フォルダを削除
                    Directory.Delete(tmpFolder, true);
                }
                catch (Exception ex)
                {
                    log.WriteLine($"画像フォルダのコピー中にエラーが発生しました: {ex.Message}");
                    throw;
                }
            }
        }

        //public void CopyAndDeleteTemporaryImages(string tmpFolder, string rootPath, string exportDir, StreamWriter log)
        //{
        //    if (Directory.Exists(tmpFolder))
        //    {
        //        try
        //        {
        //            // pict フォルダのパス
        //            string pictFolder = Path.Combine(Path.GetDirectoryName(tmpFolder), "pict");
        //            // 既存の pict フォルダがあれば削除
        //            if (Directory.Exists(pictFolder))
        //            {
        //                Directory.Delete(pictFolder, true);
        //            }
        //            // tmpFolder を pict にリネーム
        //            Directory.Move(tmpFolder, pictFolder);

        //            // pict を exportDir 配下に移動
        //            string exportPictDir = Path.Combine(rootPath, exportDir, "pict");
        //            // 既存の exportDir/pict があれば削除
        //            if (Directory.Exists(exportPictDir))
        //            {
        //                Directory.Delete(exportPictDir, true);
        //            }
        //            // exportDir がなければ作成
        //            string exportDirPath = Path.Combine(rootPath, exportDir);
        //            if (!Directory.Exists(exportDirPath))
        //            {
        //                Directory.CreateDirectory(exportDirPath);
        //            }
        //            Directory.Move(pictFolder, exportPictDir);
        //        }
        //        catch (Exception ex)
        //        {
        //            log.WriteLine($"画像フォルダの移動中にエラーが発生しました: {ex.Message}");
        //            throw;
        //        }
        //    }
        //}

        // 指定されたパスにある書誌情報を読み込み、mergeScript にデータを追加
        
        // テンプレートZIPファイルをアセンブリから取得し、指定されたパスに解凍
        public void PrepareHtmlTemplates(System.Reflection.Assembly assembly, string rootPath, string exportDir)
        {
            try
            {
                // テンプレートZIPファイルのパス
                string zipFilePath = Path.Combine(rootPath, "htmlTemplates.zip");
                string templatesDirPath = Path.Combine(rootPath, "htmlTemplates");
                string exportDirPath = Path.Combine(rootPath, exportDir);
                string tmpCoverPicDirPath = Path.Combine(rootPath, "tmpcoverpic");
                string FileNotFoundExceptionMsg = "リソース 'htmlTemplates.zip' が見つかりません。";

                // アセンブリからリソースを取得し、テンプレートZIPファイルを作成
                using (Stream stream = assembly.GetManifestResourceStream("WordAddIn1.htmlTemplates.zip"))
                {
                    if (stream == null)
                    {
                        throw new FileNotFoundException(FileNotFoundExceptionMsg);
                    }

                    using (FileStream fs = File.Create(zipFilePath))
                    {
                        stream.CopyTo(fs);
                    }
                }

                // 既存のテンプレートフォルダを削除
                if (Directory.Exists(templatesDirPath))
                {
                    Directory.Delete(templatesDirPath, true);
                }

                // ZIPファイルを解凍
                ZipFile.ExtractToDirectory(zipFilePath, rootPath);

                // 出力ディレクトリを削除
                if (Directory.Exists(exportDirPath))
                {
                    Directory.Delete(exportDirPath, true);
                }

                // 一時的なカバーピクチャフォルダを削除
                if (Directory.Exists(tmpCoverPicDirPath))
                {
                    Directory.Delete(tmpCoverPicDirPath, true);
                }

                // テンプレートフォルダを出力ディレクトリに移動
                Directory.Move(templatesDirPath, exportDirPath);

                // ZIPファイルを削除
                File.Delete(zipFilePath);
            }
            catch (Exception ex)
            {
                string ErrMsgTemplatePreparation = "テンプレートの準備中にエラーが発生しました: ";
                MessageBox.Show(ErrMsgTemplatePreparation + ex.Message, ErrMsg, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        // 表紙選択ダイアログを表示し、選択されたテンプレートに応じてフラグを設定
        public bool HandleCoverSelection(loader load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2)
        {
            isEasyCloud = false;
            isEdgeTracker = false;
            isPattern1 = false;
            isPattern2 = false;

            CoverSelectionForm coverSelectionForm = new CoverSelectionForm();
            load.Visible = false;
            coverSelectionForm.ShowDialog();

            if (coverSelectionForm.DialogResult != DialogResult.OK)
            {
                load.Close();
                load.Dispose();
                return false;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.None)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("表紙のパターンをを選択してください。");
                return false;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern3)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("[汎用パターン3]テンプレートはまもなく登場します。");
                return false;
            }

            isEasyCloud = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EasyCloud;
            isEdgeTracker = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EdgeTracker;
            isPattern1 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern1;
            isPattern2 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern2;

            return true;
        }

        // キャンバスに関連する図形のプロパティを調整
        public void AdjustCanvasShapes(Document docCopy)
        {
            foreach (Shape docS in docCopy.Shapes)
            {
                if (docS.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                {
                    List<float> canvasItemsTop = new List<float>();
                    List<float> canvasItemsLeft = new List<float>();
                    List<float> canvasItemsHeight = new List<float>();
                    List<float> canvasItemsWidth = new List<float>();

                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        canvasItemsTop.Add(docS.CanvasItems[i].Top);
                        canvasItemsLeft.Add(docS.CanvasItems[i].Left);
                        canvasItemsHeight.Add(docS.CanvasItems[i].Height);
                        canvasItemsWidth.Add(docS.CanvasItems[i].Width);
                    }

                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    docS.Height = docS.Height + 30;
                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        docS.CanvasItems[i].Height = canvasItemsHeight[i - 1];
                        docS.CanvasItems[i].Width = canvasItemsWidth[i - 1];
                        docS.CanvasItems[i].Top = canvasItemsTop[i - 1] + 0.59F;
                        docS.CanvasItems[i].Left = canvasItemsLeft[i - 1];
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    }
                }
            }
        }

        public void GenerateZipArchive(string zipDirPath, string rootPath, string exportDir, string headerDir, string docFullName, string docName, StreamWriter log)
        {
            string exportDirPath = Path.Combine(rootPath, exportDir);
            string headerDirPath = Path.Combine(rootPath, headerDir);

            if (Directory.Exists(zipDirPath))
            {
                Directory.Delete(zipDirPath, true);
            }
            Directory.CreateDirectory(zipDirPath);

            copyDirectory(exportDirPath, Path.Combine(zipDirPath, exportDir));
            if (Directory.Exists(headerDirPath))
            {
                copyDirectory(headerDirPath, Path.Combine(zipDirPath, headerDir));
            }
            File.Copy(docFullName, Path.Combine(zipDirPath, docName));

            log.WriteLine(docFullName + ":" + Path.Combine(zipDirPath, docName));

            if (File.Exists(zipDirPath + ".zip"))
            {
                File.Delete(zipDirPath + ".zip");
            }

            ZipFile.CreateFromDirectory(zipDirPath, zipDirPath + ".zip", CompressionLevel.Optimal, true, Encoding.GetEncoding("Shift_JIS"));

            Directory.Delete(zipDirPath, true);
        }

        private void SetDefaultBodyId(XmlDocument objBody, string docid)
        {
            if (((XmlElement)objBody.DocumentElement.FirstChild).GetAttribute("id") == "")
            {
                ((XmlElement)objBody.DocumentElement.FirstChild).SetAttribute("id", docid + "00000");
            }
        }

        private void ShowHtmlOutputDialog(string exportDirPath, string indexHtmlPath)
        {
            DialogResult selectMsg = MessageBox.Show(exportDirPath + MsgHtmlOutputSuccess1, MsgHtmlOutputSuccess2, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectMsg == DialogResult.Yes)
            {
                try { Process.Start(indexHtmlPath); }
                catch { MessageBox.Show(ErrMsgHtmlOutputFailure1, ErrMsgHtmlOutputFailure2, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }
    }
}
