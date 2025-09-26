// GenerateHTMLButton.Helper.cs

using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // HTML生成処理の前処理を実行
        // イベントハンドラの無効化、ボタンの無効化、ファイル名のパターンチェック
        // ファイル名が規定のパターンに合致しない場合はエラーメッセージを表示して処理を中断
        private bool PreProcess(Word.Application application, Word.Document activeDocument, loader load)
        {
            // イベントハンドラを一時的に無効化して処理中の干渉を防ぐ
            application.WindowSelectionChange -= Application_WindowSelectionChange;
            button3.Enabled = false;
            application.DocumentChange -= Application_DocumentChange;
            
            // ファイル名が規定のパターンに合致するかチェック
            if (!Regex.IsMatch(activeDocument.Name, FileNamePattern))
            {
                // パターンに合致しない場合はローダーを閉じてエラーメッセージを表示
                load.Close();
                load.Dispose();
                string ErrMsgInvalidFileName = "開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」";
                string ErrMsgFileNameRule = "ファイル命名規則エラー";
                MessageBox.Show(ErrMsgInvalidFileName, ErrMsgFileNameRule, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            return true;
        }

        // Wordドキュメントのカスタムドキュメントプロパティから「webHelpFolderName」の値を取得
        // この値はHTML出力時のフォルダ名として使用される
        private string GetWebHelpFolderName(Document activeDocument)
        {
            // カスタムドキュメントプロパティのコレクションを取得
            var properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
            
            // LINQ を使用して「webHelpFolderName」プロパティを検索し、その値を返す
            // プロパティが存在しない場合はnullを返す
            return properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                .FirstOrDefault(x => x.Name == "webHelpFolderName")?.Value;
        }

        // 例外処理とエラーログの出力
        // ローダーの終了、詳細なエラー情報のログ出力、ユーザーへのエラーメッセージ表示
        private void HandleException(Exception ex, StreamWriter log, loader load)
        {
            // ローダーを閉じてリソースを解放
            load.Close();
            load.Dispose();
            
            // 詳細なスタックトレース情報を取得
            StackTrace stackTrace = new StackTrace(ex, true);
            
            // エラーの詳細情報をログファイルに出力
            log.WriteLine("[Error] Exception Details:");
            log.WriteLine($"  Source: {ex.Source ?? "Unknown Source"}");
            log.WriteLine($"  TargetSite: {ex.TargetSite}");
            log.WriteLine($"  Message: {ex.Message}");
            log.WriteLine($"  StackTrace: {stackTrace}");
            
            // ユーザーにエラーメッセージを表示
            MessageBox.Show(ErrMsg);
        }

        // 画像をコピーして一時フォルダを削除（元のコード）
        public void CopyAndDeleteTemporaryImages(string tmpFolder, string rootPath, string exportDir)
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
                catch (Exception)
                {
                    throw;
                }
            }
        }

        // テンプレートZIPファイルをアセンブリから取得し、指定されたパスに解凍
        public void PrepareHtmlTemplates(System.Reflection.Assembly assembly, string rootPath, string exportDirPath)
        {
            string zipFilePath = null;
            try
            {
                // テンプレートZIPファイルのパス
                zipFilePath = Path.Combine(rootPath, "htmlTemplates.zip");
                string templatesDirPath = Path.Combine(rootPath, "htmlTemplates");
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
            }
            catch (Exception ex)
            {
                string ErrMsgTemplatePreparation = "テンプレートの準備中にエラーが発生しました: ";
                MessageBox.Show(ErrMsgTemplatePreparation + ex.Message, ErrMsg, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // 例外発生時でもZIPファイルを確実に削除
                if (!string.IsNullOrEmpty(zipFilePath) && File.Exists(zipFilePath))
                {
                    try
                    {
                        File.Delete(zipFilePath);
                    }
                    catch (Exception deleteEx)
                    {
                        Debug.WriteLine($"ZIPファイル削除エラー: {deleteEx.Message}");
                    }
                }
            }
        }

        // 表紙選択ダイアログを表示し、選択されたテンプレートに応じてフラグを設定
        public bool HandleCoverSelection(loader load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2, out bool isPattern3)
        {
            isEasyCloud = false;
            isEdgeTracker = false;
            isPattern1 = false;
            isPattern2 = false;
            isPattern3 = false;

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

            isEasyCloud = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EasyCloud;
            isEdgeTracker = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EdgeTracker;
            isPattern1 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern1;
            isPattern2 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern2;
            isPattern3 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern3;

            return true;
        }

        // キャンバスに関連する図形のプロパティを調整
        // キャンバス内の図形の位置調整とレイアウト最適化を実行
        // Word文書内のキャンバス図形を対象に、HTML出力前の図形調整を実施
        public void AdjustCanvasShapes(Document docCopy)
        {
            // 文書内のすべての図形をループ処理
            foreach (Shape docS in docCopy.Shapes)
            {
                // キャンバス図形のみを処理対象とする
                if (docS.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                {
                    // テーブル内のキャンバスは処理をスキップ（レイアウト崩れを防ぐため）
                    if (docS.Anchor != null && docS.Anchor.Tables.Count > 0)
                    {
                        continue;
                    }

                    // キャンバス内の各アイテムの元の位置・サイズ情報を保存
                    List<float> canvasItemsTop = new List<float>();
                    List<float> canvasItemsLeft = new List<float>();
                    List<float> canvasItemsHeight = new List<float>();
                    List<float> canvasItemsWidth = new List<float>();

                    // キャンバス内の各アイテムのプロパティを取得し保存
                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        // アスペクト比ロックを解除して調整を可能にする
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        canvasItemsTop.Add(docS.CanvasItems[i].Top);
                        canvasItemsLeft.Add(docS.CanvasItems[i].Left);
                        canvasItemsHeight.Add(docS.CanvasItems[i].Height);
                        canvasItemsWidth.Add(docS.CanvasItems[i].Width);
                    }

                    // キャンバス自体のサイズ調整（高さを30ポイント拡張）
                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                    docS.Height = docS.Height + 30;
                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;

                    // 保存した位置・サイズ情報を基に各アイテムを再配置
                    for (int i = 1; i <= docS.CanvasItems.Count; i++)
                    {
                        // アスペクト比ロックを解除して位置・サイズを調整
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                        // 元のサイズを復元
                        docS.CanvasItems[i].Height = canvasItemsHeight[i - 1];
                        docS.CanvasItems[i].Width = canvasItemsWidth[i - 1];
                        // 上方向に0.59ポイント移動（レイアウト調整のため）
                        docS.CanvasItems[i].Top = canvasItemsTop[i - 1] + 0.59F;
                        docS.CanvasItems[i].Left = canvasItemsLeft[i - 1];
                        // アスペクト比ロックを再設定
                        docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                    }
                }
            }
        }

        // Webヘルプ配布用のZIPアーカイブファイルを生成
        // HTML出力ファイル、ヘッダーファイル、元のWordファイルを含む配布パッケージを作成
        public void GenerateZipArchive(string zipDirPath, string rootPath, string exportDir, string headerDir, string docFullName, string docName)
        {
            // 各種パスの組み立て
            string exportDirPath = Path.Combine(rootPath, exportDir);
            string headerDirPath = Path.Combine(rootPath, headerDir);

            // 既存のZIP作業ディレクトリがあれば削除し、新規作成
            if (Directory.Exists(zipDirPath))
            {
                Directory.Delete(zipDirPath, true);
            }
            Directory.CreateDirectory(zipDirPath);

            // HTML出力ディレクトリをZIP作業ディレクトリにコピー
            copyDirectory(exportDirPath, Path.Combine(zipDirPath, exportDir));
            
            // ヘッダーファイルディレクトリが存在すればコピー（目次作成に使用）
            if (Directory.Exists(headerDirPath))
            {
                copyDirectory(headerDirPath, Path.Combine(zipDirPath, headerDir));
            }
            
            // 元のWordファイルをZIP作業ディレクトリにコピー
            File.Copy(docFullName, Path.Combine(zipDirPath, docName));

            // ファイルコピー情報をログに記録
            //log.WriteLine(docFullName + ":" + Path.Combine(zipDirPath, docName));

            // 既存のZIPファイルがあれば削除
            if (File.Exists(zipDirPath + ".zip"))
            {
                File.Delete(zipDirPath + ".zip");
            }

            // ZIP作業ディレクトリからZIPアーカイブを作成（Shift_JISエンコーディング使用）
            ZipFile.CreateFromDirectory(zipDirPath, zipDirPath + ".zip", CompressionLevel.Optimal, true, Encoding.GetEncoding("Shift_JIS"));

            // ZIP作業ディレクトリを削除（クリーンアップ）
            Directory.Delete(zipDirPath, true);
        }

        // HTML変換されたXML文書の本文部分において、最初の要素にデフォルトIDを設定
        private void SetDefaultBodyId(XmlDocument objBody, string docid)
        {
            if (((XmlElement)objBody.DocumentElement.FirstChild).GetAttribute("id") == "")
            {
                ((XmlElement)objBody.DocumentElement.FirstChild).SetAttribute("id", docid + "00000");
            }
        }

        // HTMLファイルをブラウザで開くかどうかを確認するダイアログを表示
        private void ShowHtmlOutputDialog(string exportDirPath, string indexHtmlPath)
        {
            string MsgHtmlOutputSuccess1 = "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？";
            string MsgHtmlOutputSuccess2 = "HTML出力成功。";
            string ErrMsgHtmlOutputFailure1 = "HTMLの出力に失敗しました。";
            string ErrMsgHtmlOutputFailure2 = "HTML出力失敗。";

            DialogResult selectMsg = MessageBox.Show(exportDirPath + MsgHtmlOutputSuccess1, MsgHtmlOutputSuccess2, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (selectMsg == DialogResult.Yes)
            {
                try { Process.Start(indexHtmlPath); }
                catch { MessageBox.Show(ErrMsgHtmlOutputFailure1, ErrMsgHtmlOutputFailure2, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }
    }
}
