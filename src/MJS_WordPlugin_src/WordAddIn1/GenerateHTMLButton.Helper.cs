using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// リファクタリング完了
namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 画像をコピーして一時フォルダを削除
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

        // 指定されたパスにある書誌情報を読み込み、mergeScript にデータを追加
        public void CollectMergeScript(string documentPath, string documentName, Dictionary<string, string> mergeScript)
        {
            try
            {
                // ファイルパスを安全に結合
                // ドキュメント名の最初の3文字を抽出して、対応するヘッダーファイルのパスを生成
                string headerFilePath = Path.Combine(documentPath, "headerFile", Regex.Replace(documentName, "^(.{3}).+$", "$1") + ".txt");

                // ヘッダーファイルをUTF-8エンコーディングで読み込む
                using (StreamReader sr = new StreamReader(headerFilePath, Encoding.UTF8))
                {
                    // ファイルの終端まで1行ずつ読み込む
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

                        // 空行や空白行をスキップ
                        if (string.IsNullOrWhiteSpace(strBuffer)) continue;

                        // タブ区切りで行を分割
                        string[] info = strBuffer.Split('\t');

                        // 必要な情報が揃っている場合のみ処理を続行
                        if (info.Length == 4 && !string.IsNullOrEmpty(info[3]))
                        {
                            string key = info[2]; // 辞書のキー
                            string value = info[3].Replace("(", "").Replace(")", ""); // 値から括弧を削除

                            // 辞書に同じキーと値のペアが存在しない場合のみ追加
                            if (!mergeScript.ContainsKey(key) || mergeScript[key] != value)
                            {
                                mergeScript[key] = value;
                            }
                        }
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                // ファイルが見つからない場合のエラーメッセージを表示
                MessageBox.Show($"ファイルが見つかりません: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // その他の例外が発生した場合のエラーメッセージを表示
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

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

        // 表紙に関連する段落を収集
        public void CollectCoverParagraphs(Document docCopy, ref string manualTitle, ref string manualSubTitle, ref string manualVersion,
                                      ref string manualTitleCenter, ref string manualSubTitleCenter, ref string manualVersionCenter,
                                      ref bool coverExist)
        {
            foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
            {
                if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualTitle += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualSubTitle += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualVersion += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualTitleCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualSubTitleCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualVersionCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
            }
        }

        // 商標情報と著作権情報を収集
        public void CollectTrademarkAndCopyrightDetails(
            Document docCopy,
            int lastSectionIdx,
            StreamWriter log,
            ref string trademarkTitle,
            ref List<string> trademarkTextList,
            ref string trademarkRight)
        {
            try
            {
                bool isTradeMarksDetected = false;
                bool isRightDetected = false;

                foreach (Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
                {
                    string wpTextTrim = wp.Range.Text.Trim();
                    string wpStyleName = wp.get_Style().NameLocal;

                    // ログに段落の内容を記録
                    log.WriteLine($"[Style: {wpStyleName}] {wpTextTrim}");

                    // 空行や無効な行をスキップ
                    if (string.IsNullOrEmpty(wpTextTrim) || wpTextTrim == "/")
                    {
                        continue;
                    }

                    // 商標タイトルの検出
                    if (!isTradeMarksDetected && wpTextTrim.Contains("商標") &&
                        (wpStyleName.Contains("MJS_見出し 4") || wpStyleName.Contains("MJS_見出し 5")))
                    {
                        trademarkTitle = wpTextTrim + "<br/>";
                        isTradeMarksDetected = true;
                        continue;
                    }

                    // 商標情報のリスト追加
                    if (isTradeMarksDetected && !isRightDetected &&
                        (wpStyleName.Contains("MJS_箇条書き") || wpStyleName.Contains("MJS_箇条書き2")))
                    {
                        trademarkTextList.Add(wpTextTrim + "<br/>");
                        continue;
                    }

                    // 著作権情報の検出
                    if (!isRightDetected && wpTextTrim.Contains("All rights reserved") &&
                        wpStyleName.Contains("MJS_リード文"))
                    {
                        trademarkRight = wpTextTrim + "<br/>";
                        isRightDetected = true;
                        continue;
                    }
                }
            }
            catch (Exception ex)
            {
                log.WriteLine($"エラー: {ex.Message}");
                MessageBox.Show($"商標および著作権情報の収集中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
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

        private void CleanUpManualTitles(
            ref string manualTitle,
            ref string manualSubTitle,
            ref string manualVersion,
            ref string manualTitleCenter,
            ref string manualSubTitleCenter,
            ref string manualVersionCenter)
        {
            string bell = new string((char)7, 1);
            manualTitle = Regex.Replace(manualTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitle = Regex.Replace(manualSubTitle, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersion = Regex.Replace(manualVersion, @"<br/>$", "").Replace(bell, "").Trim();
            manualTitleCenter = Regex.Replace(manualTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualSubTitleCenter = Regex.Replace(manualSubTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
            manualVersionCenter = Regex.Replace(manualVersionCenter, @"<br/>$", "").Replace(bell, "").Trim();
        }
    }
}
