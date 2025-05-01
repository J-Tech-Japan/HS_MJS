using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        // 指定されたパスにあるテキストファイル（書誌情報）を読み込み、mergeScript にデータを追加
        public void CollectMergeScript(string documentPath, string documentName, Dictionary<string, string> mergeScript)
        {
            using (StreamReader sr = new StreamReader(
                    documentPath + "\\headerFile\\" + Regex.Replace(documentName, "^(.{3}).+$", "$1") + @".txt", Encoding.Default))
            {
                while (sr.Peek() >= 0)
                {
                    string strBuffer = sr.ReadLine();
                    string[] info = strBuffer.Split('\t');

                    if (info.Length == 4)
                    {
                        if (!info[3].Equals(""))
                        {
                            info[3] = info[3].Replace("(", "").Replace(")", "");
                            if (!mergeScript.Any(x => x.Key == info[2] && x.Value == info[3]))
                            {
                                mergeScript.Add(info[2], info[3]);
                            }
                        }
                    }
                }
            }
        }

        public void PrepareHtmlTemplates(System.Reflection.Assembly assembly, string rootPath, string exportDir)
        {
            // アセンブリからリソースを取得し、テンプレートZIPファイルを作成
            using (Stream stream = assembly.GetManifestResourceStream("WordAddIn1.htmlTemplates.zip"))
            {
                FileStream fs = File.Create(rootPath + "\\htmlTemplates.zip");
                stream.Seek(0, SeekOrigin.Begin);
                stream.CopyTo(fs);
                fs.Close();
            }

            // 既存のテンプレートフォルダを削除
            if (Directory.Exists(rootPath + "\\htmlTemplates"))
            {
                Directory.Delete(rootPath + "\\htmlTemplates", true);
            }

            // ZIPファイルを解凍
            ZipFile.ExtractToDirectory(rootPath + "\\htmlTemplates.zip", rootPath);

            // 出力ディレクトリを削除
            if (Directory.Exists(rootPath + "\\" + exportDir))
            {
                Directory.Delete(rootPath + "\\" + exportDir, true);
            }

            // 一時的なカバーピクチャフォルダを削除
            if (Directory.Exists(rootPath + "\\tmpcoverpic"))
            {
                Directory.Delete(rootPath + "\\tmpcoverpic", true);
            }

            // テンプレートフォルダを出力ディレクトリに移動
            Directory.Move(rootPath + "\\htmlTemplates", rootPath + "\\" + exportDir);

            // ZIPファイルを削除
            File.Delete(rootPath + "\\htmlTemplates.zip");
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
            ref string trademarkRight
            )
        {
            bool isTradeMarksDetected = false;
            bool isRightDetected = false;

            foreach (Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
            {
                log.WriteLine(wp.Range.Text);

                string wpTextTrim = wp.Range.Text.Trim();
                string wpStyleName = wp.get_Style().NameLocal;

                if (string.IsNullOrEmpty(wpTextTrim) || wpTextTrim == "/")
                {
                    continue;
                }

                if (wpTextTrim.Contains("商標")
                    && (wpStyleName.Contains("MJS_見出し 4") || wpStyleName.Contains("MJS_見出し 5")))
                {
                    trademarkTitle = wp.Range.Text + "<br/>";
                    isTradeMarksDetected = true;
                }
                else if (isTradeMarksDetected && (!isRightDetected)
                    && (wpStyleName.Contains("MJS_箇条書き")
                        || wpStyleName.Contains("MJS_箇条書き2")))
                {
                    trademarkTextList.Add(wp.Range.Text + "<br/>");
                }
                else if (wpTextTrim.Contains("All rights reserved")
                    && (wpStyleName.Contains("MJS_リード文")))
                {
                    trademarkRight = wp.Range.Text + "<br/>";
                    isRightDetected = true;
                }
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
            if (Directory.Exists(zipDirPath))
            {
                Directory.Delete(zipDirPath, true);
            }
            Directory.CreateDirectory(zipDirPath);

            copyDirectory(rootPath + "\\" + exportDir, Path.Combine(zipDirPath, exportDir));
            if (Directory.Exists(rootPath + "\\" + headerDir))
            {
                copyDirectory(rootPath + "\\" + headerDir, Path.Combine(zipDirPath, headerDir));
            }
            File.Copy(docFullName, Path.Combine(zipDirPath, docName));

            log.WriteLine(docFullName + ":" + Path.Combine(zipDirPath, docName));

            if (File.Exists(zipDirPath + ".zip"))
            {
                File.Delete(zipDirPath + ".zip");
            }

            System.IO.Compression.ZipFile.CreateFromDirectory(zipDirPath, zipDirPath + ".zip", System.IO.Compression.CompressionLevel.Optimal, true, Encoding.GetEncoding("Shift_JIS"));

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
