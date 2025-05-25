using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private (bool coverExist, string subTitle, int biCount, List<List<string>> productSubLogoGroups, string docTitle, string docid, bool isTmpDot) CollectInfo(Document docCopy, Word.Application application, (string rootPath, string docName, string docFullName, string exportDir, string headerDir, string exportDirPath, string logPath, string tmpHtmlPath, string indexHtmlPath, string tmpFolderForImagesSavedBySaveAs2Method, string docid, string docTitle, string zipDirPath) paths, bool isPattern1, bool isPattern2, StreamWriter log)
        {
            int biCount = 0;
            bool coverExist = false;
            string subTitle = "";
            string manualTitle = "";
            string manualSubTitle = "";
            string manualVersion = "";
            string manualTitleCenter = "";
            string manualSubTitleCenter = "";
            string manualVersionCenter = "";
            string trademarkTitle = "";
            List<string> trademarkTextList = new List<string>();
            string trademarkRight = "";
            int lastSectionIdx = docCopy.Sections.Count;
            CollectCoverParagraphs(docCopy, ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter, ref coverExist);
            CollectTrademarkAndCopyrightDetails(docCopy, lastSectionIdx, log, ref trademarkTitle, ref trademarkTextList, ref trademarkRight);
            CleanUpManualTitles(ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter);
            List<List<string>> productSubLogoGroups = new List<List<string>>();
            if (coverExist)
            {
                ProcessCoverImages(docCopy, application, paths.rootPath, paths.exportDir, ref subTitle, ref biCount, ref productSubLogoGroups, isPattern1, isPattern2, log);
            }
            application.Selection.EndKey(Word.WdUnits.wdStory);
            object selectionRange = application.Selection.Range;
            Word.Shape temporaryCanvas = docCopy.Shapes.AddCanvas(0, 0, 1, 1, ref selectionRange);
            temporaryCanvas.WrapFormat.Type = Word.WdWrapType.wdWrapInline;
            AdjustCanvasShapes(docCopy);
            temporaryCanvas.Delete();
            foreach (Word.Table wt in docCopy.Tables)
            {
                if (wt.PreferredWidthType == Word.WdPreferredWidthType.wdPreferredWidthPoints)
                    wt.AllowAutoFit = true;
            }
            foreach (Word.Style ws in docCopy.Styles)
                if (ws.NameLocal == "奥付タイトル")
                    ws.NameLocal = "titledef";
            docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
            docCopy.SaveAs2(paths.tmpHtmlPath, Word.WdSaveFormat.wdFormatFilteredHTML);
            docCopy.Close();
            log.WriteLine("画像フォルダ コピー");
            bool isTmpDot = true;
            CopyAndDeleteTemporaryImages(paths.tmpFolderForImagesSavedBySaveAs2Method, paths.rootPath, paths.exportDir, log);
            return (coverExist, subTitle, biCount, productSubLogoGroups, paths.docTitle, paths.docid, isTmpDot);
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
