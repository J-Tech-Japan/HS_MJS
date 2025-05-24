using Microsoft.Office.Tools.Ribbon;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void GenerateHTMLButton(object sender, RibbonControlEventArgs e)
        {
            blHTMLPublish = true;
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            var defaultView = application.ActiveWindow.View.Type;
            loader load = new loader();
            load.Show();

            try
            {
                if (!PreProcess(application, activeDocument, load)) return;
                var webHelpFolderName = GetWebHelpFolderName(activeDocument);
                if (!makeBookInfo(load)) { load.Close(); load.Dispose(); return; }
                var mergeScript = CollectMergeScriptDict(activeDocument);
                if (!HandleCoverSelection(load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2)) return;
                load.Visible = true;
                activeDocument.AcceptAllRevisions();
                var paths = PreparePaths(activeDocument, webHelpFolderName);
                using (StreamWriter log = new StreamWriter(paths.logPath, false, Encoding.UTF8))
                {
                    try
                    {
                        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        PrepareHtmlTemplates(assembly, paths.rootPath, paths.exportDir);
                        Application.DoEvents();
                        var docCopy = CopyDocumentToHtml(application, paths.tmpHtmlPath, log);
                        var coverInfo = CollectCoverAndTrademarkInfo(docCopy, application, paths, isPattern1, isPattern2, log);
                        var htmlStr = ReadAndProcessHtml(paths.tmpHtmlPath, coverInfo.isTmpDot);
                        var (objXml, objToc, objBody) = LoadAndProcessXml(htmlStr, coverInfo.docTitle);
                        var (className, styleName, chapterSplitClass) = ProcessCssStyles(objXml);
                        WriteIndexHtml(paths.indexHtmlPath, coverInfo.docTitle, coverInfo.docid, mergeScript);
                        var (htmlCoverTemplate1, htmlCoverTemplate2) = BuildCoverTemplates(assembly, paths, coverInfo, isEasyCloud, isEdgeTracker, isPattern1, isPattern2);
                        var htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);
                        var htmlTemplate2 = "</body>\n</html>\n";
                        var searchJs = BuildSearchJs();
                        // objTocCurrent, objBodyCurrent を取得しrefで渡す
                        XmlNode objTocCurrent = objToc.DocumentElement;
                        XmlNode objBodyCurrent = objBody.DocumentElement;
                        GenerateTocAndBody(objXml, objBody, objToc, chapterSplitClass, styleName, coverInfo.docid, bookInfoDef, ref objBodyCurrent, ref objTocCurrent, load);
                        SetDefaultBodyId(objBody, coverInfo.docid);
                        GenerateTocFiles(objToc, paths.rootPath, paths.exportDir, mergeScript);
                        objXml = null;
                        File.Delete(paths.tmpHtmlPath);
                        CleanUpXmlNodes(objBody);
                        GenerateSearchFiles(objBody, paths.rootPath, paths.exportDir, coverInfo.docid, htmlTemplate1, htmlTemplate2, htmlCoverTemplate1, htmlCoverTemplate2, objToc, mergeScript, searchJs);
                        log.WriteLine("Zipファイル作成");
                        GenerateZipArchive(paths.zipDirPath, paths.rootPath, paths.exportDir, paths.headerDir, paths.docFullName, paths.docName, log);
                    }
                    catch (Exception ex)
                    {
                        HandleException(ex, log, load);
                        button3.Enabled = true;
                        return;
                    }
                }
                application.ActiveWindow.View.Type = defaultView;
                load.Close();
                load.Dispose();
                ShowHtmlOutputDialog(paths.exportDirPath, paths.indexHtmlPath);
            }
            finally
            {
                application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            }
        }

        // --- 以下、分割したプライベートメソッド群 ---
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
        private Dictionary<string, string> CollectMergeScriptDict(Word.Document activeDocument)
        {
            var mergeScript = new Dictionary<string, string>();
            CollectMergeScript(activeDocument.Path, activeDocument.Name, mergeScript);
            return mergeScript;
        }

        private Word.Document CopyDocumentToHtml(Word.Application application, string tmpHtmlPath, StreamWriter log)
        {
            ClearClipboardSafely();
            Application.DoEvents();
            application.Selection.WholeStory();
            application.Selection.Copy();
            Application.DoEvents();
            application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Application.DoEvents();
            Word.Document docCopy = application.Documents.Add();
            Application.DoEvents();
            docCopy.TrackRevisions = false;
            docCopy.AcceptAllRevisions();
            docCopy.Select();
            Application.DoEvents();
            application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
            Application.DoEvents();
            ClearClipboardSafely();
            log.WriteLine("Number of sections: " + docCopy.Sections.Count);
            return docCopy;
        }

        private (bool coverExist, string subTitle, int biCount, List<List<string>> productSubLogoGroups, string docTitle, string docid, bool isTmpDot) CollectCoverAndTrademarkInfo(Word.Document docCopy, Word.Application application, (string rootPath, string docName, string docFullName, string exportDir, string headerDir, string exportDirPath, string logPath, string tmpHtmlPath, string indexHtmlPath, string tmpFolderForImagesSavedBySaveAs2Method, string docid, string docTitle, string zipDirPath) paths, bool isPattern1, bool isPattern2, StreamWriter log)
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
        
        private void WriteIndexHtml(string indexHtmlPath, string docTitle, string docid, Dictionary<string, string> mergeScript)
        {
            string idxHtmlTemplate = BuildIdxHtmlTemplate(docTitle, docid, mergeScript);
            using (StreamWriter sw = new StreamWriter(indexHtmlPath, false, Encoding.UTF8))
            {
                sw.Write(idxHtmlTemplate);
            }
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
    }
}
