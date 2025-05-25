using Microsoft.Office.Tools.Ribbon;
using System;
using System.IO;
using System.Text;
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
                        var coverInfo = CollectInfo(docCopy, application, paths, isPattern1, isPattern2, log);
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
    }
}
