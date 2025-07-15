using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using DocumentFormat.OpenXml.VariantTypes;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void GenerateHTMLButton(object sender, RibbonControlEventArgs e)
        {
            // HTML出力フラグをON
            blHTMLPublish = true;

            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            // 現在の表示モードを保存
            var defaultView = application.ActiveWindow.View.Type;
            
            // ローダーフォームを表示
            loader load = new loader();
            load.Show();

            try
            {
                // 前処理（ドキュメントや環境のチェック）
                if (!PreProcess(application, activeDocument, load)) return;

                // 出力先フォルダ名を取得
                var webHelpFolderName = GetWebHelpFolderName(activeDocument);
                
                // 書籍情報の作成・取得
                if (!makeBookInfo(load)) { load.Close(); load.Dispose(); return; }
                
                // マージスクリプト情報の収集
                var mergeScript = CollectMergeScriptDict(activeDocument);
                
                // カバー選択ダイアログの処理
                if (!HandleCoverSelection(load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2)) return;
                
                // ローダーを可視化
                load.Visible = true;
                
                // すべての変更履歴を反映
                activeDocument.AcceptAllRevisions();
                
                // 各種パスの準備
                var paths = PreparePaths(activeDocument, webHelpFolderName);
                
                // ログファイルの作成
                using (StreamWriter log = new StreamWriter(paths.logPath, false, Encoding.UTF8))
                {
                    bool isError = false;
                    try
                    {
                        // アセンブリ取得
                        System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                        
                        // HTMLテンプレートの準備
                        PrepareHtmlTemplates(assembly, paths.rootPath, paths.exportDir);
                        Application.DoEvents();
                        
                        // ドキュメントを一時HTML用にコピー
                        var docCopy = CopyDocumentToHtml(application, log);

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

                        // HTML保存時のエンコーディング設定
                        docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

                        // 表紙情報（タイトル・サブタイトル等）の収集
                        CollectCoverParagraphs(docCopy, ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter, ref coverExist);

                        // 商標・著作権情報の収集
                        CollectTrademarkAndCopyrightDetails(docCopy, lastSectionIdx, log, ref trademarkTitle, ref trademarkTextList, ref trademarkRight);

                        // タイトル・サブタイトル等の整形
                        CleanUpManualTitles(ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter);

                        List<List<string>> productSubLogoGroups = new List<List<string>>();



                        // カバー情報の収集
                        //var coverInfo = CollectInfo(docCopy, application, paths, isPattern1, isPattern2, log);

                        // HTMLファイルの読み込みと加工
                        var htmlStr = ReadAndProcessHtml(paths.tmpHtmlPath, coverInfo.isTmpDot);
                        
                        // XMLへの変換と各種ノード取得
                        var (objXml, objToc, objBody) = LoadAndProcessXml(htmlStr, coverInfo.docTitle);
                        
                        // CSSスタイルの処理
                        var (className, styleName, chapterSplitClass) = ProcessCssStyles(objXml);
                        
                        // index.htmlの書き出し
                        WriteIndexHtml(paths.indexHtmlPath, coverInfo.docTitle, coverInfo.docid, mergeScript);
                        
                        // カバーテンプレートの生成
                        var (htmlCoverTemplate1, htmlCoverTemplate2) = BuildCoverTemplates(assembly, paths, coverInfo, isEasyCloud, isEdgeTracker, isPattern1, isPattern2);
                        
                        // HTMLテンプレートの生成
                        var htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);
                        var htmlTemplate2 = "</body>\n</html>\n";
                        
                        // 検索用JSの生成
                        var searchJs = BuildSearchJs();
                        
                        // 目次・本文ノードの参照取得
                        XmlNode objTocCurrent = objToc.DocumentElement;
                        XmlNode objBodyCurrent = objBody.DocumentElement;
                        
                        // 目次・本文の生成
                        BuildTocBodyFromXml(objXml, objBody, objToc, chapterSplitClass, styleName, coverInfo.docid, bookInfoDef, ref objBodyCurrent, ref objTocCurrent, load);
                        
                        // 本文IDの設定
                        SetDefaultBodyId(objBody, coverInfo.docid);
                        
                        // 目次ファイルの生成
                        ExportTocAsJsFiles(objToc, paths.rootPath, paths.exportDir, mergeScript);
                        
                        // 一時XMLの解放
                        objXml = null;
                        
                        // 一時HTMLの削除
                        File.Delete(paths.tmpHtmlPath);
                        
                        // XMLノードのクリーンアップ
                        CleanUpXmlNodes(objBody);
                        
                        // 検索用ファイルの生成
                        GenerateSearchFiles(objBody, paths.rootPath, paths.exportDir, coverInfo.docid, htmlTemplate1, htmlTemplate2, htmlCoverTemplate1, htmlCoverTemplate2, objToc, mergeScript, searchJs);

                        // AppData/Local/Tempから画像をwebhelpフォルダにコピーする
                        CopyImagesFromAppDataLocalTemp(activeDocument.FullName);

                        // Zipファイル作成ログ
                        log.WriteLine("Zipファイル作成");
                        
                        // Zipアーカイブの生成
                        GenerateZipArchive(paths.zipDirPath, paths.rootPath, paths.exportDir, paths.headerDir, paths.docFullName, paths.docName, log);
                    }
                    catch (Exception ex)
                    {
                        isError = true;
                        HandleException(ex, log, load);
                        button3.Enabled = true;
                        return;
                    }
                    finally
                    {
                        log.Close();
                        if (!isError && File.Exists(paths.logPath))
                        {
                            File.Delete(paths.logPath);
                        }

                        // ドキュメント変更イベントを再登録
                        application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);

                        // tmpcoverpicのクリーンアップ
                        //var tmpCoverPicPath = Path.Combine(paths.rootPath, "tmpcoverpic");
                        //if (Directory.Exists(tmpCoverPicPath))
                        //{
                        //    try { Directory.Delete(tmpCoverPicPath, true); }
                        //    catch { /* ログ出力など必要に応じて */ }
                        //}
                    }
                }

                // 表示モードを元に戻す
                application.ActiveWindow.View.Type = defaultView;

                load.Close();
                load.Dispose();

                // 出力先フォルダをダイアログで表示
                ShowHtmlOutputDialog(paths.exportDirPath, paths.indexHtmlPath);
            }
            finally
            {
                // ドキュメント変更イベントを再登録
                application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            }
        }

        // ドキュメントを一時 HTML 用にコピー
        //private Word.Document CopyDocumentToHtml(Word.Application application, StreamWriter log)
        //{
        //    //CheckAndRestoreRefFields(application.ActiveDocument);
        //    ClearClipboardSafely();
        //    Application.DoEvents();
        //    application.Selection.WholeStory();
        //    application.Selection.Copy();
        //    Application.DoEvents();
        //    application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
        //    Application.DoEvents();
        //    Word.Document docCopy = application.Documents.Add();
        //    Application.DoEvents();
        //    docCopy.TrackRevisions = false;
        //    docCopy.AcceptAllRevisions();
        //    docCopy.Select();
        //    Application.DoEvents();
        //    application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
        //    Application.DoEvents();
        //    ClearClipboardSafely();
        //    log.WriteLine("Number of sections: " + docCopy.Sections.Count);
        //    return docCopy;
        //}


        private Word.Document CopyDocumentToHtml(Word.Application application, StreamWriter log)
        {
            // 元ドキュメントの全範囲を取得
            Word.Document srcDoc = application.ActiveDocument;
            Word.Range srcRange = srcDoc.Content;

            // 新規ドキュメントを作成
            Word.Document docCopy = application.Documents.Add();
            docCopy.TrackRevisions = false;

            // 元ドキュメントの全範囲をコピー＆ペースト（フィールドを保持）
            srcRange.Copy();
            Word.Range destRange = docCopy.Content;
            destRange.Paste();

            Application.DoEvents();

            // クリップボードをクリア（任意）
            ClearClipboardSafely();

            log.WriteLine("Number of sections: " + docCopy.Sections.Count);
            return docCopy;
        }
    }
}
