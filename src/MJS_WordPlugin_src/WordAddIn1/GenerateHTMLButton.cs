using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void GenerateHTMLButton(object sender, RibbonControlEventArgs e)
        {
            StreamWriter sw;
            blHTMLPublish = true;

            loader load = new loader();
            load.Show();

            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            // イベントハンドラの解除
            application.WindowSelectionChange -= Application_WindowSelectionChange;
            button3.Enabled = false;
            application.DocumentChange -= Application_DocumentChange;

            var defaultView = application.ActiveWindow.View.Type;

            // ファイル名の検証
            if (!Regex.IsMatch(activeDocument.Name, FileNamePattern))
            {
                load.Close();
                load.Dispose();
                MessageBox.Show(ErrMsgInvalidFileName, ErrMsgFileNameRule, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // カスタムプロパティの取得
            var properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
            string webHelpFolderName = properties.Cast<Microsoft.Office.Core.DocumentProperty>()
                                                 .FirstOrDefault(x => x.Name == "webHelpFolderName")?.Value;

            load.Visible = false;

            // BookInfoの作成
            if (!makeBookInfo(load))
            {
                load.Close();
                load.Dispose();
                return;
            }

            // マージスクリプトの収集
            Dictionary<string, string> mergeScript = new Dictionary<string, string>();
            CollectMergeScript(activeDocument.Path, activeDocument.Name, mergeScript);

            // カバー選択の処理
            bool isEasyCloud, isEdgeTracker, isPattern1, isPattern2;
            if (!HandleCoverSelection(load, out isEasyCloud, out isEdgeTracker, out isPattern1, out isPattern2))
            {
                return;
            }

            load.Visible = true;

            // ドキュメント内のすべての変更履歴を受け入れて変更を確定
            activeDocument.AcceptAllRevisions();

            string rootPath = activeDocument.Path;
            string docName = activeDocument.Name;
            string docFullName = activeDocument.FullName;
            string exportDir = "webHelp";
            string headerDir = "headerFile";
            string exportDirPath = Path.Combine(rootPath, exportDir);
            string tmpDocPath = Path.Combine(rootPath, "tmp.doc");
            string logPath = Path.Combine(rootPath, "log.txt");
            string tmpHtmlPath = Path.Combine(rootPath, "tmp.html");
            string indexHtmlPath = Path.Combine(rootPath, exportDir, "index.html");
            string tmpFolderForImagesSavedBySaveAs2Method = Path.Combine(rootPath, "tmp.files");
            string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");
            string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");
            string zipDirPath = Path.Combine(rootPath, $"{docid}_{exportDir}_{DateTime.Today:yyyyMMdd}");

            if (webHelpFolderName != null && webHelpFolderName.Length > 0)
            {
                exportDir = webHelpFolderName;
            }

            using (StreamWriter log = new StreamWriter(logPath, false, Encoding.UTF8))
            {
                try
                {
                    log.WriteLine("テンプレートデータ準備");

                    // 現在実行中のアセンブリ（DLLまたはEXE）に関する情報を取得
                    System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    
                    // HTMLテンプレートの準備
                    PrepareHtmlTemplates(assembly, rootPath, exportDir);

                    Application.DoEvents();

                    log.WriteLine("HTML保存");
                    Application.DoEvents();

                    ClearClipboardSafely();
                    //Clipboard.SetDataObject(new DataObject());
                    Application.DoEvents();

                    // ドキュメント全体を選択してクリップボードにコピー
                    application.Selection.WholeStory();
                    application.Selection.Copy();

                    Application.DoEvents();

                    // 選択範囲をドキュメントの先頭に折りたたむ
                    // （選択が解除され、カーソルがドキュメントの先頭に移動）
                    application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    // 一時ファイルを削除
                    if (File.Exists(tmpDocPath))
                    {
                        try { File.Delete(tmpDocPath); }
                        catch
                        {
                            load.Close();
                            load.Dispose();
                            MessageBox.Show(ErrMsgTmpDocOpen, ErrMsgFile, MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    Application.DoEvents();
                    Word.Document docCopy = application.Documents.Add();

                    Application.DoEvents();
                    docCopy.SaveAs2(tmpDocPath);
                    docCopy.TrackRevisions = false;
                    docCopy.AcceptAllRevisions();

                    docCopy.Select();
                    Application.DoEvents();

                    // クリップボードの内容を貼り付け、貼り付け先のスタイルを適用
                    application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);

                    Application.DoEvents();
                    ClearClipboardSafely();

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

                    // ドキュメント内のセクション数をログに記録
                    log.WriteLine("Number of sections: " + docCopy.Sections.Count);

                    // ドキュメントの最後のセクションのインデックスを取得
                    int lastSectionIdx = docCopy.Sections.Count;

                    // 表紙に関連する段落を収集
                    CollectCoverParagraphs(
                        docCopy,
                        ref manualTitle,
                        ref manualSubTitle,
                        ref manualVersion,
                        ref manualTitleCenter,
                        ref manualSubTitleCenter,
                        ref manualVersionCenter,
                        ref coverExist);

                    // 商標情報と著作権情報を収集
                    CollectTrademarkAndCopyrightDetails(
                        docCopy,
                        lastSectionIdx,
                        log,
                        ref trademarkTitle,
                        ref trademarkTextList,
                        ref trademarkRight
                        );

                    // 不要なタグや制御文字を削除
                    CleanUpManualTitles(
                        ref manualTitle,
                        ref manualSubTitle,
                        ref manualVersion,
                        ref manualTitleCenter,
                        ref manualSubTitleCenter,
                        ref manualVersionCenter);

                    List<List<string>> productSubLogoGroups = new List<List<string>>();

                    // 特定の条件に基づいて画像を抽出・変換・保存する
                    if (coverExist)
                    {
                        ProcessCoverImages(
                                docCopy,
                                application,
                                rootPath,
                                exportDir,
                                ref subTitle,
                                ref biCount,
                                ref productSubLogoGroups,
                                isPattern1,
                                isPattern2,
                                log
                            );
                    }

                    // ドキュメントの末尾にカーソルを移動
                    application.Selection.EndKey(Word.WdUnits.wdStory);
                    object selectionRange = application.Selection.Range;

                    // ドキュメントにキャンバスを追加（位置: 0, 0、サイズ: 1x1ポイント）
                    Word.Shape temporaryCanvas = docCopy.Shapes.AddCanvas(0, 0, 1, 1, ref selectionRange);

                    // キャンバスの折り返し設定を「行内」に変更
                    temporaryCanvas.WrapFormat.Type = Word.WdWrapType.wdWrapInline;

                    // キャンバス内の図形のプロパティを調整
                    AdjustCanvasShapes(docCopy);

                    temporaryCanvas.Delete();

                    // 表の幅がポイント単位で指定されている場合、自動調整を有効
                    foreach (Word.Table wt in docCopy.Tables)
                    {
                        if (wt.PreferredWidthType == Word.WdPreferredWidthType.wdPreferredWidthPoints)
                            wt.AllowAutoFit = true;
                    }

                    // スタイル名が「奥付タイトル」の場合、スタイル名を「titledef」に変更
                    foreach (Word.Style ws in docCopy.Styles)
                        if (ws.NameLocal == "奥付タイトル")
                            ws.NameLocal = "titledef";

                    // ドキュメントのWebオプションをUTF-8エンコーディングに設定
                    docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

                    // ドキュメントをHTML形式（フィルタ済み）で保存
                    // ドキュメント内の画像を tmp.files フォルダに自動的に出力
                    docCopy.SaveAs2(tmpHtmlPath, Word.WdSaveFormat.wdFormatFilteredHTML);
                    docCopy.Close();

                    File.Delete(tmpDocPath);

                    log.WriteLine("画像フォルダ コピー");

                    bool isTmpDot = true;

                    CopyAndDeleteTemporaryImages(tmpFolderForImagesSavedBySaveAs2Method, rootPath, exportDir, log);

                    string htmlStr;

                    using (StreamReader sr = new StreamReader(tmpHtmlPath, Encoding.UTF8))
                    {
                        // HTMLファイルの内容を文字列として読み取る
                        htmlStr = sr.ReadToEnd();
                    }

                    // HTML文字列を処理
                    htmlStr = ProcessHtmlString(htmlStr, isTmpDot);

                    // HTML文字列をXML形式に変換してロード
                    XmlDocument objXml = new XmlDocument();
                    objXml.LoadXml(htmlStr);

                    // 目次（TOC）と本文（Body）のXMLドキュメントを生成
                    ProcessXmlDocuments(objXml, docTitle, out XmlDocument objToc, out XmlDocument objBody);

                    // 現在の目次ノードと本文ノードを取得
                    XmlNode objTocCurrent = objToc.DocumentElement;
                    XmlNode objBodyCurrent = objBody.DocumentElement;

                    // CSSスタイル情報を取得
                    string className = "";
                    className = objXml.SelectSingleNode("/html/head/style[contains(comment(), 'mso-style-name')]").OuterXml;

                    // CSSスタイル文字列を整形（不要な空白や改行を削除）
                    className = Regex.Replace(className, "[\r\n\t ]+", "");
                    className = Regex.Replace(className, "}", "}\n");

                    // スタイル名を格納する辞書
                    Dictionary<string, string> styleName = new Dictionary<string, string>();

                    // 章分割に使用するCSSクラスを初期化
                    string chapterSplitClass = "";

                    // CSSスタイルのような文字列を解析し、特定の条件に一致するスタイルを抽出・加工
                    ProcessStyles(className, ref chapterSplitClass, styleName);

                    log.WriteLine("index.html出力");

                    // タイトル定義を格納するリスト
                    List<string> titleDeffenition = new List<string>();

                    // XMLドキュメント内の<p>タグでクラス名が'titledef'の要素を検索し、
                    // そのテキスト内容をトリムしてリストに追加
                    foreach (XmlElement link in objXml.SelectNodes("//p[@class='titledef']"))
                    {
                        titleDeffenition.Add(link.InnerText.Trim());
                    }

                    // インデックスHTMLテンプレートを生成
                    string idxHtmlTemplate = BuildIdxHtmlTemplate(docTitle, docid, mergeScript);
                    sw = new StreamWriter(indexHtmlPath, false, Encoding.UTF8);
                    sw.Write(idxHtmlTemplate);
                    sw.Close();

                    // 表紙テンプレート1と表紙テンプレート2を生成
                    string htmlCoverTemplate1 = BuildHtmlCoverTemplate1(isEdgeTracker);
                    string htmlCoverTemplate2 = "";

                    if (isEdgeTracker)
                    {
                        BuildEdgeTrackerCoverTemplate(
                            assembly,
                            rootPath,
                            exportDir,
                            manualTitle,
                            trademarkTitle,
                            trademarkTextList,
                            trademarkRight,
                            ref htmlCoverTemplate1);
                    }
                    else if (isEasyCloud)
                    {
                        BuildEasyCloudCoverTemplate(
                            rootPath,
                            exportDir,
                            manualTitle,
                            manualSubTitle,
                            manualVersion,
                            trademarkTitle,
                            trademarkTextList,
                            trademarkRight,
                            subTitle,
                            ref htmlCoverTemplate1,
                            ref htmlCoverTemplate2);
                    }
                    else if (isPattern1)
                    {
                        BuildPattern1CoverTemplate(
                            manualTitle,
                            manualTitleCenter,
                            manualSubTitle,
                            manualSubTitleCenter,
                            trademarkTitle,
                            trademarkTextList,
                            trademarkRight,
                            ref htmlCoverTemplate2);
                    }
                    else if (isPattern2)
                    {
                        BuildPattern2CoverTemplate(
                            productSubLogoGroups,
                            manualTitleCenter,
                            manualTitle,
                            manualSubTitleCenter,
                            manualSubTitle,
                            manualVersionCenter,
                            manualVersion,
                            trademarkTitle,
                            trademarkTextList,
                            trademarkRight,
                            ref htmlCoverTemplate2);
                    }

                    // すべてのパターンに共通するHTMLテンプレートの追加
                    AppendHtmlCoverTemplate2(ref htmlCoverTemplate2);

                    // HTMLテンプレート1を生成とHTMLテンプレート2を生成
                    string htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);
                    string htmlTemplate2 = "";
                    htmlTemplate2 += @"</body>" + "\n";
                    htmlTemplate2 += @"</html>" + "\n";

                    // 検索用JavaScriptコードを生成
                    string searchJs = BuildSearchJs();
                    
                    log.WriteLine("変換ループ開始");

                    // 目次（TOC）と本文（Body）を生成
                    GenerateTocAndBody(
                        objXml,
                        objBody,
                        objToc,
                        chapterSplitClass,
                        styleName,
                        docid,
                        bookInfoDef,
                        ref objBodyCurrent,
                        ref objTocCurrent,
                        load);

                    // 本文の最初のノードにIDが設定されていない場合、デフォルトのIDを設定
                    if (((XmlElement)objBody.DocumentElement.FirstChild).GetAttribute("id") == "")
                    {
                        ((XmlElement)objBody.DocumentElement.FirstChild).SetAttribute("id", docid + "00000");
                    }

                    // 目次ファイルを作成
                    GenerateTocFiles(objToc, rootPath, exportDir, mergeScript);

                    objXml = null;
                    File.Delete(tmpHtmlPath);

                    // CleanUpXmlNodes メソッドを呼び出す
                    CleanUpXmlNodes(objBody);

                    // 検索用データを生成
                    GenerateSearchFiles(
                        objBody,                // HTMLの内容を保持するXmlDocument
                        rootPath,               // ドキュメントのルートパス
                        exportDir,              // 出力ディレクトリ名
                        docid,                  // ドキュメントID
                        htmlTemplate1,          // HTMLテンプレート1
                        htmlTemplate2,          // HTMLテンプレート2
                        htmlCoverTemplate1,     // HTMLカバーテンプレート1
                        htmlCoverTemplate2,     // HTMLカバーテンプレート2
                        objToc,                 // 目次情報を保持するXmlDocument
                        mergeScript,            // マージスクリプトの辞書
                        searchJs                // 検索用JavaScriptコード
                    );

                    log.WriteLine("Zipファイル作成");

                    // Zipファイルを作成
                    GenerateZipArchive(zipDirPath, rootPath, exportDir, headerDir, docFullName, docName, log);
                }

                catch (Exception ex)
                {
                    load.Close();
                    load.Dispose();

                    // スタックトレースを取得（例外発生箇所の詳細情報を含む）
                    StackTrace stackTrace = new StackTrace(ex, true);

                    // 例外の詳細情報をログに記録
                    log.WriteLine("[Error] Exception Details:");
                    log.WriteLine($"  Source: {ex.Source ?? "Unknown Source"}");
                    log.WriteLine($"  TargetSite: {ex.TargetSite}");
                    log.WriteLine($"  Message: {ex.Message}");
                    log.WriteLine($"  StackTrace: {stackTrace}");

                    MessageBox.Show(ErrMsg);

                    button3.Enabled = true;
                    return;
                }
            }

            // log.txtを削除したい場合は以下のコードを有効にする
            //File.Delete(logPath);

            application.ActiveWindow.View.Type = defaultView;
            load.Close();
            load.Dispose();

            // ユーザーに出力したHTMLをブラウザで表示するか確認するメッセージボックスを表示
            DialogResult selectMsg = MessageBox.Show(exportDirPath + MsgHtmlOutputSuccess1, MsgHtmlOutputSuccess2, MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (selectMsg == DialogResult.Yes)
            {
                try
                {
                    // index.htmlを既定のブラウザで開く
                    Process.Start(indexHtmlPath);
                }
                catch
                {
                    // index.htmlの起動に失敗した場合、エラーメッセージを表示
                    MessageBox.Show(ErrMsgHtmlOutputFailure1, ErrMsgHtmlOutputFailure2, MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }

            // イベントハンドラを再登録
            application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }
    }
}
