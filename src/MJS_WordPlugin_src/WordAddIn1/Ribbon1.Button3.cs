using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Drawing;
using System.Xml;
using System.Threading;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using System.Reflection.Emit;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void button3_Click_1(object sender, RibbonControlEventArgs e)
        {
            StreamWriter sw;
            blHTMLPublish = true;

            loader load = new loader();
            load.Show();

            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            button3.Enabled = false;

            application.DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);

            var defaultView = application.ActiveWindow.View.Type;

            if (!Regex.IsMatch(activeDocument.Name, FileNamePattern))
            {
                load.Close();
                load.Dispose();
                MessageBox.Show(InvalidFileNameMessage, ErrFileNameRule, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
            string webHelpFolderName = null;

            if (properties.Cast<Microsoft.Office.Core.DocumentProperty>().Any(x => x.Name == "webHelpFolderName"))
            {
                webHelpFolderName = properties["webHelpFolderName"].Value;
            }

            load.Visible = false;

            if (!makeBookInfo(load))
            {
                load.Close();
                load.Dispose();
                return;
            }

            Dictionary<string, string> mergeScript = new Dictionary<string, string>();

            CollectMergeScript(activeDocument.Path, activeDocument.Name, mergeScript);

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

            if (webHelpFolderName != null && webHelpFolderName.Length > 0)
            {
                exportDir = webHelpFolderName;
            }

            using (StreamWriter log = new StreamWriter(rootPath + "\\log.txt", false, Encoding.UTF8))
            {
                try
                {
                    log.WriteLine("テンプレートデータ準備");

                    System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                    // HTMLテンプレートの準備
                    PrepareHtmlTemplates(assembly, rootPath, exportDir);

                    string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");
                    string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");
                    string zipDirPath = rootPath + "\\" + docid + "_" + exportDir + "_" + DateTime.Today.ToString("yyyyMMdd");

                    Application.DoEvents();

                    log.WriteLine("HTML保存");
                    Application.DoEvents();

                    Clipboard.Clear();
                    Clipboard.SetDataObject(new DataObject());
                    Application.DoEvents();
                    application.Selection.WholeStory();
                    application.Selection.Copy();
                    Application.DoEvents();
                    application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

                    if (File.Exists(rootPath + "\\tmp.doc"))
                    {
                        try { File.Delete(rootPath + "\\tmp.doc"); }
                        catch
                        {
                            load.Close();
                            load.Dispose();
                            MessageBox.Show("同階層のtmp.docが開かれています。\r\ntmp.docを閉じてから実行してください。", "ファイルエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                            return;
                        }
                    }

                    Application.DoEvents();
                    Word.Document docCopy = application.Documents.Add();

                    Application.DoEvents();
                    docCopy.SaveAs2(rootPath + "\\tmp.doc");
                    docCopy.TrackRevisions = false;
                    docCopy.AcceptAllRevisions();

                    docCopy.Select();
                    Application.DoEvents();
                    application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);

                    load.Invoke((MethodInvoker)delegate
                    {
                        Clipboard.Clear();
                    });

                    //Clipboard.Clear();
                    // クリップボードをクリアする代わりに、空のデータを設定して内容を上書きする
                    //Clipboard.SetDataObject(new DataObject());

                    Application.DoEvents();

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

                    log.WriteLine("Number of sections: " + docCopy.Sections.Count);
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

                    bool isTradeMarksDetected = false;
                    bool isRightDetected = false;

                    // 商標情報と著作権情報を収集
                    CollectTrademarkAndCopyrightDetails(
                        docCopy,
                        lastSectionIdx,
                        log,
                        ref trademarkTitle,
                        ref trademarkTextList,
                        ref trademarkRight,
                        ref isTradeMarksDetected,
                        ref isRightDetected);

                    CleanUpManualTitles(
                        ref manualTitle,
                        ref manualSubTitle,
                        ref manualVersion,
                        ref manualTitleCenter,
                        ref manualSubTitleCenter,
                        ref manualVersionCenter);

                    List<List<string>> productSubLogoGroups = new List<List<string>>();

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

                    application.Selection.EndKey(Word.WdUnits.wdStory);
                    object selectionRange = application.Selection.Range;
                    Word.Shape wst = docCopy.Shapes.AddCanvas(0, 0, 1, 1, ref selectionRange);
                    wst.WrapFormat.Type = Word.WdWrapType.wdWrapInline;

                    // キャンバスに関連する図形のプロパティを調整
                    AdjustCanvasShapes(docCopy);

                    wst.Delete();

                    foreach (Word.Table wt in docCopy.Tables)
                    {
                        if (wt.PreferredWidthType == Word.WdPreferredWidthType.wdPreferredWidthPoints)
                            wt.AllowAutoFit = true;
                    }
                    
                    foreach (Word.Style ws in docCopy.Styles)
                        if (ws.NameLocal == "奥付タイトル")
                            ws.NameLocal = "titledef";

                    docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                    docCopy.SaveAs2(rootPath + "\\tmp.html", Word.WdSaveFormat.wdFormatFilteredHTML);
                    docCopy.Close();
                    File.Delete(rootPath + "\\tmp.doc");

                    log.WriteLine("画像フォルダ コピー");

                    string tmpFolderForImagesSavedBySaveAs2Method = rootPath + "\\tmp.files";
                    bool isTmpDot = true;

                    if (!Directory.Exists(tmpFolderForImagesSavedBySaveAs2Method))
                    {
                        isTmpDot = false;
                        tmpFolderForImagesSavedBySaveAs2Method = rootPath + "\\tmp_files";
                    }

                    if (Directory.Exists(tmpFolderForImagesSavedBySaveAs2Method))
                    {
                        foreach (string pict in Directory.GetFiles(tmpFolderForImagesSavedBySaveAs2Method))
                        {
                            File.Copy(pict, rootPath + "\\" + exportDir + "\\pict\\" + Path.GetFileName(pict));
                        }

                        Directory.Delete(tmpFolderForImagesSavedBySaveAs2Method, true);
                    }

                    StreamReader sr = new StreamReader(rootPath + "\\tmp.html", Encoding.UTF8);
                    string htmlStr = sr.ReadToEnd();
                    sr.Close();

                    htmlStr = ProcessHtmlString(htmlStr, isTmpDot);

                    XmlDocument objXml = new XmlDocument();

                    objXml.LoadXml(htmlStr);

                    ProcessXmlDocuments(objXml, docTitle, out XmlDocument objToc, out XmlDocument objBody);

                    XmlNode objTocCurrent = objToc.DocumentElement;
                    XmlNode objBodyCurrent = objBody.DocumentElement;

                    string className = "";
                    className = objXml.SelectSingleNode("/html/head/style[contains(comment(), 'mso-style-name')]").OuterXml;
                    className = Regex.Replace(className, "[\r\n\t ]+", "");
                    className = Regex.Replace(className, "}", "}\n");

                    Dictionary<string, string> styleName = new Dictionary<string, string>();

                    string chapterSplitClass = "";

                    // CSSスタイルのような文字列を解析し、特定の条件に一致するスタイルを抽出・加工
                    ProcessStyles(className, ref chapterSplitClass, styleName);

                    log.WriteLine("index.html出力");

                    List<string> titleDeffenition = new List<string>();

                    foreach (XmlElement link in objXml.SelectNodes("//p[@class='titledef']"))
                    {
                        titleDeffenition.Add(link.InnerText.Trim());
                    }

                    string idxHtmlTemplate = BuildIdxHtmlTemplate(docTitle, docid, mergeScript);

                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\index.html", false, Encoding.UTF8);
                    sw.Write(idxHtmlTemplate);
                    sw.Close();

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
                    
                    string htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);
                    string htmlTemplate2 = "";
                    htmlTemplate2 += @"</body>" + "\n";
                    htmlTemplate2 += @"</html>" + "\n";

                    string searchJs = BuildSearchJs();
                    
                    log.WriteLine("変換ループ開始");

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

                    if (((XmlElement)objBody.DocumentElement.FirstChild).GetAttribute("id") == "")
                    {
                        ((XmlElement)objBody.DocumentElement.FirstChild).SetAttribute("id", docid + "00000");
                    }

                    // 目次ファイルを作成
                    GenerateTocFiles(objToc, rootPath, exportDir, mergeScript);

                    //objXml.Save(rootPath + "\\base.xml");
                    objXml = null;
                    File.Delete(rootPath + "\\tmp.html");

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
                    GenerateZipArchive(zipDirPath, rootPath, exportDir, headerDir, docFullName, docName, log);

                }

                catch (Exception ex)
                {
                    load.Close();
                    load.Dispose();

                    StackTrace stackTrace = new StackTrace(ex, true);

                    log.WriteLine(ex.Message);
                    log.WriteLine(ex.HelpLink);
                    log.WriteLine(ex.Source);
                    log.WriteLine(ex.StackTrace);
                    log.WriteLine(ex.TargetSite);
                    MessageBox.Show("エラーが発生しました");

                    button3.Enabled = true;
                    return;
                }
            }

            // log.txtを削除したい場合は以下のコードを有効にする
            //File.Delete(rootPath + "\\log.txt");

            application.ActiveWindow.View.Type = defaultView;
            load.Close();
            load.Dispose();

            DialogResult selectMess = MessageBox.Show(rootPath + "\\" + exportDir + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？", "HTML出力成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            
            if (selectMess == DialogResult.Yes)
            {
                try
                {
                    Process.Start(rootPath + "\\" + exportDir + @"\index.html");
                }
                catch
                {
                    MessageBox.Show("HTMLの出力に失敗しました。", "HTML出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
           
            application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }
    }
}
