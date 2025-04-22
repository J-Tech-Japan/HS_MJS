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
            blHTMLPublish = true;

            loader load = new loader();
            load.Show();

            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            StreamWriter sw;
            application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            button3.Enabled = false;
            application.DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);

            var defaultView = application.ActiveWindow.View.Type;

            if (!Regex.IsMatch(activeDocument.Name, @"^[A-Z]{3}(_[^_]*?){2}\.docx*$"))
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」", "ファイル命名規則エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // read word properties
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)activeDocument.CustomDocumentProperties;
            string webHelpFolderName = null;

            // check webHelpFolderName property exists
            if (properties.Cast<Microsoft.Office.Core.DocumentProperty>().Any(x => x.Name == "webHelpFolderName"))
            {
                webHelpFolderName = properties["webHelpFolderName"].Value;
            }

            // SOURCELINK追加==========================================================================START
            load.Visible = false;
            if (!makeBookInfo(load))
            {
                load.Close();
                load.Dispose();
                return;
            }

            // Collect merge
            Dictionary<string, string> mergeScript = new Dictionary<string, string>();
            using (StreamReader sr = new StreamReader(
                    activeDocument.Path + "\\headerFile\\" + Regex.Replace(activeDocument.Name, "^(.{3}).+$", "$1") + @".txt", System.Text.Encoding.Default))
            {
                // 書誌情報番号の最大値取得
                while (sr.Peek() >= 0)
                {
                    string strBuffer = sr.ReadLine();

                    // SOURCELINK追加==========================================================================START
                    string[] info = strBuffer.Split('\t');

                    if (info.Length == 4)
                    {
                        if (!info[3].Equals(""))
                        {
                            // this page will in that page
                            info[3] = info[3].Replace("(", "").Replace(")", "");
                            if (!mergeScript.Any(x => x.Key == info[2] && x.Value == info[3]))
                            {
                                mergeScript.Add(info[2], info[3]);
                            }
                        }
                    }
                }
            }

            bool isEasyCloud, isEdgeTracker, isPattern1, isPattern2;

            if (!HandleCoverSelection(load, out isEasyCloud, out isEdgeTracker, out isPattern1, out isPattern2))
            {
                return;
            }

            load.Visible = true;
            // SOURCELINK追加==========================================================================END

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
                    using (Stream stream = assembly.GetManifestResourceStream("WordAddIn1.htmlTemplates.zip"))
                    {
                        FileStream fs = File.Create(rootPath + "\\htmlTemplates.zip");
                        stream.Seek(0, SeekOrigin.Begin);
                        stream.CopyTo(fs);
                        fs.Close();
                    }

                    if (Directory.Exists(rootPath + "\\htmlTemplates"))
                    {
                        Directory.Delete(rootPath + "\\htmlTemplates", true);
                    }

                    System.IO.Compression.ZipFile.ExtractToDirectory(rootPath + "\\htmlTemplates.zip", rootPath);

                    if (Directory.Exists(rootPath + "\\" + exportDir))
                    {
                        Directory.Delete(rootPath + "\\" + exportDir, true);
                    }
                    if (Directory.Exists(rootPath + "\\tmpcoverpic")) Directory.Delete(rootPath + "\\tmpcoverpic", true);
                    Directory.Move(rootPath + "\\htmlTemplates", rootPath + "\\" + exportDir);

                    File.Delete(rootPath + "\\htmlTemplates.zip");

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
                    Clipboard.Clear();
                    Application.DoEvents();
                    //docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                    //docCopy.SaveAs2(rootPath + "\\tmp.doc");
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
                    //string strOutFileName = "";

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

                    
                    string bell = new string((char)7, 1);
                    manualTitle = Regex.Replace(manualTitle, @"<br/>$", "").Replace(bell, "").Trim();
                    manualSubTitle = Regex.Replace(manualSubTitle, @"<br/>$", "").Replace(bell, "").Trim();
                    manualVersion = Regex.Replace(manualVersion, @"<br/>$", "").Replace(bell, "").Trim();
                    manualTitleCenter = Regex.Replace(manualTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
                    manualSubTitleCenter = Regex.Replace(manualSubTitleCenter, @"<br/>$", "").Replace(bell, "").Trim();
                    manualVersionCenter = Regex.Replace(manualVersionCenter, @"<br/>$", "").Replace(bell, "").Trim();
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

                    //docCopy.SaveAs2 method save images files into tmp.files folder, but sometimes it's tmp_files folder (wtf?!?), so need to check
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

                    System.Xml.XmlDocument objXml = new System.Xml.XmlDocument();

                    objXml.LoadXml(htmlStr);

                    // 新しいメソッドを呼び出す
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
                    foreach (System.Xml.XmlElement link in objXml.SelectNodes("//p[@class='titledef']"))
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

                    htmlCoverTemplate2 += @"<script type=""text/javascript"" language=""javascript1.2"">//<![CDATA[" + "\n";
                    htmlCoverTemplate2 += @"<!--" + "\n";
                    htmlCoverTemplate2 += @"if (window.writeIntopicBar)" + "\n";
                    htmlCoverTemplate2 += @"   writeIntopicBar(0);" + "\n";
                    htmlCoverTemplate2 += @"//-->" + "\n";
                    htmlCoverTemplate2 += @"//]]></script>" + "\n";
                    htmlCoverTemplate2 += @"</body>" + "\n";
                    htmlCoverTemplate2 += @"</html>" + "\n";

                    string htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);

                    string htmlTemplate2 = "";
                    htmlTemplate2 += @"</body>" + "\n";
                    htmlTemplate2 += @"</html>" + "\n";

                    string searchJs = BuildSearchJs();
                    
                    string htmlToc = "";
                    string htmlToc1 = "";
                    string htmlToc2 = "";
                    string htmlToc3 = "";

                    string lv1styleName = "";
                    string lv2styleName = "";
                    string lv3styleName = "";

                    int lv1count = 0;
                    int lv2count = 0;
                    int lv3count = 0;

                    bool chapterSplit = false;

                    log.WriteLine("変換ループ開始");
                    //return;

                    foreach (System.Xml.XmlElement sectionNode in objXml.SelectNodes("/html/body/div"))
                    {
                        objBodyCurrent = (System.Xml.XmlElement)objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));

                        if (chapterSplit)
                        {
                            chapterSplit = false;
                        }

                        if (sectionNode.SelectSingleNode(chapterSplitClass) != null)
                        {
                            ((System.Xml.XmlElement)objBodyCurrent).SetAttribute("style", "width:714px");
                            lv1styleName = chapterSplitClass;
                            chapterSplit = true;
                        }

                        bool breakFlg = false;

                        foreach (System.Xml.XmlNode childs in sectionNode.SelectNodes("*"))
                        {
                            string thisStyleName = "";

                            if (childs.SelectSingleNode("@class") == null)
                            {
                                if (styleName.ContainsKey(childs.Name))
                                {
                                    thisStyleName = styleName[childs.Name];
                                }
                            }
                            else
                            {
                                if (styleName.ContainsKey(childs.Name + "." + childs.SelectSingleNode("@class").InnerText))
                                {
                                    thisStyleName = styleName[childs.Name + "." + childs.SelectSingleNode("@class").InnerText];
                                }
                            }

                            if ((thisStyleName == "") && (childs.SelectSingleNode("*[@class != '']") != null))
                            {
                                if (styleName.ContainsKey(childs.SelectSingleNode("*[@class != '']").Name + "." + childs.SelectSingleNode("*[@class != '']/@class").InnerText))
                                {
                                    thisStyleName = styleName[childs.SelectSingleNode("*[@class != '']").Name + "." + childs.SelectSingleNode("*[@class != '']/@class").InnerText];
                                }
                            }
                            else if ((thisStyleName == "") && (childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']") != null))
                            {
                                if (styleName.ContainsKey(childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name))
                                {
                                    thisStyleName = styleName[childs.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name];
                                }
                            }

                            if (childs.SelectSingleNode(".//text()[1]") != null)
                            {
                                if (Regex.IsMatch(childs.SelectSingleNode(".//text()[1]").InnerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                                {
                                    breakFlg = true;
                                    break;
                                }

                                if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[\d０-９](?:[^・用]+|)$") && Regex.IsMatch(childs.SelectSingleNode(".//text()[1]").InnerText, @"^(?:\d+\.)*\d+[\s　]+"))
                                {
                                    childs.SelectSingleNode(".//text()[1]").InnerText = Regex.Replace(childs.SelectSingleNode(".//text()[1]").InnerText, @"^(?:\d+\.)*\d+[\s　]+", "");
                                }
                            }

                            string setid = "";
                            if (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1２2３3](?:[^・用]+|)$"))
                            {
                                if (childs.SelectSingleNode(".//a[starts-with(@name, '" + docid + bookInfoDef + "')]") != null)
                                {
                                    //aaa
                                    setid = ((System.Xml.XmlElement)childs.SelectSingleNode(".//a[starts-with(@name, '" + docid + bookInfoDef + "')]")).GetAttribute("name");
                                }
                                else
                                {
                                    load.Visible = false;
                                    MessageBox.Show(childs.InnerText + ":書誌情報ブックマークの設定が行われていません。");
                                    load.Visible = true;
                                }
                            }


                            if (Regex.IsMatch(thisStyleName, "目[　 ]*次"))
                            {
                            }
                            else if (Regex.IsMatch(thisStyleName, "章[　 ]*扉.*タイトル"))
                            {
                                lv1count++;
                                lv2styleName = "";
                                lv2count = 0;
                                lv3styleName = "";
                                lv2count = 0;

                                objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");
                                objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", Regex.Replace(childs.InnerText, @"^第[\d０-９]+章[　\s]*", ""));

                                ((System.Xml.XmlElement)objBodyCurrent).SetAttribute("id", setid);
                            }
                            else if (Regex.IsMatch(thisStyleName, "章[　 ]*扉"))
                            {
                            }
                            else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$"))
                            {
                                if (!Regex.IsMatch(childs.InnerText, @"目\s*次\s*$"))
                                {
                                    if ((lv1styleName == "") || (lv1styleName == thisStyleName) || Regex.IsMatch(lv1styleName, @"(見出し|Heading)\s*[２2]"))
                                    {
                                        lv1count++;
                                        lv2styleName = "";
                                        lv2count = 0;
                                        lv3styleName = "";
                                        lv3count = 0;

                                        objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");

                                        objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                        ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                        ((System.Xml.XmlElement)objTocCurrent).SetAttribute("href", setid);

                                        lv1styleName = thisStyleName;
                                    }
                                    else
                                    {
                                        lv2count++;
                                        lv3styleName = "";
                                        lv3count = 0;

                                        if ((objTocCurrent == null) || (objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]") == null))
                                        {
                                        }
                                        else
                                        {
                                            objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]");

                                            objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                            ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                            ((System.Xml.XmlElement)objTocCurrent).SetAttribute("href", setid);
                                        }
                                        lv2styleName = thisStyleName;

                                    }
                                    objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                                    ((System.Xml.XmlElement)objBodyCurrent).SetAttribute("id", setid);

                                    ((System.Xml.XmlElement)objBodyCurrent).AppendChild(objBody.CreateElement("p"));
                                    ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading1");


                                    foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                    {
                                        innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                    }
                                }
                            }
                            else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[２2](?![・用])"))
                            {
                                if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                                {
                                    lv1count++;
                                    lv2styleName = "";
                                    lv2count = 0;
                                    lv3styleName = "";
                                    lv3count = 0;

                                    objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");
                                    objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("href", setid);
                                }
                                else
                                {
                                    if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                                    {
                                        lv2count++;
                                        lv3styleName = "";
                                        lv3count = 0;

                                        objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]");
                                    }
                                    else
                                    {
                                        lv3count++;

                                        objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 2]");
                                    }

                                    objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("href", setid);
                                }

                                objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                                ((System.Xml.XmlElement)objBodyCurrent).SetAttribute("id", setid);

                                objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading1 NoPageBreak");

                                foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                {
                                    innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                }

                                if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                                {
                                    lv1styleName = thisStyleName;
                                }
                                else if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                                {
                                    lv2styleName = thisStyleName;
                                }
                                else
                                {
                                    lv3styleName = thisStyleName;
                                }
                            }
                            else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                            {
                                //if ((lv1styleName == "") || (lv1styleName == thisStyleName) ||
                                //   (lv2styleName == "") || (lv2styleName == thisStyleName) ||
                                //   (lv3styleName == "") || (lv3styleName == thisStyleName))
                                //{
                                //    if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                                //    {
                                //        lv1count++;
                                //        lv2styleName = "";
                                //        lv2count = 0;
                                //        lv3styleName = "";
                                //        lv3count = 0;

                                //        objTocCurrent = objTocCurrent.SelectSingleNode("/result/item");
                                //    }
                                //    else if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                                //    {
                                //        lv2count++;
                                //        lv3styleName = "";
                                //        lv3count = 0;

                                //        objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 1]");
                                //    }
                                //    else if ((lv3styleName == "") || (lv3styleName == thisStyleName))
                                //    {
                                //        lv3count++;

                                //        objTocCurrent = objTocCurrent.SelectSingleNode("ancestor-or-self::item[count(ancestor::item) = 2]");

                                //    }

                                //    objTocCurrent = objTocCurrent.AppendChild(objToc.CreateElement("item"));
                                //    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("title", childs.InnerText);
                                //    ((System.Xml.XmlElement)objTocCurrent).SetAttribute("href", setid);

                                //    objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                                //    ((System.Xml.XmlElement)objBodyCurrent).SetAttribute("id", setid);

                                //    objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                //    ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading1");

                                //    foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                //    {
                                //        innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                //    }

                                //    if ((lv1styleName == "") || (lv1styleName == thisStyleName))
                                //    {
                                //        lv1styleName = thisStyleName;
                                //    }
                                //    else if ((lv2styleName == "") || (lv2styleName == thisStyleName))
                                //    {
                                //        lv2styleName = thisStyleName;
                                //    }
                                //    else
                                //    {
                                //        lv3styleName = thisStyleName;
                                //    }
                                //}
                                
                                objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading3");
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("id", Regex.Replace(setid, "^.*?♯(.*?)$", "$1"));

                                foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                {
                                    innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                }
                            }
                            else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[４4](?![・用])"))
                            {
                                objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading4");
                                foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                {
                                    innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                }
                            }
                            else if (Regex.IsMatch(thisStyleName, @"(見出し|Heading)\s*[５5]"))
                            {
                                objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading5");
                                foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                {
                                    innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                }
                            }
                            else
                            {
                                if (objBodyCurrent.Name == "result")
                                {
                                    objBodyCurrent = objBody.DocumentElement.AppendChild(objBody.CreateElement("div"));
                                }
                                innerNode(styleName, objBodyCurrent, childs);
                            }
                        }

                        if (breakFlg) break;
                    }

                    if (((System.Xml.XmlElement)objBody.DocumentElement.FirstChild).GetAttribute("id") == "")
                    {
                        ((System.Xml.XmlElement)objBody.DocumentElement.FirstChild).SetAttribute("id", docid + "00000");
                    }

                    //目次出力
                    foreach (System.Xml.XmlNode toc in objToc.SelectNodes("/result/item"))
                    {
                        htmlToc = @"{""type"":""book"",""name"":""" + ((System.Xml.XmlElement)toc).GetAttribute("title") + @""",""key"":""toc1""}";

                        foreach (System.Xml.XmlNode toc1 in toc.SelectNodes("item"))
                        {
                            if (htmlToc1 != "")
                            {
                                htmlToc1 = htmlToc1 + ",";
                            }

                            htmlToc1 = htmlToc1 + @"{""type"":""";

                            if (toc1.SelectNodes("item").Count != 0)
                            {
                                htmlToc1 = htmlToc1 + "book";
                            }
                            else
                            {
                                htmlToc1 = htmlToc1 + "item";
                            }

                            htmlToc1 += @""",""name"":""" + ((System.Xml.XmlElement)toc1).GetAttribute("title") + @"""";

                            if (toc1.SelectNodes("item").Count != 0)
                            {
                                htmlToc1 += @",""key"":""toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + @"""";
                            }

                            if (((System.Xml.XmlElement)toc1).GetAttribute("href") != "")
                            {
                                htmlToc1 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((System.Xml.XmlElement)toc1).GetAttribute("href")) + @"""";
                            }

                            htmlToc1 += "}";

                            foreach (System.Xml.XmlNode toc2 in toc1.SelectNodes("item"))
                            {
                                if (htmlToc2 != "")
                                {
                                    htmlToc2 = htmlToc2 + ",";
                                }

                                htmlToc2 += @"{""type"":""";

                                if (toc2.SelectNodes("item").Count != 0)
                                {
                                    htmlToc2 += "book";
                                }
                                else
                                {
                                    htmlToc2 += "item";
                                }

                                htmlToc2 += @""",""name"":""" + ((System.Xml.XmlElement)toc2).GetAttribute("title") + @"""";

                                if (toc2.SelectNodes("item").Count != 0)
                                {
                                    htmlToc2 += @",""key"":""toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + @"""";
                                }
                                if (((System.Xml.XmlElement)toc2).GetAttribute("href") != "")
                                {
                                    htmlToc2 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((System.Xml.XmlElement)toc2).GetAttribute("href")) + @"""";
                                }

                                htmlToc2 += "}";

                                foreach (System.Xml.XmlNode toc3 in toc2.SelectNodes("item"))
                                {
                                    if (htmlToc3 != "")
                                    {
                                        htmlToc3 += ",";
                                    }

                                    htmlToc3 += @"{""type"":""item"",""name"":""" + ((System.Xml.XmlElement)toc3).GetAttribute("title") + @""",""url"":""" + makeHrefWithMerge(mergeScript, ((System.Xml.XmlElement)toc3).GetAttribute("href")) + @"""}";
                                }

                                if (htmlToc3 != "")
                                {
                                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + ".new.js", false, Encoding.UTF8);
                                    sw.WriteLine("(function() {");
                                    sw.WriteLine("var toc =  [" + htmlToc3 + "];");
                                    sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                                    sw.WriteLine("})();");
                                    sw.Close();
                                    htmlToc3 = "";
                                }
                            }

                            if (htmlToc2 != "")
                            {
                                sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + ".new.js", false, Encoding.UTF8);
                                sw.WriteLine("(function() {");
                                sw.WriteLine("var toc =  [" + htmlToc2 + "];");
                                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                                sw.WriteLine("})();");
                                sw.Close();
                                htmlToc2 = "";
                            }
                        }

                        sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc1.new.js", false, Encoding.UTF8);
                        sw.WriteLine("(function() {");
                        sw.WriteLine("var toc =  [" + htmlToc1 + "];");
                        sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                        sw.WriteLine("})();");
                        sw.Close();

                    }

                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc.new.js", false, Encoding.UTF8);
                    sw.WriteLine("(function() {");
                    sw.WriteLine("var toc =  [" + htmlToc + "];");
                    sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                    sw.WriteLine("})();");
                    sw.Close();

                    //objXml.Save(rootPath + "\\base.xml");
                    objXml = null;
                    File.Delete(rootPath + "\\tmp.html");

                    //objBody.Save(rootPath + "\\body.xml");
                    //objToc.Save(rootPath + "\\toc.xml");

                    foreach (System.Xml.XmlElement langSpan in objBody.SelectNodes(".//span[boolean(@lang)]|.//a"))
                    {
                        langSpan.RemoveAttribute("lang");

                        if (langSpan.Name == "a")
                        {
                            langSpan.RemoveAttribute("name");
                        }

                        if (langSpan.Attributes.Count == 0)
                        {
                            while (langSpan.ChildNodes.Count != 0)
                            {
                                langSpan.ParentNode.InsertBefore(langSpan.ChildNodes[0], langSpan);
                            }
                            langSpan.ParentNode.RemoveChild(langSpan);
                        }
                    }

                    while (objBody.SelectSingleNode("/result/div//*[((name() = 'div') or (name() = 'br')) and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ") != null)
                    {
                        System.Xml.XmlNode lineBreak = objBody.SelectSingleNode("/result/div//*[((name() = 'div') or (name() = 'br')) and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ");
                        lineBreak.ParentNode.RemoveChild(lineBreak);
                    }

                    while (objBody.SelectSingleNode("/result/div//*[not(img)][(name() = 'p') and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ") != null)
                    {
                        System.Xml.XmlNode lineBreak = objBody.SelectSingleNode("/result/div//*[not(img)][(name() = 'p') and not(boolean(ancestor::table)) and not(boolean(node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != ''])) and not(boolean(following-sibling::node()[translate(., ' &#9;" + ((char)160).ToString() + "　', '') != '']))] ");
                        lineBreak.ParentNode.RemoveChild(lineBreak);
                    }

                    System.Xml.XmlDocument searchWords = new System.Xml.XmlDocument();
                    searchWords.LoadXml("<div class='search'></div>");

                    foreach (System.Xml.XmlNode splithtml in objBody.SelectNodes("/result/div"))
                    {
                        string thisId = ((System.Xml.XmlElement)splithtml).GetAttribute("id");
                        ((System.Xml.XmlElement)splithtml).RemoveAttribute("id");
                        ((System.Xml.XmlElement)splithtml).RemoveAttribute("style");

                        if (thisId == docid + "00000")
                        {
                            sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + thisId + ".html", false, Encoding.UTF8);
                            string coverBody = "";
                            foreach (System.Xml.XmlNode coverItem in splithtml.SelectNodes(".//*[starts-with(@class, 'manual_')]"))
                            {
                                coverBody += coverItem.OuterXml;
                            }

                            //sw.Write(htmlCoverTemplate1 + coverBody + htmlCoverTemplate2);
                            sw.Write(htmlCoverTemplate1 + htmlCoverTemplate2);
                            sw.Close();
                        }
                        else
                        {
                            string htmlTemplate1cpy = htmlTemplate1;
                            if (objToc.SelectSingleNode(".//item[@href = '" + thisId + "']") != null)
                            {
                                htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, "<title></title>", "<title>" + ((System.Xml.XmlElement)objToc.SelectSingleNode(".//item[@href = '" + thisId + "']")).GetAttribute("title") + "</title>");
                                string breadcrumb = "";
                                System.Xml.XmlElement breadcrumbDisplay = objBody.CreateElement("div");
                                breadcrumbDisplay.SetAttribute("style", "text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;");

                                string tocId = "";

                                foreach (System.Xml.XmlNode tocItem in objToc.SelectNodes(".//item[@href = '" + thisId + "']/ancestor-or-self::item"))
                                {
                                    if (breadcrumb != "")
                                    {
                                        breadcrumb += " > ";
                                        breadcrumbDisplay.AppendChild(objBody.CreateTextNode(" > "));
                                    }
                                    breadcrumb += ((System.Xml.XmlElement)tocItem).GetAttribute("title");

                                    if (tocItem.SelectSingleNode("@href") != null)
                                    {
                                        breadcrumbDisplay.AppendChild(objBody.CreateElement("a"));
                                        ((System.Xml.XmlElement)breadcrumbDisplay.LastChild).SetAttribute("href", "./" + makeHrefWithMerge(mergeScript, ((System.Xml.XmlElement)tocItem).GetAttribute("href")) + "");
                                        breadcrumbDisplay.LastChild.InnerText = ((System.Xml.XmlElement)tocItem).GetAttribute("title");
                                    }
                                    else
                                    {
                                        breadcrumbDisplay.AppendChild(objBody.CreateTextNode(((System.Xml.XmlElement)tocItem).GetAttribute("title")));
                                    }

                                    if (tocId != "")
                                    {
                                        tocId += ".";
                                    }
                                    int precedingItemCount = tocItem.SelectNodes("preceding-sibling::item[boolean(item)]|self::item[boolean(item)]").Count;
                                    tocId += precedingItemCount.ToString();
                                    if (tocItem.SelectSingleNode("item") == null)
                                    {
                                        tocId += "_";
                                        tocId += (tocItem.SelectNodes("preceding-sibling::item[not(boolean(item)) and (count(preceding-sibling::item[boolean(item)]) = " + precedingItemCount + ")]").Count + 1).ToString();
                                    }
                                }
                                htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, "♪", tocId);

                                string searchText = splithtml.InnerText.Replace("&", "&amp;").Replace("<", "&lt;");
                                string displayText = searchText;
                                if (searchText.Length >= 90)
                                {
                                    displayText = displayText.Substring(0, 90) + " ...";
                                }

                                string[] wide = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン" };
                                string[] narrow = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" };

                                for (int i = 0; i < wide.Length; i++)
                                {
                                    searchText = Regex.Replace(searchText, wide[i], narrow[i]);
                                }
                                searchText = searchText.ToLower();

                                searchWords.DocumentElement.AppendChild(searchWords.CreateElement("div"));
                                ((System.Xml.XmlElement)searchWords.DocumentElement.LastChild).SetAttribute("id", thisId);
                                searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>" + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>" + ((System.Xml.XmlElement)objToc.SelectSingleNode(".//item[@href = '" + thisId + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='displayText'>" + displayText + "</div><div class='search_word'>" + searchText + "</div>";

                                htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, @"<meta name=""topic-breadcrumbs"" content="""" />", @"<meta name=""topic-breadcrumbs"" content=""" + breadcrumb + @""" />");
                                splithtml.InsertBefore(breadcrumbDisplay, splithtml.FirstChild);
                            }

                            if (!String.IsNullOrEmpty(thisId))
                            {
                                foreach (System.Xml.XmlNode nd in splithtml.SelectNodes(".//a[contains(@href, '" + thisId + ".html')]"))
                                {
                                    if (((System.Xml.XmlElement)nd).GetAttribute("href").Contains("#"))
                                        ((System.Xml.XmlElement)nd).SetAttribute("href", Regex.Replace(((System.Xml.XmlElement)nd).GetAttribute("href"), @"^.*?(#.*?)$", "$1", RegexOptions.Singleline));
                                    else
                                        ((System.Xml.XmlElement)nd).SetAttribute("href", "#");
                                }
                            }

                            sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + thisId + ".html", false, Encoding.UTF8);
                            string htmlBody = htmlTemplate1cpy + splithtml.OuterXml + htmlTemplate2;
                            // find tag span has class manual_  in tag p has class manual_ and add class manual_ to tag span with unicode
                            htmlBody = Regex.Replace(htmlBody, @"<p[^>]*?class=""MJS_oflow_step([^""]*?)""[^>]*?>(.*?)<span[^>]*?>(.*?)</span>(.*?)</p>", @"<p class=""MJS_oflow_step$1""><span class=""MJS_oflow_stepNum$2"">$3</span>$4</p>", RegexOptions.Singleline);
                            //find charactor è in tag span with class manual_ and replace 
                            htmlBody = Regex.Replace(htmlBody, @"<span class=""MJS_oflow_stepNum"">(è)</span>", @"<span class=""MJS_oflow_stepResult""></span>", RegexOptions.Singleline);
                            // find tag p has class manual_ and remove tag span with class manual_
                            htmlBody = Regex.Replace(htmlBody, @"<p[^>]*?class=""MJS_oflow_stepResult([^""]*?)""[^>]*?>(.*?)<span[^>]*?>(.*?)</span>(.*?)</p>", @"<p class=""MJS_oflow_stepResult"">$4</p>", RegexOptions.Singleline);
                            // find tag span has class manual_ and remove tag span in span
                            htmlBody = Regex.Replace(htmlBody, @"<span class=""MJS_oflow_stepNum""><span[^>]*?>(.*?)</span>(.*?)</span>", @"<span class=""MJS_oflow_stepNum"">$1$2</span>", RegexOptions.Singleline);

                            sw.Write(htmlBody);
                            sw.Close();
                        }
                    }

                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\search.js", false, Encoding.UTF8);
                    sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;").Replace(@"\u", @"\\u").Replace(@"\U", @"\\U")));
                    sw.Close();

                    if (!File.Exists(rootPath + "\\" + exportDir + "\\" + docid + "00000.html"))
                    {
                        sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + docid + "00000.html", false, Encoding.UTF8);
                        sw.Write(htmlCoverTemplate1 + htmlCoverTemplate2);
                        sw.Close();
                    }

                    log.WriteLine("Zipファイル作成");

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

                catch (Exception ex)
                {
                    load.Close();
                    load.Dispose();
                    //m_nowLoading.Abort();
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
                //m_nowLoading.Abort();
            }

            File.Delete(rootPath + "\\log.txt");

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

            //button3.Enabled = true;

            //ActiveDocumentのパスは、「WordAddIn1.application.ActiveDocument.Path」で取得できます。
            //index.htmlが出力されるとして、「WordAddIn1.application.ActiveDocument.Path + @"\index.html"」に
            //出力されるindex.htmlのパスという想定で、以下に出力後のHTMLをブラウザで閲覧するか否かの
            //メッセージボックス表示のコードを書いています。


            //DialogResult selectMess = MessageBox.Show(WordAddIn1.application.ActiveDocument.Path + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？", "HTML出力成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (selectMess == DialogResult.Yes)
            //{
            //    try
            //    {
            //        Process.Start(WordAddIn1.application.ActiveDocument.Path + @"\index.html");
            //    }
            //    catch
            //    {
            //        MessageBox.Show("HTMLの出力に失敗しました。", "HTML出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}

            /*以下は、次期対応変更履歴保存用コードの一部です。
            var activeDoc = WordAddIn1.application.ActiveDocument as Microsoft.Office.Interop.Word.Document;
            Word.Selection ws = WordAddIn1.application.Selection;
            string text = "No,Page,Type,Revision,User\r\n";
            foreach (Word.Revision r in activeDoc.Revisions)
            {
                string word = r.Range.Text;
                if(word.Contains("\r"))
                {
                    word = @"""" + word + @"""";
                    word = word.Replace("\r", "\n");
                }
                text += r.Index + "," + r.Range.Information[Word.WdInformation.wdActiveEndPageNumber] + "," + cordConvert((int)r.Type) + "," + word + "," + r.Author + "\r\n";
            }
            using (StreamWriter sw = new StreamWriter(@"./revision.csv", false, Encoding.UTF8))
            {
                sw.Write(text);
            }
            */
            application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }
    }
}
