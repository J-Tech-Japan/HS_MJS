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
using Microsoft.Office.Interop.Word;
using Table = Microsoft.Office.Interop.Word.Table;
using System.Drawing;
using Application = System.Windows.Forms.Application;

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

                        if (coverExist)
                        {
                            if (!Directory.Exists(paths.rootPath + "\\tmpcoverpic")) Directory.CreateDirectory(paths.rootPath + "\\tmpcoverpic");
                            string strOutFileName = paths.rootPath + "\\tmpcoverpic";

                            try
                            {
                                bool repeatUngroup = true;
                                while (repeatUngroup)
                                {
                                    repeatUngroup = false;
                                    foreach (Shape ws in docCopy.Shapes)
                                    {
                                        ws.Select();
                                        if (Globals.ThisAddIn.Application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1)
                                        {
                                            if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                            {
                                                ws.Ungroup();
                                                repeatUngroup = true;
                                            }
                                        }
                                    }
                                }


                                foreach (Shape ws in docCopy.Shapes)
                                {
                                    ws.Select();
                                    if (WordAddIn1.Globals.ThisAddIn.Application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1)
                                    {
                                        if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                                        {
                                            bool checkCanvas = true;
                                            while (checkCanvas)
                                            {
                                                checkCanvas = false;
                                                foreach (Word.Shape wsp in ws.CanvasItems)
                                                {
                                                    if (wsp.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                                    {
                                                        wsp.Ungroup();
                                                        checkCanvas = true;
                                                    }
                                                }
                                            }
                                            foreach (Word.Shape wsp in ws.CanvasItems)
                                            {
                                                wsp.Select();
                                                string tempSubTitle = "";
                                                try
                                                {
                                                    tempSubTitle = wsp.TextFrame.TextRange.Text;
                                                }
                                                catch { }
                                                if (!String.IsNullOrEmpty(tempSubTitle) && tempSubTitle != "/" && subTitle == "")
                                                {
                                                    subTitle = tempSubTitle;
                                                    break;
                                                }
                                            }
                                            if (String.IsNullOrEmpty(subTitle))
                                            {
                                                ws.Select();
                                                if (!Directory.Exists(paths.rootPath + "\\tmpcoverpic")) Directory.CreateDirectory(paths.rootPath + "\\tmpcoverpic");

                                                strOutFileName = paths.rootPath + "\\tmpcoverpic";
                                                byte[] vData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                                                if (vData != null)
                                                {
                                                    MemoryStream ms = new MemoryStream(vData);
                                                    Image temp = Image.FromStream(ms);
                                                    float aspectTemp = (float)temp.Width / (float)temp.Height;
                                                    if (aspectTemp > 2.683 || aspectTemp < 2.681)
                                                    {
                                                        biCount++;
                                                        temp.Save(strOutFileName + "\\" + biCount + ".png", ImageFormat.Png);
                                                    }
                                                }
                                            }
                                        }
                                        else if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                                        {
                                            ws.ConvertToInlineShape();
                                        }
                                    }
                                }

                                foreach (Shape ws in docCopy.Shapes)
                                {
                                    ws.Select();
                                    if (Globals.ThisAddIn.Application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1)
                                    {
                                        if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                                        {
                                            ws.ConvertToInlineShape();
                                        }
                                    }
                                }

                                if (isPattern1 || isPattern2)
                                {
                                    int productSubLogoCount = 0;

                                    foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
                                    {
                                        if (wp.get_Style().NameLocal == "MJS_製品ロゴ（メイン）")
                                        {
                                            try
                                            {
                                                foreach (Word.InlineShape wis in wp.Range.InlineShapes)
                                                {
                                                    wis.Range.Select();
                                                    Clipboard.Clear();
                                                    WordAddIn1.Globals.ThisAddIn.Application.Selection.CopyAsPicture();
                                                    Image img = Clipboard.GetImage();
                                                    img.Save(strOutFileName + "\\product_logo_main.png", ImageFormat.Png);

                                                    break; //get first product main logo only
                                                }
                                            }
                                            catch (Exception ex)
                                            {
                                                log.WriteLine("Error when extracting [MJS_製品ロゴ（メイン）]: " + ex.ToString());
                                            }
                                        }
                                        else if (wp.get_Style().NameLocal == "MJS_製品ロゴ（サブ）" && productSubLogoCount < 3)
                                        {
                                            try
                                            {
                                                List<string> productSubLogoFileNames = new List<string>();

                                                foreach (InlineShape wis in wp.Range.InlineShapes)
                                                {
                                                    wis.Range.Select();
                                                    Clipboard.Clear();
                                                    Globals.ThisAddIn.Application.Selection.CopyAsPicture();
                                                    Image img = Clipboard.GetImage();

                                                    productSubLogoCount++;
                                                    string subLogoFileName = string.Format("product_logo_sub{0}.png", productSubLogoCount);
                                                    img.Save(strOutFileName + "\\" + subLogoFileName, ImageFormat.Png);
                                                    productSubLogoFileNames.Add(subLogoFileName);

                                                    Clipboard.Clear();

                                                    if (productSubLogoCount == 3)
                                                    {
                                                        break; //get first 3 sub logos only
                                                    }
                                                }

                                                productSubLogoGroups.Add(productSubLogoFileNames);
                                            }
                                            catch (Exception ex)
                                            {
                                                log.WriteLine("Error when extracting [MJS_製品ロゴ（サブ）]: " + ex.ToString());
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    foreach (Word.InlineShape wis in docCopy.Sections[1].Range.InlineShapes)
                                    {
                                        byte[] vData = (byte[])wis.Range.EnhMetaFileBits;
                                        //MessageBox.Show(vData.Length.ToString());

                                        if (vData != null)
                                        {
                                            MemoryStream ms = new MemoryStream(vData);
                                            Image temp = Image.FromStream(ms);
                                            float aspectTemp = (float)temp.Width / (float)temp.Height;
                                            if ((float)temp.Height < 360) continue;
                                            if (aspectTemp > 12.225 && aspectTemp < 12.226) continue;
                                            if (aspectTemp > 2.681 && aspectTemp < 2.683) continue;
                                            biCount++;
                                            temp.Save(strOutFileName + "\\" + biCount + ".png", ImageFormat.Png);
                                        }
                                    }
                                }

                                Dictionary<string, float> dicStrFlo = new Dictionary<string, float>();

                                string[] coverPics = Directory.GetFiles(strOutFileName, "*.png", SearchOption.AllDirectories);

                                foreach (string coverPic in coverPics)
                                {
                                    using (FileStream fs = new FileStream(coverPic, FileMode.Open, FileAccess.Read))
                                    {
                                        dicStrFlo.Add(coverPic, (float)Image.FromStream(fs).Width * (float)Image.FromStream(fs).Height);
                                    }
                                }

                                List<KeyValuePair<string, float>> pairs = new List<KeyValuePair<string, float>>(dicStrFlo);
                                pairs.Sort(CompareKeyValuePair);

                                if (isPattern1 || isPattern2)
                                {
                                    for (int p = 0; p < pairs.Count; p++)
                                    {
                                        string destF = paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\" + Path.GetFileName(pairs[p].Key);

                                        if (File.Exists(destF))
                                        {
                                            File.Delete(destF);
                                        }

                                        File.Move(pairs[p].Key, destF);
                                    }
                                }
                                else
                                {
                                    for (int p = 0; p < pairs.Count; p++)
                                    {

                                        if (p == 0 || p + 1 != pairs.Count)
                                        {
                                            if (File.Exists(paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-4.png")) File.Delete(paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-4.png");
                                            File.Move(pairs[p].Key, paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-4.png");
                                        }
                                        else
                                        {
                                            using (Bitmap src = new Bitmap(pairs[p].Key))
                                            {
                                                int w = src.Width / 5;
                                                int h = src.Height / 5;
                                                Bitmap dst = new Bitmap(w, h);
                                                Graphics g = Graphics.FromImage(dst);
                                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bicubic;
                                                g.DrawImage(src, 0, 0, w, h);
                                                dst.Save(paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-background.png", ImageFormat.Png);
                                            }
                                            // Saves result.
                                            File.Delete(pairs[p].Key);
                                        }
                                    }
                                }

                                if (Directory.Exists(paths.rootPath + "\\tmpcoverpic")) Directory.Delete(paths.rootPath + "\\tmpcoverpic", true);
                            }
                            catch (Exception ex)
                            {
                                log.WriteLine(ex.ToString());
                            }
                        }

                        // ドキュメント末尾に移動し、一時キャンバスを追加
                        application.Selection.EndKey(WdUnits.wdStory);
                        object selectionRange = application.Selection.Range;
                        Shape temporaryCanvas = docCopy.Shapes.AddCanvas(0, 0, 1, 1, ref selectionRange);
                        temporaryCanvas.WrapFormat.Type = WdWrapType.wdWrapInline;

                        // キャンバス内の図形調整
                        AdjustCanvasShapes(docCopy);

                        // 一時キャンバス削除
                        temporaryCanvas.Delete();

                        // テーブル幅の自動調整
                        foreach (Table wt in docCopy.Tables)
                        {
                            if (wt.PreferredWidthType == WdPreferredWidthType.wdPreferredWidthPoints)
                                wt.AllowAutoFit = true;
                        }

                        // スタイル名の置換
                        foreach (Style ws in docCopy.Styles)
                            if (ws.NameLocal == "奥付タイトル")
                                ws.NameLocal = "titledef";


                        docCopy.SaveAs2(
                            paths.tmpHtmlPath,
                            WdSaveFormat.wdFormatFilteredHTML,
                            SaveNativePictureFormat: true
                        );

                        // ドキュメントを閉じる
                        docCopy.Close(false);

                        // ファイル解放待ち（100ms程度の遅延を入れる）
                        System.Threading.Thread.Sleep(100);

                        // 画像フォルダのコピー処理
                        log.WriteLine("画像フォルダ コピー");

                        bool isTmpDot = true;
                        CopyAndDeleteTemporaryImages(paths.tmpFolderForImagesSavedBySaveAs2Method, paths.rootPath, paths.exportDir, log);

                        // カバー情報の収集
                        //var coverInfo = CollectInfo(docCopy, application, paths, isPattern1, isPattern2, log);

                        // HTMLファイルの読み込みと加工
                        var htmlStr = ReadAndProcessHtml(paths.tmpHtmlPath, isTmpDot);
                        
                        // XMLへの変換と各種ノード取得
                        var (objXml, objToc, objBody) = LoadAndProcessXml(htmlStr, paths.docTitle);
                        
                        // CSSスタイルの処理
                        var (className, styleName, chapterSplitClass) = ProcessCssStyles(objXml);
                        
                        // index.htmlの書き出し
                        WriteIndexHtml(paths.indexHtmlPath, paths.docTitle, paths.docid, mergeScript);

                        // カバーテンプレートの生成
                        //var (htmlCoverTemplate1, htmlCoverTemplate2) = BuildCoverTemplates(assembly, paths, coverInfo, isEasyCloud, isEdgeTracker, isPattern1, isPattern2);

                        string htmlCoverTemplate1 = "";
                        htmlCoverTemplate1 += @"<!DOCTYPE HTML>" + "\n";
                        htmlCoverTemplate1 += @"<html>" + "\n";
                        htmlCoverTemplate1 += @"<head>" + "\n";
                        htmlCoverTemplate1 += @"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" + "\n";
                        htmlCoverTemplate1 += @" <meta name=""generator"" content=""Adobe RoboHelp 2017"" />" + "\n";
                        htmlCoverTemplate1 += @"<title>表紙</title>" + "\n";
                        htmlCoverTemplate1 += @"<link rel=""stylesheet"" href=""cover.css"" type=""text/css"" />" + "\n";
                        htmlCoverTemplate1 += @"<link rel=""stylesheet"" href=""font.css"" type=""text/css"" />" + "\n";
                        htmlCoverTemplate1 += @"<link rel=""StyleSheet"" href=""resp.css"" type=""text/css"" />" + "\n";
                        htmlCoverTemplate1 += @"<style type=""text/css"">" + "\n";
                        htmlCoverTemplate1 += @"<!--" + "\n";
                        htmlCoverTemplate1 += @"A:visited { color:#954F72; }" + "\n";
                        htmlCoverTemplate1 += @"A:link { color:#000000; }" + "\n";
                        htmlCoverTemplate1 += @"-->" + "\n";
                        htmlCoverTemplate1 += @"</style>" + "\n";
                        htmlCoverTemplate1 += @"<script type=""text/javascript"" language=""JavaScript"">" + "\n";
                        htmlCoverTemplate1 += @"//<![CDATA[" + "\n";
                        htmlCoverTemplate1 += @"function reDo() {" + "\n";
                        htmlCoverTemplate1 += @"  if (innerWidth != origWidth || innerHeight != origHeight)" + "\n";
                        htmlCoverTemplate1 += @"     location.reload();" + "\n";
                        htmlCoverTemplate1 += @"}" + "\n";
                        htmlCoverTemplate1 += @"if ((parseInt(navigator.appVersion) == 4) && (navigator.appName == ""Netscape"")) {" + "\n";
                        htmlCoverTemplate1 += @"   origWidth = innerWidth;" + "\n";
                        htmlCoverTemplate1 += @"   origHeight = innerHeight;" + "\n";
                        htmlCoverTemplate1 += @"   onresize = reDo;" + "\n";
                        htmlCoverTemplate1 += @"}" + "\n";
                        htmlCoverTemplate1 += @"onerror = null;" + "\n";
                        htmlCoverTemplate1 += @"//]]>" + "\n";
                        htmlCoverTemplate1 += @"</script>" + "\n";
                        htmlCoverTemplate1 += @"<style type=""text/css"">" + "\n";
                        htmlCoverTemplate1 += @"<!--" + "\n";
                        htmlCoverTemplate1 += @"div.WebHelpPopupMenu { position:absolute;" + "\n";
                        htmlCoverTemplate1 += @"left:0px;" + "\n";
                        htmlCoverTemplate1 += @"top:0px;" + "\n";
                        htmlCoverTemplate1 += @"z-index:4;" + "\n";
                        htmlCoverTemplate1 += @"visibility:hidden; }" + "\n";
                        htmlCoverTemplate1 += @"-->" + "\n";
                        if (isEdgeTracker)
                        {
                            htmlCoverTemplate1 += "\n";
                            htmlCoverTemplate1 += @"p.HyousiLogo {" + "\n";
                            htmlCoverTemplate1 += @"text-align       : center;" + "\n";
                            htmlCoverTemplate1 += @"margin-top       : 60pt;" + "\n";
                            htmlCoverTemplate1 += @"margin-bottom    : 40pt;" + "\n";
                            htmlCoverTemplate1 += @"margin-right     : 0mm;" + "\n";
                            htmlCoverTemplate1 += @"line-height      : 15pt;" + "\n";
                            htmlCoverTemplate1 += @"}" + "\n";
                            htmlCoverTemplate1 += "\n";
                            htmlCoverTemplate1 += @"div.HyousiBackground {" + "\n";
                            htmlCoverTemplate1 += @"display : table;" + "\n";
                            htmlCoverTemplate1 += @"width   : 100%;" + "\n";
                            htmlCoverTemplate1 += @"height  : 65px;" + "\n";
                            htmlCoverTemplate1 += @"}" + "\n";
                            htmlCoverTemplate1 += "\n";
                            htmlCoverTemplate1 += @"p.HyousiText {" + "\n";
                            htmlCoverTemplate1 += @"display             : table-cell;" + "\n";
                            htmlCoverTemplate1 += @"background-image    : url('pict/hyousi.png');" + "\n";
                            htmlCoverTemplate1 += @"background-repeat   : no-repeat;" + "\n";
                            htmlCoverTemplate1 += @"background-position : center;" + "\n";
                            htmlCoverTemplate1 += @"text-align          : center;" + "\n";
                            htmlCoverTemplate1 += @"vertical-align      : middle;" + "\n";
                            htmlCoverTemplate1 += @"font-size           : 1.8em;" + "\n";
                            htmlCoverTemplate1 += @"font-weight         : bold;" + "\n";
                            htmlCoverTemplate1 += @"color               : #FFF;" + "\n";
                            htmlCoverTemplate1 += @"letter-spacing      : 10px;" + "\n";
                            htmlCoverTemplate1 += @"}" + "\n";
                        }
                        htmlCoverTemplate1 += @"</style>" + "\n";
                        htmlCoverTemplate1 += @"</head>" + "\n";
                        string htmlCoverTemplate2 = "";


                        if (isEdgeTracker)
                        {
                            string[] hyousiGazo = { "EdgeTracker_logo50mm.png", "MJS_LOGO_255.gif", "hyousi.png" };
                            foreach (var hyousi in hyousiGazo)
                            {
                                Bitmap bmp = new Bitmap(assembly.GetManifestResourceStream("WordAddIn1.Resources." + hyousi));
                                bmp.Save(paths.rootPath + "\\" + paths.exportDir + "\\pict\\" + hyousi);
                            }
                            htmlCoverTemplate1 += @"<body>" + "\n";
                            htmlCoverTemplate1 += @"<p class=""HyousiLogo""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 553px;"" alt="""" src=""pict/EdgeTracker_logo50mm.png"" border=""0""></p>" + "\n";
                            htmlCoverTemplate1 += @"<div class=""HyousiBackground"">" + "\n";
                            htmlCoverTemplate1 += @"<p class=""HyousiText"">" + manualTitle + "</p>\n";
                            htmlCoverTemplate1 += @"</div>" + "\n";
                            htmlCoverTemplate1 += @"<div class=""product_trademarks"">" + "\n";
                            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
                            foreach (string trademarkText in trademarkTextList)
                            {
                                htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
                            }
                            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
                            htmlCoverTemplate1 += @"</div>" + "\n";
                            htmlCoverTemplate1 += @"<p style=""text-align: center; margin-top: 80pt;""><a href=""https://www.mjs.co.jp"" target=""_blank""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 255px;"" alt="""" src=""pict/MJS_LOGO_255.gif"" border=""0""></a></p>" + "\n";
                        }
                        else if (isEasyCloud)
                        {
                            if (File.Exists(paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-background.png"))
                                htmlCoverTemplate1 += @"<body style=""text-justify-trim: punctuation; background-image: url('template/images/cover-background.png');background-repeat: no-repeat; background-position: 0px 300px;"">" + "\n";
                            else
                                htmlCoverTemplate1 += @"<body>" + "\n";

                            htmlCoverTemplate1 += @"<p class=""manual_title"" style=""line-height: 130%;"">" + manualTitle + "</p>" + "\n";
                            htmlCoverTemplate1 += @"<p class=""manual_subtitle"">" + manualSubTitle + "</p>" + "\n";

                            if (File.Exists(paths.rootPath + "\\" + paths.exportDir + "\\template\\images\\cover-4.png"))
                                htmlCoverTemplate1 += @"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""><img src=""template/images/cover-4.png"" width=""650"" /></p>" + "\n";
                            else
                                htmlCoverTemplate1 += @"<p class=""manual_title"" style=""margin: 80px 0px 80px 100px; ""></p>" + "\n";

                            htmlCoverTemplate1 += @"<p class=""manual_version"">" + manualVersion + "</p>" + "\n";
                            htmlCoverTemplate1 += @"<div class=""product_trademarks"">" + "\n";
                            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
                            foreach (string trademarkText in trademarkTextList)
                            {
                                htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
                            }
                            htmlCoverTemplate1 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
                            htmlCoverTemplate1 += @"</div>" + "\n";
                            if (!String.IsNullOrEmpty(subTitle))
                            {
                                htmlCoverTemplate2 += @"<p style=""margin-left: 700px; margin-top: 150px; font-size: 15pt; font-family: メイリオ;" + "\n";
                                htmlCoverTemplate2 += @"    font-weight: bold;"">" + subTitle + "</p>" + "\n";
                                htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                                htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 10px;""" + "\n";
                                htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
                            }
                            else
                            {
                                htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                                htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 100px;""" + "\n";
                                htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
                            }
                            htmlCoverTemplate2 += @" </p>" + "\n";
                        }
                        else if (isPattern1)
                        {
                            htmlCoverTemplate2 += string.Format(@"<p class=""manual_title"" style=""line-height: 130%; "">{0}</p>" + "\n", !string.IsNullOrWhiteSpace(manualTitle) ? manualTitle : manualTitleCenter);
                            htmlCoverTemplate2 += string.Format(@"<p class=""manual_subtitle"">{0}</p>" + "\n", !string.IsNullOrWhiteSpace(manualSubTitle) ? manualSubTitle : manualSubTitleCenter);
                            htmlCoverTemplate2 += @"<p class=""product_logo_main_nosub"">" + "\n";
                            htmlCoverTemplate2 += @"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">" + "\n";
                            htmlCoverTemplate2 += @"</p>" + "\n";
                            htmlCoverTemplate2 += @"<div class=""product_trademarks"">" + "\n";
                            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
                            foreach (string trademarkText in trademarkTextList)
                            {
                                htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
                            }
                            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
                            htmlCoverTemplate2 += @"</div>" + "\n";
                            htmlCoverTemplate2 += @"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>" + "\n";
                            htmlCoverTemplate2 += @"</p>" + "\n";
                        }
                        else if (isPattern2)
                        {
                            htmlCoverTemplate2 += @"<p class=""product_logo_main"">" + "\n";
                            htmlCoverTemplate2 += @"  <img src = ""template/images/product_logo_main.png"" alt=""製品ロゴ（メイン）"">" + "\n";
                            htmlCoverTemplate2 += @"</p>" + "\n";
                            htmlCoverTemplate2 += @"<div class=""product_logo_sub"">" + "\n";
                            foreach (List<string> subLogoGroup in productSubLogoGroups)
                            {
                                htmlCoverTemplate2 += @"<div>" + "\n";
                                foreach (string subLogoFileName in subLogoGroup)
                                {
                                    htmlCoverTemplate2 += string.Format(@"  <img src = ""template/images/{0}"" alt=""製品ロゴ（サブ）"">" + "\n", subLogoFileName);
                                }
                                htmlCoverTemplate2 += @"</div>" + "\n";
                            }
                            htmlCoverTemplate2 += @"</div>" + "\n";
                            htmlCoverTemplate2 += string.Format(@"<p class=""manual_title_center"" style=""line-height: 130%; "">{0}</p>" + "\n", !string.IsNullOrWhiteSpace(manualTitleCenter) ? manualTitleCenter : manualTitle);
                            htmlCoverTemplate2 += string.Format(@"<p class=""manual_subtitle_center"">{0}</p>" + "\n", !string.IsNullOrWhiteSpace(manualSubTitleCenter) ? manualSubTitleCenter : manualSubTitle);
                            htmlCoverTemplate2 += string.Format(@"<p class=""manual_version_center"">{0}</p>" + "\n", !string.IsNullOrWhiteSpace(manualVersionCenter) ? manualVersionCenter : manualVersion);
                            htmlCoverTemplate2 += @"<div class=""product_trademarks"">" + "\n";
                            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_title"">{0}</p>" + "\n", trademarkTitle);
                            foreach (string trademarkText in trademarkTextList)
                            {
                                htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_text"">{0}</p>" + "\n", trademarkText);
                            }
                            htmlCoverTemplate2 += string.Format(@"  <p class=""trademark_right"">{0}</p>" + "\n", trademarkRight);
                            htmlCoverTemplate2 += @"</div>" + "\n";
                            htmlCoverTemplate2 += @"<p><a href = ""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）"" style=""margin-left: 700px; margin-top: 100px;"" width=""132"" height=""48"" /></a>" + "\n";
                            htmlCoverTemplate2 += @"</p>" + "\n";
                        }

                        htmlCoverTemplate2 += @"<script type=""text/javascript"" language=""javascript1.2"">//<![CDATA[" + "\n";
                        htmlCoverTemplate2 += @"<!--" + "\n";
                        htmlCoverTemplate2 += @"if (window.writeIntopicBar)" + "\n";
                        htmlCoverTemplate2 += @"   writeIntopicBar(0);" + "\n";
                        htmlCoverTemplate2 += @"//-->" + "\n";
                        htmlCoverTemplate2 += @"//]]></script>" + "\n";
                        htmlCoverTemplate2 += @"</body>" + "\n";
                        htmlCoverTemplate2 += @"</html>" + "\n";
                        // HTMLテンプレートの生成
                        var htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript);
                        var htmlTemplate2 = "</body>\n</html>\n";
                        
                        // 検索用JSの生成
                        var searchJs = BuildSearchJs();
                        
                        // 目次・本文ノードの参照取得
                        XmlNode objTocCurrent = objToc.DocumentElement;
                        XmlNode objBodyCurrent = objBody.DocumentElement;
                        
                        // 目次・本文の生成
                        BuildTocBodyFromXml(objXml, objBody, objToc, chapterSplitClass, styleName, paths.docid, bookInfoDef, ref objBodyCurrent, ref objTocCurrent, load);
                        
                        // 本文IDの設定
                        SetDefaultBodyId(objBody, paths.docid);
                        
                        // 目次ファイルの生成
                        ExportTocAsJsFiles(objToc, paths.rootPath, paths.exportDir, mergeScript);
                        
                        // 一時XMLの解放
                        objXml = null;
                        
                        // 一時HTMLの削除
                        File.Delete(paths.tmpHtmlPath);
                        
                        // XMLノードのクリーンアップ
                        CleanUpXmlNodes(objBody);
                        
                        // 検索用ファイルの生成
                        GenerateSearchFiles(objBody, paths.rootPath, paths.exportDir, paths.docid, htmlTemplate1, htmlTemplate2, htmlCoverTemplate1, htmlCoverTemplate2, objToc, mergeScript, searchJs);

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