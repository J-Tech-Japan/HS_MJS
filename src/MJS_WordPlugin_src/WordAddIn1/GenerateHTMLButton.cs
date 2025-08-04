// GenerateHTMLButton.cs

using System;
using System.Collections.Generic;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
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

            // アウトラインレベルと見出しテキストをメッセージボックスで表示（動作確認用）
            //ShowHeadingsWithOutlineLevels();


            try
            {
                // TODO: 指定スタイルの見出しを取得
                //var headings = GetHeadingsByStyles(new List<string> { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）" });
                //var specifiedHeadings = new List<string> { "はじめに", "マニュアル内の記号・表記について" };

                //List<string> commonHeadings = headings.Intersect(specifiedHeadings).ToList();

                // "##検索対象外トピック##" というコメントがついている、特定スタイルの見出しを取得
                var headingsWithComment = GetHeadingsWithComment(new List<string> { "見出し 1,MJS_見出し 1", "見出し 2,MJS_見出し 2", "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）" }, "##検索対象外トピック##");

                // commonHeadingsとheadingsWithCommentを重複要素なしで結合
                //var mergedHeadings = commonHeadings.Union(headingsWithComment).ToList();

                // 特定の見出しとその配下の見出しを取得
                var headings = GetSpecificHeadingsWithSubheadings();

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
                        //PrepareHtmlTemplates(assembly, paths.rootPath, paths.exportDir);

                        log.WriteLine("テンプレートデータ準備");

                        using (Stream stream = assembly.GetManifestResourceStream("WordAddIn1.htmlTemplates.zip"))
                        {
                            FileStream fs = File.Create(Path.Combine(paths.rootPath, "htmlTemplates.zip"));
                            stream.Seek(0, SeekOrigin.Begin);
                            stream.CopyTo(fs);
                            fs.Close();
                        }

                        if (Directory.Exists(Path.Combine(paths.rootPath, "htmlTemplates")))
                        {
                            Directory.Delete(Path.Combine(paths.rootPath, "htmlTemplates"), true);
                        }

                        System.IO.Compression.ZipFile.ExtractToDirectory(Path.Combine(paths.rootPath, "htmlTemplates.zip"), paths.rootPath);

                        if (Directory.Exists(Path.Combine(paths.rootPath, paths.exportDir)))
                        {
                            Directory.Delete(Path.Combine(paths.rootPath, paths.exportDir), true);
                        }

                        if (Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.Delete(Path.Combine(paths.rootPath, "tmpcoverpic"), true);
                        Directory.Move(Path.Combine(paths.rootPath, "htmlTemplates"), Path.Combine(paths.rootPath, paths.exportDir));

                        File.Delete(Path.Combine(paths.rootPath, "htmlTemplates.zip"));

                        string docid = Regex.Replace(paths.docName, "^(.{3}).+$", "$1");
                        string docTitle = Regex.Replace(paths.docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");

                        string zipDirPath = Path.Combine(paths.rootPath, docid + "_" + paths.exportDir + "_" + DateTime.Today.ToString("yyyyMMdd"));
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

                        #region
                        List<List<string>> productSubLogoGroups = new List<List<string>>();

                        if (coverExist)
                        {
                            if (!Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.CreateDirectory(Path.Combine(paths.rootPath, "tmpcoverpic"));
                            string strOutFileName = Path.Combine(paths.rootPath, "tmpcoverpic");

                            try
                            {
                                bool repeatUngroup = true;
                                while (repeatUngroup)
                                {
                                    repeatUngroup = false;
                                    foreach (Shape ws in docCopy.Shapes)
                                    {
                                        ws.Select();
                                        if (Globals.ThisAddIn.Application.Selection.Information[WdInformation.wdActiveEndSectionNumber] == 1)
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
                                    if (Globals.ThisAddIn.Application.Selection.Information[WdInformation.wdActiveEndSectionNumber] == 1)
                                    {
                                        if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                                        {
                                            bool checkCanvas = true;
                                            while (checkCanvas)
                                            {
                                                checkCanvas = false;
                                                foreach (Shape wsp in ws.CanvasItems)
                                                {
                                                    if (wsp.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                                    {
                                                        wsp.Ungroup();
                                                        checkCanvas = true;
                                                    }
                                                }
                                            }
                                            foreach (Shape wsp in ws.CanvasItems)
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
                                                if (!Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.CreateDirectory(Path.Combine(paths.rootPath, "tmpcoverpic"));

                                                strOutFileName = Path.Combine(paths.rootPath, "tmpcoverpic");
                                                byte[] vData = (byte[])Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
                                                if (vData != null)
                                                {
                                                    MemoryStream ms = new MemoryStream(vData);
                                                    Image temp = Image.FromStream(ms);
                                                    float aspectTemp = (float)temp.Width / (float)temp.Height;
                                                    if (aspectTemp > 2.683 || aspectTemp < 2.681)
                                                    {
                                                        biCount++;
                                                        temp.Save(Path.Combine(strOutFileName, biCount + ".png"), ImageFormat.Png);
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
                                                foreach (InlineShape wis in wp.Range.InlineShapes)
                                                {
                                                    wis.Range.Select();
                                                    Clipboard.Clear();
                                                    Globals.ThisAddIn.Application.Selection.CopyAsPicture();
                                                    Image img = Clipboard.GetImage();
                                                    img.Save(Path.Combine(strOutFileName, "product_logo_main.png"), ImageFormat.Png);

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
                                                    img.Save(Path.Combine(strOutFileName, subLogoFileName), ImageFormat.Png);
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
                                    foreach (InlineShape wis in docCopy.Sections[1].Range.InlineShapes)
                                    {
                                        byte[] vData = (byte[])wis.Range.EnhMetaFileBits;

                                        if (vData != null)
                                        {
                                            MemoryStream ms = new MemoryStream(vData);
                                            Image temp = Image.FromStream(ms);
                                            float aspectTemp = (float)temp.Width / (float)temp.Height;
                                            if ((float)temp.Height < 360) continue;
                                            if (aspectTemp > 12.225 && aspectTemp < 12.226) continue;
                                            if (aspectTemp > 2.681 && aspectTemp < 2.683) continue;
                                            biCount++;
                                            temp.Save(Path.Combine(strOutFileName, biCount + ".png"), ImageFormat.Png);
                                        }
                                    }
                                }

                                // 一時フォルダ内のPNG画像をすべて取得し、
                                // 画像ごとの面積（幅×高さ）を計算し、
                                // 面積順に並べたリスト（pairs）を作る
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
                                        string destF = Path.Combine(paths.rootPath, paths.exportDir, "template", "images", Path.GetFileName(pairs[p].Key));

                                        if (File.Exists(destF))
                                        {
                                            File.Delete(destF);
                                        }

                                        File.Move(pairs[p].Key, destF);
                                    }
                                }
                                else
                                {
                                    // pairs（画像ファイルパスと面積のペアのリスト）をループし、
                                    // 画像ファイルを用途別にコピー・リサイズ・リネームする
                                    for (int p = 0; p < pairs.Count; p++)
                                    {
                                        // 先頭画像または最後以外の画像の場合
                                        if (p == 0 || p + 1 != pairs.Count)
                                        {
                                            // cover-4.png として保存（既存なら削除してから移動）
                                            if (File.Exists(Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-4.png")))
                                                File.Delete(Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-4.png"));
                                            File.Move(pairs[p].Key, Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-4.png"));
                                        }
                                        else
                                        {
                                            // 最後の画像の場合
                                            // 1/5サイズに縮小して cover-background.png として保存
                                            using (Bitmap src = new Bitmap(pairs[p].Key))
                                            {
                                                int w = src.Width / 5;
                                                int h = src.Height / 5;
                                                Bitmap dst = new Bitmap(w, h);
                                                Graphics g = Graphics.FromImage(dst);
                                                g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bicubic;
                                                g.DrawImage(src, 0, 0, w, h);
                                                dst.Save(Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-background.png"), ImageFormat.Png);
                                            }
                                            File.Delete(pairs[p].Key);
                                        }
                                    }
                                }

                                // cover-4.pngが存在しない場合、セクション1の最初のShape画像をcover-4.pngとして保存
                                //string cover4Path = Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-4.png");
                                //if (!File.Exists(cover4Path))
                                //{
                                //    var section1 = docCopy.Sections[1];
                                //    foreach (Word.Shape shape in section1.Range.ShapeRange)
                                //    {
                                //        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                                //        {
                                //            try
                                //            {
                                //                shape.Select();
                                //                Clipboard.Clear();
                                //                Globals.ThisAddIn.Application.Selection.CopyAsPicture();
                                //                Image img = Clipboard.GetImage();
                                //                if (img != null)
                                //                {
                                //                    img.Save(cover4Path, ImageFormat.Png);
                                //                }
                                //            }
                                //            catch { }
                                //            break; // 最初のShapeのみ
                                //        }
                                //    }
                                //}

                                if (Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.Delete(Path.Combine(paths.rootPath, "tmpcoverpic"), true);
                            }
                            catch (Exception ex)
                            {
                                log.WriteLine(ex.ToString());
                            }
                        }

                        #endregion
                        
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

                        // ドキュメントをHTML形式で保存
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

                        // HTMLファイルの読み込みと加工
                        var htmlStr = ReadAndProcessHtml(paths.tmpHtmlPath, isTmpDot);
                        
                        // XMLへの変換と各種ノード取得
                        var (objXml, objToc, objBody) = LoadAndProcessXml(htmlStr, paths.docTitle);
                        
                        // CSSスタイルの処理
                        var (className, styleName, chapterSplitClass) = ProcessCssStyles(objXml);
                        
                        // index.htmlの書き出し
                        WriteIndexHtml(paths.indexHtmlPath, paths.docTitle, paths.docid, mergeScript);

                        // カバーテンプレートの生成
                        string htmlCoverTemplate1 = BuildHtmlCoverTemplate1Header();

                        if (isEdgeTracker)
                        {
                            htmlCoverTemplate1 += BuildEdgeTrackerCoverCss();
                        }

                        htmlCoverTemplate1 += @"</style>" + "\n";
                        htmlCoverTemplate1 += @"</head>" + "\n";

                        string htmlCoverTemplate2 = "";

                        if (isEdgeTracker)
                        {
                            htmlCoverTemplate1 += BuildEdgeTrackerCoverHtml(
                                assembly,
                                paths.rootPath,
                                paths.exportDir,
                                manualTitle,
                                trademarkTitle,
                                trademarkTextList,
                                trademarkRight
                             );
                        }
                        else if (isEasyCloud)
                        {
                            htmlCoverTemplate1 += BuildEasyCloudCoverHtml(
                                paths.rootPath,
                                paths.exportDir,
                                manualTitle,
                                manualSubTitle,
                                manualVersion,
                                trademarkTitle,
                                trademarkTextList,
                                trademarkRight
                            );

                            //if (!String.IsNullOrEmpty(subTitle))
                            //{
                            //    htmlCoverTemplate2 += @"<p style=""margin-left: 700px; margin-top: 150px; font-size: 15pt; font-family: メイリオ;" + "\n";
                            //    htmlCoverTemplate2 += @"    font-weight: bold;"">" + subTitle + "</p>" + "\n";
                            //    htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                            //    htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 10px;""" + "\n";
                            //    htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
                            //}

                            //else
                            //{
                            //    htmlCoverTemplate2 += @"<p><a href=""http://www.mjs.co.jp"" target=""_blank""><img src=""template/images/cover-3.png"" alt=""株式会社ミロク情報サービス （http://www.mjs.co.jp）""" + "\n";
                            //    htmlCoverTemplate2 += @"                                        style=""margin-left: 700px; margin-top: 100px;""" + "\n";
                            //    htmlCoverTemplate2 += @"                                        width=""132"" height=""48"" /></a>" + "\n";
                            //}

                            htmlCoverTemplate2 += BuildEasyCloudSubTitleSection(subTitle);
                            htmlCoverTemplate2 += @" </p>" + "\n";
                        }
                        else if (isPattern1)
                        {
                            htmlCoverTemplate2 += GeneratePattern1CoverHtml(
                                manualTitle,
                                manualTitleCenter,
                                manualSubTitle,
                                manualSubTitleCenter,
                                trademarkTitle,
                                trademarkTextList,
                                trademarkRight
                            );
                        }
                        else if (isPattern2)
                        {
                         
                            htmlCoverTemplate2 += GeneratePattern2CoverHtml(
                                productSubLogoGroups,
                                manualTitleCenter,
                                manualTitle,
                                manualSubTitleCenter,
                                manualSubTitle,
                                manualVersionCenter,
                                manualVersion,
                                trademarkTitle,
                                trademarkTextList,
                                trademarkRight
                            );
                        }

                        htmlCoverTemplate2 += BuildHtmlCoverFooter();

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

                        // TODO: 検索ブロックの削除
                        foreach (string heading in headings)
                        {
                            RemoveSearchBlockByTitle(
                            heading,
                            paths.rootPath,
                            paths.exportDir);
                        }

                        foreach (string heading in headingsWithComment)
                        {
                            RemoveSearchBlockByTitle(
                            heading,
                            paths.rootPath,
                            paths.exportDir);
                        }

                        //foreach (string heading in mergedHeadings)
                        //{
                        //    RemoveSearchBlockByTitle(
                        //    heading,
                        //    paths.rootPath,
                        //    paths.exportDir);
                        //}

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

                        // tmpcoverpicのクリーンアップ（必要であれば）
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
    }
}