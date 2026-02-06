// GenerateHTMLButton.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Xml;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Tools.Ribbon;
using Table = Microsoft.Office.Interop.Word.Table;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        /// <summary>
        /// HTML出力ボタンのイベントハンドラー（高画質画像抽出を有効にして実行）
        /// </summary>
        private void GenerateHTMLButton(object sender, RibbonControlEventArgs e)
        {
            // デフォルトで高画質画像抽出を有効にし、beta版モードはfalseに設定
            GenerateHTMLButton(sender, e, extractHighQualityImages: true, isBetaMode: true);
        }

        /// <summary>
        /// HTML出力処理の本体
        /// </summary>
        /// <param name="sender">イベント送信元</param>
        /// <param name="e">イベント引数</param>
        /// <param name="extractHighQualityImages">高画質画像抽出機能を実行するかどうか</param>
        /// <param name="isBetaMode">beta版モードかどうか（trueの場合、詳細ログとCSV出力を実行）</param>
        private void GenerateHTMLButton(object sender, RibbonControlEventArgs e, bool extractHighQualityImages, bool isBetaMode)
        {
            // HTML出力フラグをON
            blHTMLPublish = true;

            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;

            // beta版モードの場合のみ、アクティブドキュメントのフォルダにログを出力するよう設定
            if (isBetaMode)
            {
                Utils.ConfigureLogToDocumentFolder(application);
            }

            // 現在の表示モードを保存
            var defaultView = application.ActiveWindow.View.Type;

            // ローダーフォームを表示
            loader load = new loader();
            load.Show();

            try
            {
                // "##検索対象外トピック##" というコメントがついている、特定スタイルの見出しを取得
                var headingsWithComment = GetHeadingsWithComment(new List<string> { "見出し 1,MJS_見出し 1", "見出し 2,MJS_見出し 2", "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）" }, "##検索対象外トピック##");

                // 特定の見出しとその配下の見出しを取得
                var headings = GetSpecificHeadingsWithSubheadings();

                // 重複なしで結合
                var mergedHeadings = headingsWithComment.Union(headings).ToList();

                // 前処理（ドキュメントや環境のチェック）
                if (!PreProcess(application, activeDocument, load)) return;

                // 出力先フォルダ名を取得
                var webHelpFolderName = GetWebHelpFolderName(activeDocument);

                // 書籍情報の作成・取得
                if (!makeBookInfo(load)) { load.Close(); load.Dispose(); return; }

                // マージスクリプト情報の収集
                var mergeScript = CollectMergeScriptDict(activeDocument);

                // カバー選択ダイアログの処理
                if (!HandleCoverSelection(load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2, out bool isPattern3)) return;

                // ローダーを可視化
                load.Visible = true;

                // すべての変更履歴を反映
                activeDocument.AcceptAllRevisions();

                // 各種パスの準備
                var paths = PreparePaths(activeDocument, webHelpFolderName);

                System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                PrepareHtmlTemplates(assembly, paths.rootPath, paths.exportDirPath);

                // ドキュメントを一時HTML用にコピー
                var docCopy = CopyDocumentToHtml(application);

                using (StreamWriter log = new StreamWriter(paths.logPath, false, Encoding.UTF8))
                {
                    bool isError = false;
                    log.WriteLine("Number of sections: " + docCopy.Sections.Count);

                    try
                    {
                        // 高画質画像抽出機能が有効な場合のみ実行
                        List<Utils.ExtractedImageInfo> extractedImages = null;
                        if (extractHighQualityImages)
                        {
                            // 高画質の画像とキャンバスの抽出
                            log.WriteLine("高画質画像とキャンバスの抽出開始");

                            // パラメータオブジェクトの作成
                            var imageExtractionOptions = new Utils.ImageExtractionOptions
                            {
                                IncludeInlineShapes = true,      // インライン図形を抽出
                                IncludeShapes = true,            // フローティング図形を抽出
                                IncludeFreeforms = false,        // フリーフォーム図形は抽出しない
                                AddMarkers = true,               // マーカーを追加
                                SkipCoverMarkers = false,         // 表紙の画像にはマーカーをつけない
                                MinOriginalWidth = 50.0f,        // 元画像の最小幅（ポイント）
                                MinOriginalHeight = 60.0f,       // 元画像の最小高さ（ポイント）
                                IncludeMjsTableImages = true,    // MJS_画像（表内）スタイルの画像を抽出
                                MaxOutputWidth = 1024,           // 出力画像の最大幅
                                MaxOutputHeight = 1024,           // 出力画像の最大高さ
                                OutputScaleMultiplier = 1.4f,     // 出力スケール倍率
                                //DisableResize = true    // リサイズ無効化
                            };

                            extractedImages = Utils.ExtractImagesFromWord(
                                docCopy,
                                Path.Combine(paths.rootPath, paths.exportDir, "extracted_images"),
                                imageExtractionOptions
                            );

                            // 抽出統計をログに出力
                            if (extractedImages != null)
                            {
                                log.WriteLine($"画像抽出完了: {extractedImages.Count}個の画像を抽出しました");
                            }
                            else
                            {
                                log.WriteLine("画像抽出結果: 抽出された画像はありません");
                            }
                        }
                        else
                        {
                            log.WriteLine("高画質画像抽出機能: スキップ（extractHighQualityImages = false）");
                        }

                        // beta版モードの場合のみCSVファイルに出力
                        if (isBetaMode)
                        {
                            log.WriteLine("beta版モード: 画像比較CSVファイル出力開始");
                            Utils.ExportCompleteWidthHeightComparisonListToCsvFile(extractedImages, Path.Combine(paths.rootPath, "complete_comparison.csv"));
                            log.WriteLine("beta版モード: 画像比較CSVファイル出力完了");
                        }
                        else
                        {
                            log.WriteLine("正式版モード: 画像比較CSVファイル出力をスキップ");
                        }

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
                            if (!Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.CreateDirectory(Path.Combine(paths.rootPath, "tmpcoverpic"));
                            string strOutFileName = Path.Combine(paths.rootPath, "tmpcoverpic");

                            try
                            {
                                //UngroupAllShapesInFirstSection(docCopy, application);

                                ProcessCanvasAndPictureShapesInFirstSection(docCopy, application, ref subTitle, ref biCount, strOutFileName, paths);

                                ConvertPictureShapesToInlineInFirstSection(docCopy, application);

                                if (isPattern1 || isPattern2)
                                {
                                    ExtractProductLogosPattern1Pattern2(docCopy, application, strOutFileName, productSubLogoGroups);
                                }
                                else if (isPattern3)
                                {
                                    // LucaTech GX用の処理（pattern3.pngを使用）
                                    ExtractProductLogosPattern3(docCopy, application, strOutFileName, productSubLogoGroups);
                                }
                                else
                                {
                                    ExtractInlineShapesDefaultPattern(docCopy, strOutFileName, ref biCount);
                                }

                                // 一時フォルダ内のPNG画像をすべて取得し、
                                // 画像ごとの面積（幅×高さ）を計算し、
                                // 面積順に並べたリスト（pairs）を作る
                                Dictionary<string, float> dicStrFlo = new Dictionary<string, float>();

                                string[] coverPics = Directory.GetFiles(strOutFileName, "*.png", SearchOption.AllDirectories);

                                foreach (string coverPic in coverPics)
                                {
                                    using (FileStream fs = new FileStream(coverPic, FileMode.Open, FileAccess.Read))
                                    using (Image img = Image.FromStream(fs))
                                    {
                                        dicStrFlo.Add(coverPic, (float)img.Width * (float)img.Height);
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
                                else if (isPattern3)
                                {
                                    // LucaTech GX用の処理 - Pattern1と同様
                                    for (int p = 0; p < pairs.Count; p++)
                                    {
                                        string destF = Path.Combine(paths.rootPath, paths.exportDir, "template", "images", Path.GetFileName(pairs[p].Key));

                                        if (File.Exists(destF))
                                        {
                                            File.Delete(destF);
                                        }

                                        File.Move(pairs[p].Key, destF);
                                    }

                                    // pattern3.pngをcover-4.pngとしてコピー（リソースからpattern3を使用）
                                    try
                                    {
                                        using (var pattern3Image = Properties.Resources.pattern3)
                                        {
                                            string cover4Dest = Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-4.png");

                                            if (File.Exists(cover4Dest))
                                            {
                                                File.Delete(cover4Dest);
                                            }

                                            pattern3Image.Save(cover4Dest, ImageFormat.Png);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        log.WriteLine($"Pattern3画像保存エラー: {ex.Message}");
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
                                                using (Bitmap dst = new Bitmap(w, h))
                                                using (Graphics g = Graphics.FromImage(dst))
                                                {
                                                    g.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.Bicubic;
                                                    g.DrawImage(src, 0, 0, w, h);
                                                    dst.Save(Path.Combine(paths.rootPath, paths.exportDir, "template", "images", "cover-background.png"), ImageFormat.Png);
                                                }
                                            }
                                            File.Delete(pairs[p].Key);
                                        }
                                    }
                                }

                                if (Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic"))) Directory.Delete(Path.Combine(paths.rootPath, "tmpcoverpic"), true);
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

                        Utils.AdjustCanvasShapes(docCopy,
                            heightExpansion: 30.0f,      // 高さ拡張
                            positionOffset: 0.59f        // 位置オフセット
                        );

                        // 一時キャンバス削除
                        temporaryCanvas.Delete();

                        // テーブル幅の自動調整
                        foreach (Table wt in docCopy.Tables)
                        {
                            if (wt.PreferredWidthType == WdPreferredWidthType.wdPreferredWidthPoints)
                            {
                                // 指定されたスタイルを持つセルを含むテーブルは自動調整を適用しない
                                var excludeStyles = new[] { "MJS_コラム-本文", "MJS_コラム-タイトル" };
                                if (!ShouldApplyAutoFitToTable(wt, excludeStyles))
                                {
                                    continue;
                                }

                                wt.AllowAutoFit = true;
                            }
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
                        CopyAndDeleteTemporaryImages(paths.tmpFolderForImagesSavedBySaveAs2Method, paths.rootPath, paths.exportDir);

                        // HTMLファイルの読み込みと加工
                        var htmlStr = ReadAndProcessHtml(paths.tmpHtmlPath, isTmpDot);

                        // XMLへの変換と各種ノード取得
                        var (objXml, objToc, objBody) = LoadAndProcessXml(htmlStr, paths.docTitle);

                        // CSSスタイルの処理
                        var (className, styleName, chapterSplitClass) = ProcessCssStyles(objXml);

                        // index.htmlの書き出し
                        WriteIndexHtml(paths.indexHtmlPath, paths.docTitle, paths.docid, mergeScript, isPattern3);

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
                        else if (isPattern3)
                        {
                            htmlCoverTemplate2 += GeneratePattern3CoverHtml(
                                manualTitle,
                                manualTitleCenter,
                                manualSubTitle,
                                manualSubTitleCenter,
                                trademarkTitle,
                                trademarkTextList,
                                trademarkRight
                            );
                        }

                        htmlCoverTemplate2 += BuildHtmlCoverFooter();

                        // HTMLテンプレートの生成
                        var htmlTemplate1 = BuildHtmlTemplate1(title4Collection, mergeScript, paths.rootPath, paths.exportDir);
                        var htmlTemplate2 = "</body>\n</html>\n";

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
                        GenerateSearchFiles(objBody, paths.rootPath, paths.exportDir, paths.docid, htmlTemplate1, htmlTemplate2, htmlCoverTemplate1, htmlCoverTemplate2, objToc, mergeScript);

                        // AppData/Local/Tempから画像をwebhelpフォルダにコピーする
                        CopyImagesFromAppDataLocalTemp(activeDocument.FullName);

                        // 高画質画像抽出機能が有効な場合のみ、画像マーカー処理を実行
                        if (extractHighQualityImages)
                        {
                            // 画像マーカーの処理（HTMLファイル内の[IMAGEMARKER: xxx]を処理し、画像参照を更新）
                            log.WriteLine("画像マーカー処理");
                            try
                            {
                                int processedFiles = Utils.ProcessImageMarkersInWebhelp(paths.exportDirPath, "extracted_images");
                                log.WriteLine($"画像マーカー処理完了: {processedFiles}個のファイルを処理しました");
                            }
                            catch (Exception ex)
                            {
                                log.WriteLine($"画像マーカー処理エラー: {ex.Message}");
                            }

                            // search.jsファイルからイメージマーカーを削除
                            log.WriteLine("search.jsファイルのイメージマーカー削除");
                            try
                            {
                                int removedMarkers = Utils.RemoveImageMarkersFromSearchJsInDirectory(paths.exportDirPath);
                                if (removedMarkers >= 0)
                                {
                                    log.WriteLine($"search.jsファイルからイメージマーカー削除完了: {removedMarkers}個のマーカーを削除しました");
                                }
                                else
                                {
                                    log.WriteLine("search.jsファイルのイメージマーカー削除でエラーが発生しました");
                                }
                            }
                            catch (Exception ex)
                            {
                                log.WriteLine($"search.jsファイルのイメージマーカー削除エラー: {ex.Message}");
                            }
                        }
                        else
                        {
                            log.WriteLine("画像マーカー処理: スキップ（extractHighQualityImages = false）");
                            log.WriteLine("search.jsファイルのイメージマーカー削除: スキップ（extractHighQualityImages = false）");
                        }

                        // HTMLファイルからシンプルなspanタグを削除
                        try
                        {
                            string[] excludeFiles = { "index.html", "indexBase.html" };
                            int processedHtmlFiles = Utils.RemoveSimpleSpanTagsFromHtmlFolder(paths.exportDirPath, excludeFiles);
                            if (processedHtmlFiles >= 0)
                            {
                                log.WriteLine($"HTMLファイルのシンプルなspanタグ削除完了: {processedHtmlFiles}個のファイルを処理しました（index.htmlは除外）");
                            }
                            else
                            {
                                log.WriteLine("HTMLファイルのシンプルなspanタグ削除でエラーが発生しました");
                            }
                        }
                        catch (Exception ex)
                        {
                            log.WriteLine($"HTMLファイルのシンプルなspanタグ削除エラー: {ex.Message}");
                        }

                        // 検索ブロック削除
                        foreach (string heading in mergedHeadings)
                        {
                            RemoveSearchBlockByTitle(
                                heading,
                                paths.rootPath,
                                paths.exportDir);
                        }

                        // 検索除外する見出しリストの生成
                        if (headingsWithComment.Count > 0)
                        {
                            Utils.WriteLinesToFile(
                                paths.exportDirPath,
                                "headingsWithComment.txt",
                                headingsWithComment
                            );
                        }

                        if (headings.Count > 0)
                        {
                            Utils.WriteLinesToFile(
                                paths.exportDirPath,
                                "headings.txt",
                                headings
                            );
                        }

                        // Zipアーカイブの生成
                        log.WriteLine("Zipファイル作成");
                        GenerateZipArchive(paths.zipDirPath, paths.rootPath, paths.exportDir, paths.headerDir, paths.docFullName, paths.docName);
                    }
                    catch (Exception)
                    {
                        isError = true;
                        //HandleException(ex, log, load);
                        load.Close();
                        load.Dispose();
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

                        // tmpcoverpicフォルダを削除
                        if (Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic")))
                        {
                            Directory.Delete(Path.Combine(paths.rootPath, "tmpcoverpic"), true);
                        }

                        // indexBase.htmlファイルを削除
                        string indexBaseHtmlPath = Path.Combine(paths.exportDirPath, "indexBase.html");
                        if (File.Exists(indexBaseHtmlPath))
                        {
                            try
                            {
                                File.Delete(indexBaseHtmlPath);
                            }
                            catch (Exception)
                            {
                                //System.Diagnostics.Debug.WriteLine($"indexBase.html削除エラー: {ex.Message}");
                            }
                        }

                        application.DocumentChange += new ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
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
                application.DocumentChange += new ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            }
        }
    }
}