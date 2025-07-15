using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Packaging;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using Microsoft.Office.Interop.Word;
using OpenXmlPowerTools;
using Table = Microsoft.Office.Interop.Word.Table;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private (bool coverExist,
            string subTitle,
            int biCount,
            List<List<string>> productSubLogoGroups,
            string docTitle,
            string docid,
            bool isTmpDot,
            string trademarkTitle,
            List<string> trademarkTextList,
            string trademarkRight)
            CollectInfo(Document docCopy,
            Word.Application application,
            (string rootPath, string docName, string docFullName, string exportDir, string headerDir, string exportDirPath, string logPath, string tmpHtmlPath, string indexHtmlPath, string tmpFolderForImagesSavedBySaveAs2Method, string docid, string docTitle, string zipDirPath) paths,bool isPattern1, bool isPattern2, StreamWriter log)
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

            // HTML保存時のエンコーディング設定
            docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;

            // 表紙情報（タイトル・サブタイトル等）の収集
            CollectCoverParagraphs(docCopy, ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter, ref coverExist);
            
            // 商標・著作権情報の収集
            CollectTrademarkAndCopyrightDetails(docCopy, lastSectionIdx, log, ref trademarkTitle, ref trademarkTextList, ref trademarkRight);
            
            // タイトル・サブタイトル等の整形
            CleanUpManualTitles(ref manualTitle, ref manualSubTitle, ref manualVersion, ref manualTitleCenter, ref manualSubTitleCenter, ref manualVersionCenter);
            
            List<List<string>> productSubLogoGroups = new List<List<string>>();

            //if (coverExist)
            //{
            //    // 表紙画像やロゴの処理
            //    ProcessCoverImages(docCopy, application, paths.rootPath, paths.exportDir, ref subTitle, ref biCount, ref productSubLogoGroups, isPattern1, isPattern2, log);
            //}


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

            // manualTitle, manualTitleCenter などが空でなければそれを優先
            string resolvedDocTitle = !string.IsNullOrEmpty(manualTitle) ? manualTitle
                : !string.IsNullOrEmpty(manualTitleCenter) ? manualTitleCenter
                : paths.docTitle;

            return (coverExist, subTitle, biCount, productSubLogoGroups, resolvedDocTitle, paths.docid, isTmpDot, trademarkTitle, trademarkTextList, trademarkRight);
        }

        // 表紙に関連する段落を収集
        public void CollectCoverParagraphs(Document docCopy, ref string manualTitle, ref string manualSubTitle, ref string manualVersion,
                                      ref string manualTitleCenter, ref string manualSubTitleCenter, ref string manualVersionCenter,
                                      ref bool coverExist)
        {
            foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
            {
                string styleName = wp.get_Style().NameLocal;
                string text = wp.Range.Text.Trim();
                if (string.IsNullOrEmpty(text) || text == "/")
                    continue;

                switch (styleName)
                {
                    case "MJS_マニュアルタイトル":
                        manualTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル":
                        manualSubTitle += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン":
                        manualVersion += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルタイトル（中央）":
                        manualTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルサブタイトル（中央）":
                        manualSubTitleCenter += text + "<br/>";
                        coverExist = true;
                        break;
                    case "MJS_マニュアルバージョン（中央）":
                        manualVersionCenter += text + "<br/>";
                        coverExist = true;
                        break;
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
                    //log.WriteLine($"[Style: {wpStyleName}] {wpTextTrim}");

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
