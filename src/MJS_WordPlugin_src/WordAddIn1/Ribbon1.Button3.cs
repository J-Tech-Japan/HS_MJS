﻿using System;
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
            //Thread m_nowLoading = null;
            //m_nowLoading = new Thread(new ThreadStart(NowLoadingProc));
            //m_nowLoading.IsBackground = true;
            //m_nowLoading.Start();

            loader load = new loader();
            load.Show();

            StreamWriter sw;
            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
            button3.Enabled = false;
            WordAddIn1.Globals.ThisAddIn.Application.DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);

            ///////////////////////////////////////////////////////////////////////////
            //ここにHTML出力のコードを配置。HTMLはActiveDocumentと同階層に出力する想定でいます。
            //
            //
            ///////////////////////////////////////////////////////////////////////////

            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            Word.WdViewType defaultView = WordAddIn1.Globals.ThisAddIn.Application.ActiveWindow.View.Type;

            if (!Regex.IsMatch(thisDocument.Name, @"^[A-Z]{3}(_[^_]*?){2}\.docx*$"))
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」", "ファイル命名規則エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // read word properties
            Microsoft.Office.Core.DocumentProperties properties;
            properties = (Microsoft.Office.Core.DocumentProperties)thisDocument.CustomDocumentProperties;
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
                    thisDocument.Path + "\\headerFile\\" + Regex.Replace(thisDocument.Name, "^(.{3}).+$", "$1") + @".txt", System.Text.Encoding.Default))
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

            CoverSelectionForm coverSelectionForm = new CoverSelectionForm();
            load.Visible = false;
            coverSelectionForm.ShowDialog();
            bool isEasyCloud = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EasyCloud;
            bool isEdgeTracker = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EdgeTracker;
            bool isPattern1 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern1;
            bool isPattern2 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern2;
            bool isPattern3 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern3;

            if (coverSelectionForm.DialogResult != DialogResult.OK)
            {
                load.Close();
                load.Dispose();
                return;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.None)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("表紙のパターンをを選択してください。");
                return;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern3)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("[汎用パターン3]テンプレートはまもなく登場します。");
                return;
            }

            load.Visible = true;
            // SOURCELINK追加==========================================================================END

            thisDocument.AcceptAllRevisions();

            string rootPath = thisDocument.Path;
            string docName = thisDocument.Name;
            string docFullName = thisDocument.FullName;
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

                    //書誌情報出力
                    //スタイルチェックでもやるので、たぶん必要ない
                    //if (!makeBookInfo(log))
                    //{
                    //    load.Close();
                    //    load.Dispose();
                    //    return;
                    //}
                    Application.DoEvents();

                    log.WriteLine("HTML保存");
                    Application.DoEvents();
                    Clipboard.Clear();
                    Clipboard.SetDataObject(new DataObject());
                    Application.DoEvents();
                    WordAddIn1.Globals.ThisAddIn.Application.Selection.WholeStory();
                    WordAddIn1.Globals.ThisAddIn.Application.Selection.Copy();
                    Application.DoEvents();
                    WordAddIn1.Globals.ThisAddIn.Application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);

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
                    Word.Document docCopy = WordAddIn1.Globals.ThisAddIn.Application.Documents.Add();

                    Application.DoEvents();
                    docCopy.SaveAs2(rootPath + "\\tmp.doc");
                    docCopy.TrackRevisions = false;
                    docCopy.AcceptAllRevisions();

                    docCopy.Select();
                    Application.DoEvents();
                    WordAddIn1.Globals.ThisAddIn.Application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
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
                    string strOutFileName = "";

                    log.WriteLine("Number of sections: " + docCopy.Sections.Count);
                    int lastSectionIdx = docCopy.Sections.Count;

                    foreach (Word.Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
                    {
                        if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualTitle += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                        else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualSubTitle += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                        else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualVersion += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                        else if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル（中央）")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualTitleCenter += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                        else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル（中央）")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualSubTitleCenter += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                        else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン（中央）")
                        {
                            if (!String.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                            {
                                manualVersionCenter += wp.Range.Text + "<br/>";
                                coverExist = true;
                            }
                            continue;
                        }
                    }

                    bool isTradeMarksDetected = false;
                    bool isRightDetected = false;
                    foreach (Word.Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
                    {
                        log.WriteLine(wp.Range.Text);

                        string wpTextTrim = wp.Range.Text.Trim();
                        string wpStyleName = wp.get_Style().NameLocal;

                        if (string.IsNullOrEmpty(wpTextTrim) || wpTextTrim == "/")
                        {
                            continue;
                        }

                        if (wpTextTrim.Contains("商標")
                            && (wpStyleName.Contains("MJS_見出し 4") || wpStyleName.Contains("MJS_見出し 5")))
                        {
                            trademarkTitle = wp.Range.Text + "<br/>";
                            isTradeMarksDetected = true;
                        }
                        else if (isTradeMarksDetected && (!isRightDetected)
                            && (wpStyleName.Contains("MJS_箇条書き")
                                || wpStyleName.Contains("MJS_箇条書き2")))
                        {
                            trademarkTextList.Add(wp.Range.Text + "<br/>");
                        }
                        else if (wpTextTrim.Contains("All rights reserved")
                            && (wpStyleName.Contains("MJS_リード文")))
                        {
                            trademarkRight = wp.Range.Text + "<br/>";
                            isRightDetected = true;
                        }
                    }

                    //Word.ParagraphFormat wf = new Word.ParagraphFormat();
                    //Word.Style trademarkTitleStyle = docCopy.Styles["MJS_商標タイトル"];

                    //Word.Range lastSectionRange = docCopy.Sections[lastSectionIdx].Range;
                    //Word.Find lastSectionFind = lastSectionRange.Find;
                    //lastSectionFind.ClearFormatting();
                    //lastSectionFind.Forward = true;
                    //lastSectionFind.Format = true;
                    //lastSectionFind.set_Style("MJS_商標タイトル");

                    //object findText = "";
                    //object matchCase = false;
                    //object matchWholeWord = true;
                    //object matchWildCards = false;
                    //object matchSoundsLike = false;
                    //object matchAllWordForms = false;
                    //object forward = true;
                    //object format = true;
                    //object matchKashida = false;
                    //object matchDiacritics = false;
                    //object matchAlefHamza = false;
                    //object matchControl = false;
                    //object read_only = false;
                    //object visible = true;
                    //object replaceWith = "";
                    //object replace = 0;
                    //object wrap = 1;

                    //lastSectionFind.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                    //    ref matchSoundsLike, ref matchAllWordForms, ref forward, ref wrap, ref format, ref replaceWith,
                    //    ref replace, ref matchKashida, ref matchAlefHamza, ref matchControl);

                    //if (lastSectionFind.Found)
                    //{
                    //}

                    //foreach (Word.Table wt in docCopy.Sections[1].Range.Tables)
                    //{
                    //    foreach (Word.Column wc in wt.Columns)
                    //    {
                    //        foreach (Word.Cell wcell in wc.Cells)
                    //        {
                    //            //MessageBox.Show(wcell.Range.get_Style().NameLocal);
                    //            if (wcell.Range.get_Style().NameLocal.Trim() == "MJS_マニュアルタイトル")
                    //            {
                    //                if(!String.IsNullOrEmpty(wcell.Range.Text.Trim()) && wcell.Range.Text.Trim() != "/")
                    //                {
                    //                    manualTitle += wcell.Range.Text + "<br/>";
                    //                    coverExist = true;
                    //                }
                    //                continue;
                    //            }
                    //            else if (wcell.Range.get_Style().NameLocal.Trim() == "MJS_マニュアルサブタイトル")
                    //            {
                    //                if (!String.IsNullOrEmpty(wcell.Range.Text.Trim()) && wcell.Range.Text.Trim() != "/")
                    //                {
                    //                    manualSubTitle += wcell.Range.Text + "<br/>";
                    //                    coverExist = true;
                    //                }
                    //                continue;
                    //            }
                    //            else if (wcell.Range.get_Style().NameLocal.Trim() == "MJS_マニュアルバージョン")
                    //            {
                    //                if (!String.IsNullOrEmpty(wcell.Range.Text.Trim()) && wcell.Range.Text.Trim() != "/")
                    //                {
                    //                    manualVersion += wcell.Range.Text + "<br/>";
                    //                    coverExist = true;
                    //                }
                    //                continue;
                    //            }
                    //        }
                    //    }
                    //}

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
                        if (!Directory.Exists(rootPath + "\\tmpcoverpic")) Directory.CreateDirectory(rootPath + "\\tmpcoverpic");
                        strOutFileName = rootPath + "\\tmpcoverpic";

                        try
                        {
                            bool repeatUngroup = true;
                            while (repeatUngroup)
                            {
                                repeatUngroup = false;
                                foreach (Word.Shape ws in docCopy.Shapes)
                                {
                                    ws.Select();
                                    if (WordAddIn1.Globals.ThisAddIn.Application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1)
                                    {
                                        if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                                        {
                                            ws.Ungroup();
                                            repeatUngroup = true;
                                        }
                                    }
                                }
                            }


                            foreach (Word.Shape ws in docCopy.Shapes)
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
                                            if (!Directory.Exists(rootPath + "\\tmpcoverpic")) Directory.CreateDirectory(rootPath + "\\tmpcoverpic");

                                            strOutFileName = rootPath + "\\tmpcoverpic";
                                            byte[] vData = (byte[])WordAddIn1.Globals.ThisAddIn.Application.Selection.EnhMetaFileBits;
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

                            foreach (Word.Shape ws in docCopy.Shapes)
                            {
                                ws.Select();
                                if (WordAddIn1.Globals.ThisAddIn.Application.Selection.Information[Word.WdInformation.wdActiveEndSectionNumber] == 1)
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

                                foreach (Word.Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
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

                                            foreach (Word.InlineShape wis in wp.Range.InlineShapes)
                                            {
                                                wis.Range.Select();
                                                Clipboard.Clear();
                                                WordAddIn1.Globals.ThisAddIn.Application.Selection.CopyAsPicture();
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
                                    string destF = rootPath + "\\" + exportDir + "\\template\\images\\" + Path.GetFileName(pairs[p].Key);

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
                                        if (File.Exists(rootPath + "\\" + exportDir + "\\template\\images\\cover-4.png")) File.Delete(rootPath + "\\" + exportDir + "\\template\\images\\cover-4.png");
                                        File.Move(pairs[p].Key, rootPath + "\\" + exportDir + "\\template\\images\\cover-4.png");
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
                                            dst.Save(rootPath + "\\" + exportDir + "\\template\\images\\cover-background.png", ImageFormat.Png);
                                        }
                                        // Saves result.
                                        File.Delete(pairs[p].Key);
                                    }
                                }
                            }

                            if (Directory.Exists(rootPath + "\\tmpcoverpic")) Directory.Delete(rootPath + "\\tmpcoverpic", true);
                        }
                        catch (Exception ex)
                        {
                            log.WriteLine(ex.ToString());
                        }
                    }

                    WordAddIn1.Globals.ThisAddIn.Application.Selection.EndKey(Word.WdUnits.wdStory);
                    object selectionRange = WordAddIn1.Globals.ThisAddIn.Application.Selection.Range;
                    Word.Shape wst = docCopy.Shapes.AddCanvas(0, 0, 1, 1, ref selectionRange);
                    wst.WrapFormat.Type = Word.WdWrapType.wdWrapInline;

                    foreach (Word.Shape docS in docCopy.Shapes)
                    {
                        if (docS.Type == Microsoft.Office.Core.MsoShapeType.msoCanvas)
                        {
                            List<float> canvasItemsTop = new List<float>();
                            List<float> canvasItemsLeft = new List<float>();
                            List<float> canvasItemsHeight = new List<float>();
                            List<float> canvasItemsWidth = new List<float>();

                            for (int i = 1; i <= docS.CanvasItems.Count; i++)
                            {
                                docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                                canvasItemsTop.Add(docS.CanvasItems[i].Top);
                                canvasItemsLeft.Add(docS.CanvasItems[i].Left);
                                canvasItemsHeight.Add(docS.CanvasItems[i].Height);
                                canvasItemsWidth.Add(docS.CanvasItems[i].Width);
                            }
                            //float canvasWidth = docS.Width;
                            docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                            //docS.Height = docS.Height + 15;
                            docS.Height = docS.Height + 30;
                            docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                            for (int i = 1; i <= docS.CanvasItems.Count; i++)
                            {
                                docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
                                docS.CanvasItems[i].Height = canvasItemsHeight[i - 1];
                                docS.CanvasItems[i].Width = canvasItemsWidth[i - 1];
                                docS.CanvasItems[i].Top = canvasItemsTop[i - 1] + 0.59F;
                                docS.CanvasItems[i].Left = canvasItemsLeft[i - 1];
                                docS.CanvasItems[i].LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoTrue;
                            }
                        }
                    }
                    wst.Delete();

                    foreach (Word.Table wt in docCopy.Tables)
                    {
                        if (wt.PreferredWidthType == Word.WdPreferredWidthType.wdPreferredWidthPoints)
                            wt.AllowAutoFit = true;
                    }
                    //                    File.Copy(docFullName, rootPath + "\\tmp.doc", true);

                    //                  Word.Document docCopy = WordAddIn1.Globals.ThisAddIn.Application.Documents.Open(rootPath + "\\tmp.doc");

                    foreach (Word.Style ws in docCopy.Styles)
                        if (ws.NameLocal == "奥付タイトル")
                            ws.NameLocal = "titledef";

                    docCopy.WebOptions.Encoding = Microsoft.Office.Core.MsoEncoding.msoEncodingUTF8;
                    docCopy.SaveAs2(rootPath + "\\tmp.html", Word.WdSaveFormat.wdFormatFilteredHTML);
                    docCopy.Close();
                    File.Delete(rootPath + "\\tmp.doc");

                    log.WriteLine("画像フォルダ コピー");

                    //if (Directory.Exists(rootPath + "\\" + exportDir + "\\pict"))
                    //{
                    //    Directory.Delete(rootPath + "\\" + exportDir + "\\pict");
                    //}
                    //Directory.Move(rootPath + "\\tmp.files", rootPath + "\\" + exportDir + "\\pict");

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

                    htmlStr = Regex.Replace(htmlStr, "\r\n", " ");
                    htmlStr = Regex.Replace(htmlStr, "<meta[^>]*?>", "");
                    htmlStr = Regex.Replace(htmlStr, "(<(?:input|br|img)[^>]*)>", "$1/>");
                    htmlStr = Regex.Replace(htmlStr, "<span [^>]+>(?:&nbsp;)+ </span>", "　");
                    htmlStr = Regex.Replace(htmlStr, "&nbsp;", ((char)160).ToString());
                    htmlStr = Regex.Replace(htmlStr, "&copy;", ((char)169).ToString());
                    while (Regex.IsMatch(htmlStr, @"(src\s*=\s*""[^""]*?)\\([^""]*?"")"))
                        htmlStr = Regex.Replace(htmlStr, @"(src\s*=\s*""[^""]*?)\\([^""]*?"")", "$1/$2");

                    while (Regex.IsMatch(htmlStr, @"(<[A-z]+[^>]* [A-z-]+=)([^""'][^ />]*)"))
                    {
                        htmlStr = Regex.Replace(htmlStr, @"(<[A-z]+[^>]* [A-z-]+=)([^""'][^ />]*)", @"$1""$2""");
                    }

                    if (isTmpDot)
                    {
                        htmlStr = Regex.Replace(htmlStr, @"src=""tmp\.files/", @"src=""pict/");
                    }
                    else
                    {
                        htmlStr = Regex.Replace(htmlStr, @"src=""tmp_files/", @"src=""pict/");
                    }
                    htmlStr = Regex.Replace(htmlStr, @"<a name=""_Toc\d+?""></a>", "");
                    htmlStr = Regex.Replace(htmlStr, @"<span lang=""[^""]+?"">([^<]+?)</span>", "$1");
                    htmlStr = Regex.Replace(htmlStr, @"(<hr(?: [^/>]*)?)(>)", "$1/$2");
                    htmlStr = Regex.Replace(htmlStr, @"z-index:-?\d{3,};", "$1");
                    htmlStr = Regex.Replace(htmlStr, @"(?<=<[^>]+?) style=['""]?[^'"" ]+['""]?( (?:[^>]*)style=['""]?[^'"" >/]+['""]?)", "$1");
                    htmlStr = Regex.Replace(htmlStr, @"(<p[^>]*?(?<!/)>)([^<]*)(</(?!p))", @"$1$2</p>$3");
                    htmlStr = htmlStr.Replace("MJS--", "MJSTT");

                    System.Xml.XmlDocument objXml = new System.Xml.XmlDocument();

                    objXml.LoadXml(htmlStr);

                    foreach (System.Xml.XmlElement imgNode in objXml.SelectNodes("//img"))
                    {
                        imgNode.RemoveAttribute("height");
                        imgNode.RemoveAttribute("width");
                    }
                    foreach (System.Xml.XmlElement pageBreak in objXml.SelectNodes("//span[(translate(., ' &#10;&#13;&#9;', '') = '') and (count(*) = 1) and boolean(br[@style = 'page-break-before:always'])]"))
                    {
                        pageBreak.ParentNode.RemoveChild(pageBreak);
                    }
                    foreach (System.Xml.XmlElement pageBreak in objXml.SelectNodes("//br[translate(@style, ' &#10;&#13;&#9;', '') = 'page-break-before:always']"))
                    {
                        pageBreak.ParentNode.RemoveChild(pageBreak);
                    }
                    foreach (System.Xml.XmlElement comment in objXml.SelectNodes("//*[boolean(./*/@class[starts-with(., 'msocom')])]"))
                    {
                        comment.ParentNode.RemoveChild(comment);
                    }
                    foreach (System.Xml.XmlElement link in objXml.SelectNodes("//a[boolean(@href)]"))
                    {
                        if (link.InnerText.Contains("http")) continue;
                        //if (Regex.IsMatch(link.GetAttribute("href"), @"^.*?#.*?\.html$"))
                        //    link.GetAttributeNode("href").Value = Regex.Replace(link.GetAttribute("href"), @"^(.*?)#(.*?)\.html$", "$1.html#$2");
                        link.InnerText = Regex.Replace(link.InnerText, @"^(.*?)(?=[\s　](\d+\.\d+|[^\s|　]*?章))", "");
                        //link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:[^\s|　]*?編)*[\s　]+", "");
                        link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:第[\d０-９]+章)*[\s　]+", "");
                        link.InnerText = Regex.Replace(link.InnerText, @"^[\s　]*(?:\d+\.)*\d+[\s　]+", "");
                    }
                    System.Xml.XmlDocument objToc = new System.Xml.XmlDocument();
                    objToc.LoadXml(@"<result><item title=""" + docTitle + @"""></item></result>");
                    System.Xml.XmlNode objTocCurrent = objToc.DocumentElement;

                    System.Xml.XmlDocument objBody = new System.Xml.XmlDocument();
                    objBody.LoadXml("<result></result>");
                    System.Xml.XmlNode objBodyCurrent = objBody.DocumentElement;

                    string className = "";
                    className = objXml.SelectSingleNode("/html/head/style[contains(comment(), 'mso-style-name')]").OuterXml;
                    className = Regex.Replace(className, "[\r\n\t ]+", "");
                    className = Regex.Replace(className, "}", "}\n");

                    Dictionary<string, string> styleName = new Dictionary<string, string>();

                    string chapterSplitClass = "";

                    foreach (string clsName in className.Split('\n'))
                    {
                        string clsBefore, clsAfter;

                        if (Regex.IsMatch(clsName, "mso-style-name:"))
                        {
                            clsBefore = Regex.Replace(clsName, "^(.+?){.+?}$", "$1");
                            clsAfter = Regex.Replace(clsName, @"^.+?{.*mso-style-name:""(.+?)\\,.*"";.*}", "$1");
                            clsAfter = Regex.Replace(clsAfter, "^.+?{.*mso-style-name:(.+?);.*}", "$1");

                            foreach (string cls in clsBefore.Split(','))
                            {
                                if (Regex.IsMatch(clsAfter, "章[　 ]*扉.*タイトル"))
                                {
                                    if (chapterSplitClass != "")
                                    {
                                        chapterSplitClass += "|";
                                    }
                                    chapterSplitClass += Regex.Replace(cls, @"^(.+?)\.(.+?)$", "$1[@class='$2']");
                                }

                                styleName.Add(cls, Regex.Replace(clsAfter, @"\\", ""));
                            }
                        }
                        else if (Regex.IsMatch(clsName, "mso-style-link:"))
                        {
                            clsBefore = Regex.Replace(clsName, "^(.+?){.+?}$", "$1");
                            clsAfter = Regex.Replace(clsName, @"^.+?{.*mso-style-link:""(.+?)\\,.*"";.*}", "$1");
                            clsAfter = Regex.Replace(clsAfter, "^.+?{.*mso-style-link:(.+?);.*}", "$1");

                            foreach (string cls in clsBefore.Split(','))
                            {
                                if (Regex.IsMatch(clsAfter, "章[　 ]*扉.*タイトル"))
                                {
                                    if (chapterSplitClass != "")
                                    {
                                        chapterSplitClass += "|";
                                    }
                                    chapterSplitClass += Regex.Replace(cls, @"^(.+?)\.(.+?)$", "$1[@class='$2']");
                                }

                                styleName.Add(cls, Regex.Replace(clsAfter, @"\\", ""));
                            }
                        }
                    }
                    log.WriteLine("index.html出力");

                    List<string> titleDeffenition = new List<string>();
                    foreach (System.Xml.XmlElement link in objXml.SelectNodes("//p[@class='titledef']"))
                    {
                        titleDeffenition.Add(link.InnerText.Trim());
                    }

                    //if (!isEdgeTracker && titleDeffenition.Contains("Edge Tracker"))
                    //{
                    //    isEdgeTracker = true;
                    //}
                    #region
                    string idxHtmlTemplate = "";
                    idxHtmlTemplate += @"<?xml version=""1.0"" encoding=""utf-8"" ?>" + "\n";
                    idxHtmlTemplate += @"<!DOCTYPE HTML>" + "\n";
                    idxHtmlTemplate += @"<html>" + "\n";
                    idxHtmlTemplate += @"<head>" + "\n";
                    idxHtmlTemplate += @" <meta http-equiv=""X-UA-Compatible"" content=""IE=edge"" />" + "\n";
                    idxHtmlTemplate += @"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" + "\n";
                    idxHtmlTemplate += @"<meta name=""viewport"" content=""width=device-width, initial-scale=1, maximum-scale=3, user-scalable=yes"" />" + "\n";
                    idxHtmlTemplate += @"<meta name=""generator"" content=""Adobe Framemaker 2017"" />" + "\n";
                    idxHtmlTemplate += @"<title>" + docTitle + "</title>" + "\n";
                    idxHtmlTemplate += @"<link rel=""StyleSheet"" href=""template/Azure_Blue01/layout.css"" type=""text/css"" />" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/rh.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/common.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/layout.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/constants.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/projectdata.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/utils.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/mhutils.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/mhlang.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/mhver.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/settings.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/XmlJsReader.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/loadscreen.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/loadcsh.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/loadparentdata.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/loadprojdata.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/showhidecontrols.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/pageloader.js""></script>" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/widgets.min.css"">" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/layoutfix.min.css"">" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/layout.min.css"">" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/font.css"">" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/resp.css"">" + "\n";
                    idxHtmlTemplate += @"<link rel=""stylesheet"" type=""text/css"" href=""template/styles/pdf.css"" >" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/mhfhost.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""template/scripts/search.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""promise-6.1.0.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""jquery-3.1.0.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""html2canvas.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""jspdf.min.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""jquery.cookie.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""resp.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""fontchange.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""pdf.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" src=""search.js""></script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" >" + "\n";
                    idxHtmlTemplate += @"gTopicFrameName = ""rh_default_topic_frame_name"";" + "\n";
                    idxHtmlTemplate += @"gDefaultTopic = ""#" + docid + @"00000.html"";" + "\n";
                    idxHtmlTemplate += @"</script>" + "\n";
                    idxHtmlTemplate += @"<script type=""text/javascript"" >" + "\n";
                    idxHtmlTemplate += @"gRootRelPath = ""."";" + "\n";
                    idxHtmlTemplate += @"gCommonRootRelPath = ""."";" + "\n";
                    idxHtmlTemplate += @"mergePage = {" + "\n";
                    foreach (var item in mergeScript)
                    {
                        idxHtmlTemplate += item.Value.Split(new char[] { '♯' })[0] + ":'" + item.Key.Split(new char[] { '♯' })[0] + "',";
                    }
                    idxHtmlTemplate += @"};" + "\n";
                    idxHtmlTemplate += @"function findFirstPageInMerge(page) {" + "\n";
                    idxHtmlTemplate += @"    var found = false;" + "\n";
                    idxHtmlTemplate += @"    for (let key in mergePage) {" + "\n";
                    idxHtmlTemplate += @"        if (mergePage[key] == page) {" + "\n";
                    idxHtmlTemplate += @"            found = true;" + "\n";
                    idxHtmlTemplate += @"            return findFirstPageInMerge(key);" + "\n";
                    idxHtmlTemplate += @"            break;" + "\n";
                    idxHtmlTemplate += @"        }" + "\n";
                    idxHtmlTemplate += @"    }" + "\n";
                    idxHtmlTemplate += @"    if (!found) {" + "\n";
                    idxHtmlTemplate += @"        return page;" + "\n";
                    idxHtmlTemplate += @"    }" + "\n";
                    idxHtmlTemplate += @"}" + "\n";
                    idxHtmlTemplate += @"$(function() {" + "\n";
                    idxHtmlTemplate += @"    updateLink();" + "\n";
                    idxHtmlTemplate += @"});" + "\n";
                    idxHtmlTemplate += @"function updateLink() {" + "\n";
                    idxHtmlTemplate += @"    $(""ul.toc a, .wSearchResultItemsBlock a.nolink"").each(function() {" + "\n";
                    idxHtmlTemplate += @"        var href = $(this).attr(""href"");" + "\n";
                    idxHtmlTemplate += @"        if (href.match(""([A-Z]{3}[0-9]{5}[.]html)[?]rhtocid=.+[#]([A-Z]{3}[0-9]{5})"")) {" + "\n";
                    idxHtmlTemplate += @"            var rex = /([A-Z]{3}[0-9]{5}[.]html)[?]rhtocid=.+[#]([A-Z]{3}[0-9]{5})/g;" + "\n";
                    idxHtmlTemplate += @"            var arr = rex.exec(href);" + "\n";
                    idxHtmlTemplate += @"            $(this).attr(""href"", findFirstPageInMerge(arr[1].replace("".html"","""")) + "".html#"" + arr[2]);" + "\n";
                    idxHtmlTemplate += @"        }else if($(this).hasClass(""nolink"")){" + "\n";
                    idxHtmlTemplate += @"            if(!$(this).hasClass(""changed"")){" + "\n";
                    idxHtmlTemplate += @"                var rex = /([A-Z]{3}[0-9]{5}[.]html)/g;" + "\n";
                    idxHtmlTemplate += @"                var arr = rex.exec(href);" + "\n";
                    idxHtmlTemplate += @"                $(this).addClass(""changed"");" + "\n";
                    idxHtmlTemplate += @"                var from = findFirstPageInMerge(arr[1].replace("".html"",""""));" + "\n";
                    idxHtmlTemplate += @"                var to=arr[1].replace("".html"","""");" + "\n";
                    idxHtmlTemplate += @"                if(from==to){" + "\n";
                    idxHtmlTemplate += @"                    to="""";" + "\n";
                    idxHtmlTemplate += @"                }" + "\n";
                    idxHtmlTemplate += @"                $(this).attr(""href"", from + "".html#"" + to);" + "\n";
                    idxHtmlTemplate += @"            }" + "\n";
                    idxHtmlTemplate += @"        }" + "\n";
                    idxHtmlTemplate += @"    });" + "\n";
                    idxHtmlTemplate += @"    setTimeout(function() {" + "\n";
                    idxHtmlTemplate += @"        updateLink();" + "\n";
                    idxHtmlTemplate += @"    }, 200);" + "\n";
                    idxHtmlTemplate += @"}" + "\n";
                    idxHtmlTemplate += @"</script>" + "\n";
                    idxHtmlTemplate += @"</head>" + "\n";
                    idxHtmlTemplate += @"<body class=""hide-children loading"" data-rhwidget=""Basic"" data-class=""media-desktop: KEY_SCREEN_DESKTOP; media-landscape: KEY_SCREEN_TABLET; media-mobile: KEY_SCREEN_PHONE; ios: KEY_SCREEN_IOS"" data-controller=""ModernLayoutController: mc; JsLoadingController"" data-attr=""dir:KEY_DIR"">" + "\n";
                    //idxHtmlTemplate += @"<script type=""text/javascript"" src=""ehlpdhtm.js""></script>" + "\n";
                    idxHtmlTemplate += @"<!-- Extra mobile header with special functions -->" + "\n";
                    idxHtmlTemplate += @"<div class=""mobilespecialfunctions"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); mobile-header-visible: @.l.mobile_header_visible; searchpage-mode: @KEY_ACTIVE_TAB == 'fts'"">" + "\n";
                    idxHtmlTemplate += @"   <a class=""menubutton"" data-attr=""href: '#'; title:@KEY_LNG.NavTip;"" data-click=""$mc.toggleSideBar()"" data-if=""@.l.mobile_menu_enabled === true""></a> " + "\n";
                    idxHtmlTemplate += @"   <a class=""fts"" data-attr=""href: '#'; title:@KEY_LNG.SearchTitle;"" data-click=""$mc.toggleActiveTab('fts')"">&#160;</a>" + "\n";
                    idxHtmlTemplate += @"   <a class=""filter"" data-attr=""href: '#'; title:@KEY_LNG.Filter"" data-if=""KEY_FEATURE.filter"" data-class=""filter-applied: @.l.tag_expression.length""  data-click=""$mc.toggleActiveTab('filter')"">&#160;</a>" + "\n";
                    idxHtmlTemplate += @"   <div class=""brs-holder"">" + "\n";
                    idxHtmlTemplate += @"     <div class=""brs_previous"" data-if=""@active_content != 'search' && @KEY_SEARCH_LOCATION == 'content'"" data-attr=""title:@KEY_LNG.Prev""><a id=""browseSeqBackMobile"" class=""wBSBackButton"" data-rhwidget=""Basic"" data-attr=""href:.l.brsBack"" data-css=""visibility: @.l.brsBack?'visible':'hidden'"">&nbsp;</a>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"     <div class=""brs_next"" data-if=""@active_content != 'search' && @KEY_SEARCH_LOCATION == 'content'"" data-attr=""title:@KEY_LNG.Next""><a id=""browseSeqNextMobile"" class=""wBSNextButton"" data-rhwidget=""Basic"" data-attr=""href:.l.brsNext"" data-css=""visibility: @.l.brsNext?'visible':'hidden'"">&nbsp;</a>" + "\n";
                    idxHtmlTemplate += @"</div>   " + "\n";
                    idxHtmlTemplate += @"   </div>  " + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Function bar with TOC/IDS/GLO/FILTER/FTS buttons -->" + "\n";
                    idxHtmlTemplate += @"<div class=""functionbar"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; mobile-header-visible: @.l.mobile_header_visible"">" + "\n";
                    idxHtmlTemplate += @"   <div class=""nav"">" + "\n";
                    idxHtmlTemplate += @"       <a class=""toc"" data-if=""KEY_FEATURE.toc"" data-class=""active: @KEY_ACTIVE_TAB == 'toc'"" data-click=""$mc.toggleActiveTab('toc')"" data-attr=""title:@KEY_LNG.TableOfContents; href: '#'"">&#160;</a>" + "\n";
                    idxHtmlTemplate += @"       <a class=""filter"" data-if=""KEY_FEATURE.filter"" data-class=""active: @KEY_ACTIVE_TAB == 'filter'; filter-applied: @.l.tag_expression.length""  data-click=""$mc.toggleActiveTab('filter')"" data-attr=""title:@KEY_LNG.Filter; href: '#'"">&#160;</a>" + "\n";
                    idxHtmlTemplate += @"       <a class=""fts"" data-if=""@KEY_SEARCH_LOCATION == 'tabbar'"" data-class=""active: @KEY_ACTIVE_TAB == 'fts'; search-sidebar: @KEY_SEARCH_LOCATION == 'tabbar'"" data-click=""$mc.toggleActiveTab('fts')"" data-attr=""title:@KEY_LNG.SearchTitle; href: '#'"">&#160;</a>" + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Table of contents -->" + "\n";
                    idxHtmlTemplate += @"<div class=""toc-holder left-pane"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; search-sidebar: @KEY_SEARCH_LOCATION == 'tabbar'; search-content: @KEY_SEARCH_LOCATION == 'content'; layout-visible: @KEY_ACTIVE_TAB == 'toc'; drill-down: KEY_TOC_DRILL_DOWN; mobile-header-visible: @.l.mobile_header_visible; loading: !@EVT_TOC_LOADED"">" + "\n";
                    idxHtmlTemplate += @"   <ul class=""toc"" data-rhwidget=""List: .p.toc"" data-controller=""TocController: toc"" data-click=""$toc.onClick(event)"">" + "\n";
                    idxHtmlTemplate += @"       <li data-rif=""item.type === 'item' || item.type === 'remoteitem'"" data-i-class=""$toc.class(item)"" data-class=""inactive: @bookid != '#{@pid}'"" data-childorder=""#{childOrder}"" data-rhtags=""#{$toc.tags(item)}"" data-itemid=""#{@id}"">" + "\n";
                    idxHtmlTemplate += @"           <a data-itext=""item.name"" data-i-href=""$toc.url(item, '#{@id}')""></a>" + "\n";
                    idxHtmlTemplate += @"       </li>" + "\n";
                    idxHtmlTemplate += @"       <li class=""book"" data-rif=""item.type === 'book'"" data-class=""active: @bookid == '#{@id}'; inactive: @bookid != '#{@pid}' &amp;&amp; @bookid != '#{@id}'"" data-childorder=""#{childOrder}""" + "\n";
                    idxHtmlTemplate += @"           data-itemkey=""#{$toc.key(item.absRef, item.key)}"" data-itemid=""#{@id}"" data-itemlevel=""#{@level}"" data-rhtags=""#{$toc.tags(item)}"">" + "\n";
                    idxHtmlTemplate += @"           <a data-itext=""item.name"" data-i-href=""$toc.url(item, '#{@id}')""></a>" + "\n";
                    idxHtmlTemplate += @"       </li>" + "\n";
                    idxHtmlTemplate += @"       <li class=""child max-height-transition"" data-rif=""item.key"" data-class=""show: @show_child#{@id}"" data-childorder=""#{childOrder}"">" + "\n";
                    idxHtmlTemplate += @"         <ul class=""child"" data-child=""$toc.key(item.absRef, item.key)""></ul>" + "\n";
                    idxHtmlTemplate += @"       </li>" + "\n";
                    idxHtmlTemplate += @"   </ul>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Index -->" + "\n";
                    idxHtmlTemplate += @"<div class=""idx-holder left-pane"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; layout-visible: @KEY_ACTIVE_TAB == 'idx'; mobile-header-visible: @.l.mobile_header_visible"" data-scroll=""@.l.load_more_index(true) | debounce: 50, delta: 100"">" + "\n";
                    idxHtmlTemplate += @"   <div id=""idx"" class=""Index"">" + "\n";
                    idxHtmlTemplate += @"       <input class=""IdxFilter"" data-attr=""placeholder:@KEY_LNG.IndexFilterKewords"" type=""text"" data-keyup=""@.l.idxfilter(node.value)"" />" + "\n";
                    idxHtmlTemplate += @"        <ul class=""index-list"" data-rhwidget=""List: key: PROJECT_INDEX_DATA, , filter: $ic.showItem(item.name), spliton: index % 60 == 59"" data-config=""loadmore: '.l.load_more_index'"" data-controller=""IndexController: ic""> " + "\n";
                    idxHtmlTemplate += @"           <li class=""treeitem IndexAlphabet"" data-rif=""$ic.showCategory(item.name, this.path.length)"" data-itemlevel=""#{@level}""> " + "\n";
                    idxHtmlTemplate += @"             <span class=""IndexAlphabetText"" data-itext=""$ic.alphaText(item.name)""></span> " + "\n";
                    idxHtmlTemplate += @"           </li>" + "\n";
                    idxHtmlTemplate += @"           <li class=""treeitem IndexKeyword"" data-i-data-rhtags=""item['data-rhtags']"">" + "\n";
                    idxHtmlTemplate += @"               <a data-rif=""item.topics &amp;&amp; item.topics.length == 1"" class=""nolink IndexKeywordText"" data-i-href=""item.topics[0].url"" data-itext=""item.name""> " + "\n";
                    idxHtmlTemplate += @"               </a> " + "\n";
                    idxHtmlTemplate += @"               <span class=""IndexKeywordText IndexKeyword unselectable"" data-itext=""item.name"" data-rif=""item.topics &amp;&amp; item.topics.length != 1"" data-i-title=""item.name"" data-click=""@show.#{@id}(!@show.#{@id})""></span> " + "\n";
                    idxHtmlTemplate += @"           <ul data-if=""@show.#{@id}"" style=""list-style-type: none;"">" + "\n";
                    idxHtmlTemplate += @"            <li data-repeat=""i, topic:#{@itemkey}.topics"" data-rif=""item.topics &amp;&amp; item.topics.length > 1"" class=""IndexChildBlock IndexKeyword""> " + "\n";
                    idxHtmlTemplate += @"             <a class=""nolink IndexLink IndexLinkText"" data-i-href=""$topic.url"" data-i-data-rhtags=""$topic['data-rhtags']"" data-itext=""$topic.name"" data-i-title=""$topic.name""> " + "\n";
                    idxHtmlTemplate += @"             </a> " + "\n";
                    idxHtmlTemplate += @"            </li>" + "\n";
                    idxHtmlTemplate += @"           </ul>" + "\n";
                    idxHtmlTemplate += @"            <div class=""IndexChildBlock"" data-rif=""item['keys']""> " + "\n";
                    idxHtmlTemplate += @"             <ul class=""child"" data-child=""#{@itemkey}.keys"" style=""list-style-type: none;""></ul> " + "\n";
                    idxHtmlTemplate += @"            </div>" + "\n";
                    idxHtmlTemplate += @"          </li> " + "\n";
                    idxHtmlTemplate += @"        </ul>" + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Glossary -->" + "\n";
                    idxHtmlTemplate += @"<div class=""glo-holder left-pane"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; layout-visible: @KEY_ACTIVE_TAB == 'glo'; mobile-header-visible: @.l.mobile_header_visible"">" + "\n";
                    idxHtmlTemplate += @"   <div id=""glo"" class=""Glossary"" data-controller=""GlossaryController: gc"">" + "\n";
                    idxHtmlTemplate += @"       <input class=""GloFilter"" data-attr=""placeholder:@KEY_LNG.GlossaryFilterTerms"" type=""text"" data-keyup=""$gc.filterGlo(node.value)""/>" + "\n";
                    idxHtmlTemplate += @"           <ul style=""list-style: none;""> " + "\n";
                    idxHtmlTemplate += @"               <div data-repeat=""i, glossary: PROJECT_GLOSSARY_DATA""> " + "\n";
                    idxHtmlTemplate += @"                   <li class=""treeitem GloAlphabet"" data-rif=""!$gc.exists($glossary.name)&amp;&amp;!$gc.isFiltered($glossary.name)""> " + "\n";
                    idxHtmlTemplate += @"                       <span class=""GloAlphabetText"" data-itext=""$gc.alphaText($glossary.name)""></span>" + "\n";
                    idxHtmlTemplate += @"                   </li>" + "\n";
                    idxHtmlTemplate += @"                   <li class=""treeitem"" data-rif=""!$gc.isFiltered($glossary.name)""> " + "\n";
                    idxHtmlTemplate += @"                       <div class=""GlossTerm unselectable"" data-type=""11"" data-i-title=""$glossary.name"" data-term=""$glossary.name"" data-click=""@show.#{@index}(!@show.#{@index})""> " + "\n";
                    idxHtmlTemplate += @"                           <span class=""GlossaryTermText"" data-itext=""$glossary.name""></span>" + "\n";
                    idxHtmlTemplate += @"                       </div> " + "\n";
                    idxHtmlTemplate += @"                       <div class=""GlossDefinition unselectable"" data-type=""12"" data-if=""@show.#{@index}""> " + "\n";
                    idxHtmlTemplate += @"                           <span class=""GlossDefinitionText"" data-itext=""$glossary.value""></span> " + "\n";
                    idxHtmlTemplate += @"                       </div> " + "\n";
                    idxHtmlTemplate += @"                   </li> " + "\n";
                    idxHtmlTemplate += @"               </div> " + "\n";
                    idxHtmlTemplate += @"           </ul> " + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Filter -->" + "\n";
                    idxHtmlTemplate += @"<div class=""filter-holder left-pane"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; layout-visible: @KEY_ACTIVE_TAB == 'filter'; mobile-header-visible: @.l.mobile_header_visible; loading: !@KEY_MERGED_FILTER_KEY"">" + "\n";
                    idxHtmlTemplate += @"   <div class=""mobile-filter-heading"" data-if=""KEY_SCREEN_PHONE"">" + "\n";
                    idxHtmlTemplate += @"       <a class=""mobile_back"" data-click=""$mc.filterDone()"" data-attr=""title:@KEY_LNG.ApplyTip""></a>" + "\n";
                    idxHtmlTemplate += @"       <div class=""page-title"" data-text=""KEY_PROJECT_FILTER_CAPTION""></div>" + "\n";
                    idxHtmlTemplate += @"       <a class=""reset-button"" data-attr=""href: '#'; title: @KEY_LNG.Reset"" data-click=""$mc.setDefaultTagStates()"" data-class=""layout-visible: $mc.isTagStatesChanged(@KEY_TAG_EXPRESSION)""></a>" + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"   <p class=""filter-title"" data-if=""(@KEY_SCREEN_TABLET || @KEY_SCREEN_DESKTOP) &amp;&amp; @KEY_MERGED_FILTER_KEY"" data-text=""KEY_PROJECT_FILTER_CAPTION""></p>" + "\n";
                    idxHtmlTemplate += @"   <a class=""reset-button"" data-attr=""href: '#'; title: @KEY_LNG.Reset"" data-click=""$mc.setDefaultTagStates()"" data-class=""layout-visible: $mc.isTagStatesChanged(@KEY_TAG_EXPRESSION)""></a>" + "\n";
                    idxHtmlTemplate += @"   <ul class=""wFltOpts"" data-rhwidget=""List:KEY_MERGED_FILTER_KEY"" data-controller=""FilterController: fc"" data-click=""$fc.click(event)""" + "\n";
                    idxHtmlTemplate += @"   data-class=""radio: @KEY_PROJECT_FILTER_TYPE == 'radio'; checkbox: @KEY_PROJECT_FILTER_TYPE == 'checkbox'"">" + "\n";
                    idxHtmlTemplate += @"       <li data-i-class=""$fc.class(item)"" data-itemkey=""#{@path}"" data-itemvalue=""#{name}"">" + "\n";
                    idxHtmlTemplate += @"         <input data-rif=""$fc.inputType(item) == 'checkbox'"" data-i-id=""'filter#{@id}'"" type=""checkbox"" data-i-value=""#{@index}""/>" + "\n";
                    idxHtmlTemplate += @"         <input data-rif=""$fc.inputType(item) == 'radio'"" data-i-id=""'filter#{@id}'"" type=""radio"" data-i-name=""'filter_name#{@pid}'"" data-i-value=""'#{@index}'""/>          " + "\n";
                    idxHtmlTemplate += @"         <label data-i-for=""'filter#{@id}'"" data-i-title=""item.display"" data-itext=""item.display"" data-class=""tag-parent: #{@itemkey}.children; checked: KEY_PROJECT_TAG_STATES#{@path}""></label>" + "\n";
                    idxHtmlTemplate += @"         <ul class=""wFltOptsGrp"" data-child=""#{@itemkey}.children"" data-rif=""item.children""></ul>" + "\n";
                    idxHtmlTemplate += @"       </li>" + "\n";
                    idxHtmlTemplate += @"    </ul>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Sidebar sizer -->" + "\n";
                    idxHtmlTemplate += @"<div class=""sidebarsizer left-pane boundry column-resize"" data-if=""@.l.desktop_sidebar_available === true"" data-resize=""@.l.desktop_sidebar_visible(@sidebar_width == null || @sidebar_width != '0px') | x: 'sidebar_width', maxx: 0.7, screen: 'desktop'"" data-css=""left: sidebar_width | screen: 'desktop', dir: 'ltr'; right: sidebar_width | screen: 'desktop', dir: 'rtl'"" data-class=""desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false;"">" + "\n";
                    idxHtmlTemplate += @"   <a class=""sidebartoggle"" data-click=""@sidebar_width(null) | screen: 'desktop'"" data-toggle="".l.desktop_sidebar_visible"" data-attr=""title: @KEY_LNG.SidebarToggleTip"">&nbsp;</a>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Search field -->" + "\n";
                    idxHtmlTemplate += @"<div class=""searchbar left-pane"" data-css=""width: sidebar_width | screen: 'desktop'"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; mobile-header-visible: @.l.mobile_header_visible; searchpage-mode, layout-visible: @KEY_ACTIVE_TAB == 'fts'; search-sidebar: @KEY_SEARCH_LOCATION == 'tabbar'; searchbar-mobile:KEY_SCREEN_PHONE; search-content: @KEY_SEARCH_LOCATION == 'content';  @SEARCH_RESULTS_KEY!== undefined &&  (@SEARCH_RESULTS_KEY).length>0"" data-controller=""SearchController:sc"">" + "\n";
                    idxHtmlTemplate += @"   <a class=""mobile_back"" data-click=""@KEY_ACTIVE_TAB(null)"" data-attr=""title: @KEY_LNG.Back""></a>" + "\n";
                    idxHtmlTemplate += @"   <div class=""search-input"" data-class=""search-input-open:@focusin_main_searchbox && @SEARCH_RESULTS_KEY!== undefined &&  (@SEARCH_RESULTS_KEY).length>0"" data-if=""$mc.isSearchMode(@KEY_ACTIVE_TAB, @active_content)"">" + "\n";
                    idxHtmlTemplate += @"       <input class=""wSearchField"" type=""text"" data-class=""no-filter: !@KEY_FEATURE.filter"" data-attr=""placeholder: @KEY_LNG.Search""/>" + "\n";
                    idxHtmlTemplate += @"       <a class=""wSearchLink"" data-click=""@EVT_SEARCH_TERM(true)"" data-attr=""href: '#'"" data-if=""@KEY_SCREEN_PHONE"">&nbsp;</a>   " + "\n";
                    idxHtmlTemplate += @"       <div data-if=""@focusin_main_searchbox && @SEARCH_RESULTS_KEY!== undefined &&  (@SEARCH_RESULTS_KEY).length>0"" class=""search-list"" >" + "\n";
                    idxHtmlTemplate += @"         <table data-class=""search-table-desktop:KEY_SCREEN_DESKTOP; search-table-mobile:KEY_SCREEN_PHONE; search-table-tablet:KEY_SCREEN_TABLET"">" + "\n";
                    idxHtmlTemplate += @"           <tr data-repeat=""search_results"" data-class=""search-suggestion:true; search-selected:@selected===#{@index}"" data-click=""$sc.handleClick(#{@index})"">" + "\n";
                    idxHtmlTemplate += @"               <td class=""search-text-column""><div class=""search-text"" data-itext=""item.term""></div> </td>" + "\n";
                    idxHtmlTemplate += @"               <td>" + "\n";
                    idxHtmlTemplate += @"                   <div class=""search-delete"" data-if=""$sc.canDelete(#{@index})"" data-click=""$sc.handleDelete(#{@index})""></div>" + "\n";
                    idxHtmlTemplate += @"                   </td>" + "\n";
                    idxHtmlTemplate += @"           </tr>" + "\n";
                    idxHtmlTemplate += @"         </table> " + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"   </div>  " + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Search results -->" + "\n";
                    idxHtmlTemplate += @"<div class=""searchresults left-pane"" data-css=""width: sidebar_width | screen: 'desktop'; left: sidebar_width | screen: 'desktop', dir: 'ltr'; right: sidebar_width | screen: 'desktop', dir: 'rtl'"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; search-sidebar: @KEY_SEARCH_LOCATION == 'tabbar'; search-content: @KEY_SEARCH_LOCATION == 'content'; layout-visible: $mc.isSearchMode(@KEY_ACTIVE_TAB, @active_content); mobile-header-visible: @.l.mobile_header_visible"" data-scroll=""@.l.load_more_results(true) | debounce: 50, delta: 100"">" + "\n";
                    idxHtmlTemplate += @"   " + "\n";
                    idxHtmlTemplate += @"   <div class=""wSearchResults"" id=""searchresults"">" + "\n";
                    idxHtmlTemplate += @"       <div class=""wSearchResultSettings rh-hide"">" + "\n";
                    idxHtmlTemplate += @"           <div class=""wSearchHighlight"">" + "\n";
                    idxHtmlTemplate += @"               <input id=""highlightsearch"" type=""checkbox"" checked="""" class=""wSearchHighlight"" onclick=""onToggleHighlightSearch()"" />" + "\n";
                    idxHtmlTemplate += @"           </div>" + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"       <div class=""wSearchMessage"" data-if=""!@EVT_SEARCH_IN_PROGRESS"">" + "\n";
                    idxHtmlTemplate += @"           <span id=""searchMsg"" class=""wSearchMessage"">2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。</span> " + "\n";
                    idxHtmlTemplate += @"            " + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"       <div data-class=""loading: EVT_SEARCH_IN_PROGRESS"" data-if=""EVT_SEARCH_IN_PROGRESS""></div>" + "\n";
                    idxHtmlTemplate += @"       <p class=""progressbar"" data-if=""KEY_SEARCH_PROGRESS""><span data-text=""KEY_SEARCH_PROGRESS""></span>%</p>" + "\n";
                    idxHtmlTemplate += @"       <div class=""wSearchResultItemsBlock"" data-if=""!@EVT_SEARCH_IN_PROGRESS"">" + "\n";
                    idxHtmlTemplate += @"           <div data-rhwidget=""List: key: @.p.searchresults, spliton: index % @MAX_RESULTS == 14"" data-config=""loadmore: '.l.load_more_results', loaded: '.l.results_loaded'"">" + "\n";
                    idxHtmlTemplate += @"               <div class=""wSearchResultItem"">" + "\n";
                    idxHtmlTemplate += @"                   <a class=""nolink"" data-i-href=""item.strUrl+@.p.searchresultparams"">" + "\n";
                    idxHtmlTemplate += @"                       <div class=""wSearchResultTitle"" data-itext=""item.strTitle""></div>" + "\n";
                    idxHtmlTemplate += @"                   </a>" + "\n";
                    idxHtmlTemplate += @"                   <div class=""wSearchContent"">" + "\n";
                    idxHtmlTemplate += @"                       <span class=""wSearchContext"" data-itext=""item.strSummary""></span>" + "\n";
                    idxHtmlTemplate += @"                   </div>" + "\n";
                    idxHtmlTemplate += @"                   <div class=""wSearchURL"">" + "\n";
                    idxHtmlTemplate += @"                       <span class=""wSearchURL"" data-itext=""item.strBreadcrumbs""></span>" + "\n";
                    idxHtmlTemplate += @"                   </div>" + "\n";
                    idxHtmlTemplate += @"               </div>" + "\n";
                    idxHtmlTemplate += @"           </div>" + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"       <div data-if=""@.p.searchresults.length && !@EVT_SEARCH_IN_PROGRESS"" class=""wSearchResultsEnd"">" + "\n";
                    idxHtmlTemplate += @"           <span>検索結果の最後です。</span>" + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Topics -->" + "\n";
                    idxHtmlTemplate += @"<div class=""topic main"" data-css=""left: sidebar_width | screen: 'desktop', dir: 'ltr'; right: sidebar_width | screen: 'desktop', dir: 'rtl'"" data-class=""sidebar-opened: $mc.isSidebarTab(@KEY_ACTIVE_TAB); desktop-sidebar-hidden: @.l.desktop_sidebar_visible == false || @.l.desktop_sidebar_available === false; mobile-header-visible: @.l.mobile_header_visible"">" + "\n";
                    idxHtmlTemplate += @"   <div class=""functionholder"">" + "\n";
                    idxHtmlTemplate += @"     <div class=""buttons"">" + "\n";
                    idxHtmlTemplate += @"       <div class=""print_page"" id=""print_page"">" + "\n";
                    idxHtmlTemplate += @"         <div class=""print_page_area"" title=""いま表示されているページをPDFとして保存します。""><img src=""./template/Azure_Blue01/icon_pdf.png""><p class=""print_page_title"">ページ印刷</p></div>" + "\n";
                    idxHtmlTemplate += @"       </div>" + "\n";
                    idxHtmlTemplate += @"       <p class=""fontchange_title"">文字サイズ</p>" + "\n";
                    idxHtmlTemplate += @"       <div class=""fontsize_change"" id=""fontsize_small""><span>小</span></div>" + "\n";
                    idxHtmlTemplate += @"       <div class=""fontsize_change"" id=""fontsize_medium""><span>中</span></div>" + "\n";
                    idxHtmlTemplate += @"       <div class=""fontsize_change"" id=""fontsize_large""><span>大</span></div>" + "\n";
                    idxHtmlTemplate += @"       <div class=""brs_previous"" data-attr=""title:@KEY_LNG.Prev""><a id=""browseSeqBack"" class=""wBSBackButton"" data-rhwidget=""Basic"" data-attr=""href:.l.brsBack"" data-css=""visibility: @.l.brsBack?'visible':'hidden'"">&nbsp;</a>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"       <div class=""brs_next"" data-attr=""title:@KEY_LNG.Next""><a id=""browseSeqNext"" class=""wBSNextButton"" data-rhwidget=""Basic"" data-attr=""href:.l.brsNext"" data-css=""visibility: @.l.brsNext?'visible':'hidden'"">&nbsp;</a>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"     </div>" + "\n";
                    idxHtmlTemplate += @"   </div>" + "\n";
                    idxHtmlTemplate += @"   <div class=""topic-state"" data-class=""loading: EVT_TOPIC_LOADING"" data-if=""EVT_TOPIC_LOADING""></div>" + "\n";
                    idxHtmlTemplate += @"   <iframe class=""topic"" name=""rh_default_topic_frame_name""></iframe>" + "\n";
                    idxHtmlTemplate += @"   <a class=""to_top"" data-trigger=""EVT_SCROLL_TO_TOP"" data-attr=""title:@KEY_LNG.ToTopTip"">&#160;</a> " + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Social media buttons -->" + "\n";
                    idxHtmlTemplate += @"<div class=""social_buttons"" data-if=""@KEY_FEATURE.social === true && (@KEY_SCREEN_PHONE == false || (@KEY_SCREEN_PHONE == true && $mc.isSidebarTab(@KEY_ACTIVE_TAB) != true && $mc.isSearchMode(@KEY_ACTIVE_TAB, @active_content) != true && @KEY_ACTIVE_TAB != 'filter'))"" data-class=""opened: @.l.social_opened === true;"">" + "\n";
                    idxHtmlTemplate += @"   <a class=""social_buttons_controller"" href=""javascript:rh.model.publish('l.social_opened', rh.model.get('l.social_opened') === true ? false : true)""></a>" + "\n";
                    idxHtmlTemplate += @"   <div class=""fb-button"" data-if=""KEY_FEATURE.facebook""><iframe id=""bf-iframe"" style=""border:none; overflow:hidden;""></iframe></div>" + "\n";
                    idxHtmlTemplate += @"   <div class=""twitter-button"" id=""twitter-holder"" data-if=""KEY_FEATURE.twitter""></div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- pdf modal -->" + "\n";
                    idxHtmlTemplate += @"<div id=""modalPdf"">" + "\n";
                    idxHtmlTemplate += @"  <div id=""modalPdfContents"">" + "\n";
                    idxHtmlTemplate += @"    <div id=""modalPdfContentsHeader"">" + "\n";
                    idxHtmlTemplate += @"      <div id=""modalPdfContentsTitle"">PDF出力</div>" + "\n";
                    idxHtmlTemplate += @"      <div id=""buttonCloseModalPdf""></div>" + "\n";
                    idxHtmlTemplate += @"    </div>" + "\n";
                    idxHtmlTemplate += @"    <div id=""modalPdfLoaderTitle"">PDFプレビュー</div>" + "\n";
                    idxHtmlTemplate += @"    <div id=""modalPdfLoaderWrap"">" + "\n";
                    idxHtmlTemplate += @"      <div id=""modalPdfLoaderLoading""></div>" + "\n";
                    idxHtmlTemplate += @"      <div id=""modalPdfLoader""></div>" + "\n";
                    idxHtmlTemplate += @"    </div>" + "\n";
                    idxHtmlTemplate += @"    <div id=""modalPdfContentsFooter"">" + "\n";
                    idxHtmlTemplate += @"      <div id=""modalPdfPager"">" + "\n";
                    idxHtmlTemplate += @"        <div id=""modalPdfPagePrev"" class=""modalPdfPageButton off""></div>" + "\n";
                    idxHtmlTemplate += @"        <div id=""modalPdfPage""><span id=""modalPdfPageCurrent"">-</span>/<span id=""modalPdfPageAll"">-</span>ページ</div>" + "\n";
                    idxHtmlTemplate += @"        <div id=""modalPdfPageNext"" class=""modalPdfPageButton off""></div>" + "\n";
                    idxHtmlTemplate += @"      </div>" + "\n";
                    idxHtmlTemplate += @"      <ul id=""modalPdfButtons"">" + "\n";
                    idxHtmlTemplate += @"        <li id=""buttonOutputPdf"">PDF出力</li>" + "\n";
                    idxHtmlTemplate += @"        <li id=""buttonCancelPdf"">キャンセル</li>" + "\n";
                    idxHtmlTemplate += @"      </ul>" + "\n";
                    idxHtmlTemplate += @"    </div>" + "\n";
                    idxHtmlTemplate += @"  </div>" + "\n";
                    idxHtmlTemplate += @"  <div id=""modalPdfBg""></div>" + "\n";
                    idxHtmlTemplate += @"</div>" + "\n";
                    idxHtmlTemplate += @"<!-- Scripts -->" + "\n";
                    idxHtmlTemplate += @"<script src=""template/Azure_Blue01/usersettings.js"" type=""text/javascript""></script>" + "\n";
                    idxHtmlTemplate += @"<script>" + "\n";
                    idxHtmlTemplate += @"if(useTwitter === true) {" + "\n";
                    idxHtmlTemplate += @"   !function(d,s,id){var js,fjs=d.getElementsByTagName(s)[0],p=/^http:/.test(d.location)?'http':'https';if(!d.getElementById(id)){js=d.createElement(s);js.id=id;js.src=p+'://platform.twitter.com/widgets.js';fjs.parentNode.insertBefore(js,fjs);}}(document, 'script', 'twitter-wjs');" + "\n";
                    idxHtmlTemplate += @"}" + "\n";
                    idxHtmlTemplate += @"</script>" + "\n";
                    //idxHtmlTemplate += @"<script type=""text/javascript"" src=""whxdata/whtagdata.js""></script>" + "\n";
                    idxHtmlTemplate += @"</body>" + "\n";
                    idxHtmlTemplate += @"</html>" + "\n";

                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\index.html", false, Encoding.UTF8);
                    sw.Write(idxHtmlTemplate);
                    sw.Close();

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
                    #endregion

                    if (isEdgeTracker)
                    {
                        string[] hyousiGazo = { "EdgeTracker_logo50mm.png", "MJS_LOGO_255.gif", "hyousi.png" };
                        foreach (var hyousi in hyousiGazo)
                        {
                            Bitmap bmp = new Bitmap(assembly.GetManifestResourceStream("WordAddIn1.Resources." + hyousi));
                            bmp.Save(rootPath + "\\" + exportDir + "\\pict\\" + hyousi);
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
                        // Click to MSJ image, open link https://www.mjs.co.jp/
                        htmlCoverTemplate1 += @"<p style=""text-align: center; margin-top: 80pt;""><a href=""https://www.mjs.co.jp"" target=""_blank""><img style=""border: currentColor; border-image: none; width: 100%; max-width: 255px;"" alt="""" src=""pict/MJS_LOGO_255.gif"" border=""0""></a></p>" + "\n";
                    }
                    else if (isEasyCloud)
                    {
                        if (File.Exists(rootPath + "\\" + exportDir + "\\template\\images\\cover-background.png"))
                            htmlCoverTemplate1 += @"<body style=""text-justify-trim: punctuation; background-image: url('template/images/cover-background.png');background-repeat: no-repeat; background-position: 0px 300px;"">" + "\n";
                        else
                            #region
                            htmlCoverTemplate1 += @"<body>" + "\n";

                        htmlCoverTemplate1 += @"<p class=""manual_title"" style=""line-height: 130%;"">" + manualTitle + "</p>" + "\n";
                        htmlCoverTemplate1 += @"<p class=""manual_subtitle"">" + manualSubTitle + "</p>" + "\n";

                        if (File.Exists(rootPath + "\\" + exportDir + "\\template\\images\\cover-4.png"))
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

                    string htmlTemplate1 = "";
                    htmlTemplate1 += @"<!DOCTYPE HTML>" + "\n";
                    htmlTemplate1 += @"<html>" + "\n";
                    htmlTemplate1 += @"<head>" + "\n";
                    htmlTemplate1 += @"<meta http-equiv=""Content-Type"" content=""text/html; charset=utf-8"" />" + "\n";
                    htmlTemplate1 += @"<meta name=""generator"" content=""Adobe RoboHelp 2017"" />" + "\n";
                    htmlTemplate1 += @"<title></title>" + "\n";
                    htmlTemplate1 += @"<link rel=""StyleSheet"" href=""MJSHELP2002.css"" type=""text/css"" />" + "\n";
                    htmlTemplate1 += @"<link rel=""StyleSheet"" href=""font.css"" type=""text/css"" />" + "\n";
                    htmlTemplate1 += @"<link rel=""StyleSheet"" href=""resp.css"" type=""text/css"" />" + "\n";
                    htmlTemplate1 += @"<style type=""text/css"">" + "\n";
                    htmlTemplate1 += @"<!--" + "\n";
                    htmlTemplate1 += @"A:visited { color:purple; }" + "\n";
                    htmlTemplate1 += @"A:link { color:#337AB7; }" + "\n";
                    htmlTemplate1 += @"-->" + "\n";
                    htmlTemplate1 += @"</style>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" language=""JavaScript"">" + "\n";
                    htmlTemplate1 += @"//<![CDATA[" + "\n";
                    htmlTemplate1 += @"function reDo() {" + "\n";
                    htmlTemplate1 += @"  if (innerWidth != origWidth || innerHeight != origHeight)" + "\n";
                    htmlTemplate1 += @"     location.reload();" + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"if ((parseInt(navigator.appVersion) == 4) && (navigator.appName == ""Netscape"")) {" + "\n";
                    htmlTemplate1 += @"   origWidth = innerWidth;" + "\n";
                    htmlTemplate1 += @"   origHeight = innerHeight;" + "\n";
                    htmlTemplate1 += @"   onresize = reDo;" + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"onerror = null; " + "\n";
                    htmlTemplate1 += @"//]]>" + "\n";
                    htmlTemplate1 += @"</script>" + "\n";
                    htmlTemplate1 += @"<style type=""text/css"">" + "\n";
                    htmlTemplate1 += @"<!--" + "\n";
                    htmlTemplate1 += @"div.WebHelpPopupMenu { position:absolute;" + "\n";
                    htmlTemplate1 += @"left:0px;" + "\n";
                    htmlTemplate1 += @"top:0px;" + "\n";
                    htmlTemplate1 += @"z-index:4;" + "\n";
                    htmlTemplate1 += @"visibility:hidden; }" + "\n";
                    htmlTemplate1 += @"p.WebHelpNavBar { text-align:right; }" + "\n";
                    htmlTemplate1 += @"-->" + "\n";
                    htmlTemplate1 += @"</style>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/rh.min.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/common.min.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/topic.min.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/constants.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/utils.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/mhutils.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/mhlang.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/mhver.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/settings.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/XmlJsReader.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/loadparentdata.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/loadscreen.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/loadprojdata.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/mhtopic.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/jquery-3.1.0.min.js""></script>" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" src=""template/scripts/resp.js""></script>" + "\n";
                    htmlTemplate1 += @" <link rel=""stylesheet"" type=""text/css"" href=""template/styles/widgets.min.css"">" + "\n";
                    htmlTemplate1 += @" <link rel=""stylesheet"" type=""text/css"" href=""template/styles/topic.min.css"">" + "\n";
                    htmlTemplate1 += @"<script type=""text/javascript"" >" + "\n";
                    htmlTemplate1 += @"gRootRelPath = ""."";" + "\n";
                    htmlTemplate1 += @"gCommonRootRelPath = ""."";" + "\n";
                    htmlTemplate1 += @"gTopicId = ""♪"";" + "\n";
                    htmlTemplate1 += @"refPage = {" + "\n";
                    foreach (var item in title4Collection)
                    {
                        htmlTemplate1 += item.Key + ":['" + item.Value[0] + "','" + item.Value[1] + "'],";
                    }
                    //foreach (var item in headerCollection)
                    //{
                    //    htmlTemplate1 += item.Key + ":['" + item.Value[0] + "','" + item.Value[1] + "'],";
                    //}
                    htmlTemplate1 += @"};" + "\n";
                    htmlTemplate1 += @"mergePage = {" + "\n";
                    foreach (var item in mergeScript)
                    {
                        htmlTemplate1 += item.Value.Split(new char[] { '♯' })[0] + ":'" + item.Key.Split(new char[] { '♯' })[0] + "',";
                    }

                    htmlTemplate1 += @"};" + "\n";
                    htmlTemplate1 += @"	 url = [];" + "\n";

                    htmlTemplate1 += @"function areDirectoriesEqual(relativePath) {" + "\n";
                    htmlTemplate1 += @"	if (!relativePath) return false;" + "\n";
                    htmlTemplate1 += @"	if (!relativePath.match(""/"") || relativePath.match(/^\.\/.*/)) return true;" + "\n";
                    htmlTemplate1 += @"	const currentUrl = window.location.href;" + "\n";
                    htmlTemplate1 += @"	const currentDir = currentUrl.substring(0, currentUrl.lastIndexOf('/'));" + "\n";
                    htmlTemplate1 += @"	const baseUrl = window.location.origin;" + "\n";
                    htmlTemplate1 += @"	const resolvedUrl = new URL(relativePath, baseUrl).href;" + "\n";
                    htmlTemplate1 += @"	const resolvedDir = resolvedUrl.substring(0, resolvedUrl.lastIndexOf('/'));" + "\n";
                    htmlTemplate1 += @"	return currentDir === resolvedDir;" + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"function checkFileName(path) {" + "\n";
                    htmlTemplate1 += @"	const fileName = path.split('/').pop().split(/[?#]/)[0];" + "\n";
                    htmlTemplate1 += @"	return fileName.match(/^[A-Z]{3}[0-9]{5}[.]html$/);" + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"function changeFileNameWithHash(path) {" + "\n";
                    htmlTemplate1 += @"	const parts = path.split('/');" + "\n";
                    htmlTemplate1 += @"	const filePart = parts.pop();" + "\n";
                    htmlTemplate1 += @"	const dirPart = parts.join('/');" + "\n";
                    htmlTemplate1 += @"	return `${dirPart}/#t=${filePart}`;" + "\n";
                    htmlTemplate1 += @"}" + "\n";

                    htmlTemplate1 += @"$(function () {" + "\n";
                    htmlTemplate1 += @"	refPage = Object.values (refPage).sort ((a, b) => b[1].length - a[1].length).reduce (" + "\n";
                    htmlTemplate1 += @"	  (result, value) => {" + "\n";
                    htmlTemplate1 += @"		// tìm key tương ứng với value" + "\n";
                    htmlTemplate1 += @"		const key = Object.keys (refPage).find (k => refPage [k] === value);" + "\n";
                    htmlTemplate1 += @"		result [key] = value;" + "\n";
                    htmlTemplate1 += @"		return result;" + "\n";
                    htmlTemplate1 += @"	  }," + "\n";
                    htmlTemplate1 += @"	  {}" + "\n";
                    htmlTemplate1 += @"	);" + "\n";
                    htmlTemplate1 += @"	" + "\n";
                    htmlTemplate1 += @"	let text = window.location.href;" + "\n";
                    htmlTemplate1 += @"	if (text.match(""[/][A-Z]{3}[0-9]{5}[.]html"") != null) {" + "\n";
                    htmlTemplate1 += @"		let id = text.match(""[/][A-Z]{3}[0-9]{5}[.]html"")[0].replace(""/"", """").replace("".html"", """");" + "\n";
                    htmlTemplate1 += @"		// check id in mergePage" + "\n";
                    htmlTemplate1 += @"		for (let key in mergePage) {" + "\n";
                    htmlTemplate1 += @"			if(mergePage[key]==id){" + "\n";
                    htmlTemplate1 += @"				var newid=findFirstPageInMerge(id);" + "\n";
                    htmlTemplate1 += @"						let ref = window.location.href.split('.html#')[1];" + "\n";
                    htmlTemplate1 += @"				if (ref === '' || ref === undefined) {" + "\n";
                    htmlTemplate1 += @"					document.location.href = newid + "".html#"" + id;" + "\n";
                    htmlTemplate1 += @"					} else {" + "\n";
                    htmlTemplate1 += @"						document.location.href = newid + "".html#"" + ref;" + "\n";
                    htmlTemplate1 += @"					}" + "\n";
                    htmlTemplate1 += @"			}" + "\n";
                    htmlTemplate1 += @"	}" + "\n";
                    htmlTemplate1 += @"		loadNextPage(id)" + "\n";
                    htmlTemplate1 += @"		if(url.length>0){" + "\n";
                    htmlTemplate1 += @"		Promise.all(url.map(x => x.api))" + "\n";
                    htmlTemplate1 += @"		.then(responses => responses.forEach(" + "\n";
                    htmlTemplate1 += @"		(response, index) => {" + "\n";
                    htmlTemplate1 += @"		var html = $(response).find(""div"").parent();" + "\n";
                    htmlTemplate1 += @"			var di = $('body');" + "\n";
                    htmlTemplate1 += @"				$(html).find(""div:first"").html(""<a id='"" + url[index].page + ""'></a>"");" + "\n";
                    htmlTemplate1 += @"				$(html).find(""p.NoPageBreak"").attr(""class"", ""Heading2"");" + "\n";
                    htmlTemplate1 += @"					di.append(html);" + "\n";
                    htmlTemplate1 += @"		$("".ref"").each(function () {" + "\n";
                    htmlTemplate1 += @"			var refname = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"			$(""[name="" + refname + ""]"").each(function () {" + "\n";
                    htmlTemplate1 += @"				$(this).append(""<a name='"" + refname + ""'>"");" + "\n";
                    htmlTemplate1 += @"			})" + "\n";
                    htmlTemplate1 += @"		});" + "\n";
                    htmlTemplate1 += @"		$("".MJS_ref"").each(function(){" + "\n";
                    htmlTemplate1 += @"             $(this).find('a').each(function () {" + "\n";
                    htmlTemplate1 += @"        	    var name = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"        	    if (name?.indexOf(""_ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        	    } else {" + "\n";
                    //htmlTemplate1 += @"        		    let currentUri = $(this).attr('href');" + "\n";
                    htmlTemplate1 += @"        		        let currentUri = $(this).attr('href').replace(/^\.\//, '');" + "\n";
                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"        		        if (currentUri.match(/^https?:/)) {" + "\n";
                    htmlTemplate1 += @"        		    	    // 外部リンク" + "\n";
                    htmlTemplate1 += @"        		        } else if (currentUri.match(/^#[A-Z]{3}[0-9]{5}$/)) {" + "\n";
                    htmlTemplate1 += @"        		    	    // 内部リンク" + "\n";
                    htmlTemplate1 += @"        		        } else if (!areDirectoriesEqual(currentUri) && checkFileName(currentUri)) {" + "\n";
                    htmlTemplate1 += @"        		    	    // 外部参照" + "\n";
                    htmlTemplate1 += @"        		    	    $(this).attr('href', changeFileNameWithHash(currentUri));" + "\n";
                    htmlTemplate1 += @"        		        } else " + "\n";
                    htmlTemplate1 += @"" + "\n";

                    htmlTemplate1 += @"        		    if (currentUri.split('.html#')[0] == currentUri.split('.html#')[1]) {" + "\n";
                    htmlTemplate1 += @"        		        for (i = 0; i < Object.keys(refPage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        		            if (Object.keys(refPage)[i] == currentUri.split('.html#')[1]) {" + "\n";
                    htmlTemplate1 += @"        		                var key = Object.keys(refPage)[i];" + "\n";
                    htmlTemplate1 += @"        		                var expectUrl = refPage[key][0] + "".html#"" + currentUri.split("".html#"")[1];" + "\n";
                    htmlTemplate1 += @"        		                $(this).attr('href', expectUrl);" + "\n";
                    htmlTemplate1 += @"        		                break;" + "\n";
                    htmlTemplate1 += @"        		            }" + "\n";
                    htmlTemplate1 += @"        		        }" + "\n";
                    htmlTemplate1 += @"        		      } else {" + "\n";
                    htmlTemplate1 += @"        		    let subDestinationId = currentUri.split('.')[0];" + "\n";
                    htmlTemplate1 += @"        		    let destinationId = currentUri.split('#')[1] == undefined ? subDestinationId : currentUri.split('#')[1];" + "\n";
                    htmlTemplate1 += @"        		    let temp = mergePage;" + "\n";
                    htmlTemplate1 += @"        		    for (i = 0; i < Object.keys(mergePage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        			    let startId = Object.keys(temp).find(key => temp[key] === subDestinationId);" + "\n";
                    htmlTemplate1 += @"        				    if (startId == undefined) break;" + "\n";
                    htmlTemplate1 += @"        				    subDestinationId = startId;" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        		    let href = """"" + "\n";
                    htmlTemplate1 += @"        		    if (subDestinationId == destinationId && subDestinationId?.indexOf(""_Ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        		    	for (i = 0; i < Object.keys(refPage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        		    		if (Object.keys(refPage)[i] == subDestinationId) {" + "\n";
                    htmlTemplate1 += @"        		    			var key = Object.keys(refPage)[i];" + "\n";
                    htmlTemplate1 += @"        		    			href = refPage[key][0] + "".html#"" + destinationId;" + "\n";
                    htmlTemplate1 += @"        		    			break;" + "\n";
                    htmlTemplate1 += @"        		    		}" + "\n";
                    htmlTemplate1 += @"        		    	}" + "\n";
                    htmlTemplate1 += @"        		    } else {" + "\n";
                    htmlTemplate1 += @"        		    	href = subDestinationId + '.html#' + destinationId;" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        		    $(this).attr('href', href);" + "\n";
                    htmlTemplate1 += @"        		    $(this).attr('onclick', ""anchorElement(href.split('#Ref')[0])"");" + "\n";

                    htmlTemplate1 += @"        		   }" + "\n";
                    htmlTemplate1 += @"        	    }" + "\n";
                    htmlTemplate1 += @"         })" + "\n";
                    htmlTemplate1 += @"             $(this).find('.ref').each(function () {" + "\n";
                    htmlTemplate1 += @"        	    var name = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"        	    if (name?.indexOf(""_ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        		    name = name.replace(""_ref"", """");" + "\n";
                    htmlTemplate1 += @"        		    for (let key in refPage) {" + "\n";
                    htmlTemplate1 += @"        					if (key == name) {" + "\n";
                    htmlTemplate1 += @"        						var replaceStr = refPage[key][1];" + "\n";
                    htmlTemplate1 += @"        						if ($("".ref[name="" + key + ""]"").length > 0) {" + "\n";
                    htmlTemplate1 += @"        							$(this).html($(this).html().replace(replaceStr, ""<a href='#"" + key + ""'>"" + replaceStr + ""</a>""));" + "\n";
                    htmlTemplate1 += @"        						} else {" + "\n";
                    htmlTemplate1 += @"        						    let expectUrl = findFirstPageInMerge(refPage[key][0]) + "".html#"" + key;" + "\n";
                    htmlTemplate1 += @"        							$(this).attr(""href"", expectUrl)" + "\n";
                    htmlTemplate1 += @"        						}" + "\n";
                    htmlTemplate1 += @"        						break;" + "\n";
                    htmlTemplate1 += @"        					}" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        	    }" + "\n";
                    htmlTemplate1 += @"         });" + "\n";
                    htmlTemplate1 += @"		    });" + "\n";
                    htmlTemplate1 += @"		" + "\n";
                    htmlTemplate1 += @"		if(text.indexOf(""#"")>0){" + "\n";
                    htmlTemplate1 += @"			window.location.href=text;" + "\n";
                    htmlTemplate1 += @"		}" + "\n";
                    htmlTemplate1 += @"		var di = $('body');" + "\n";
                    htmlTemplate1 += @"     di.html(""<div></div><main>"" + di.html() + ""</main>"");" + "\n";
                    htmlTemplate1 += @"	}" + "\n";
                    htmlTemplate1 += @"))" + "\n";
                    htmlTemplate1 += @"		" + "\n";
                    htmlTemplate1 += @"		}else{" + "\n";
                    htmlTemplate1 += @"		$("".ref"").each(function () {" + "\n";
                    htmlTemplate1 += @"			var refname = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"			$(""[name="" + refname + ""]"").each(function () {" + "\n";
                    htmlTemplate1 += @"				$(this).append(""<a name='"" + refname + ""'>"");" + "\n";
                    htmlTemplate1 += @"			})" + "\n";
                    htmlTemplate1 += @"		});" + "\n";
                    htmlTemplate1 += @"		$("".MJS_ref"").each(function(){" + "\n";
                    htmlTemplate1 += @"             $(this).find('a').each(function () {" + "\n";
                    htmlTemplate1 += @"        	    var name = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"        	    if (name?.indexOf(""_ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        	    } else {" + "\n";
                    //htmlTemplate1 += @"        		    let currentUri = $(this).attr('href');" + "\n";
                    htmlTemplate1 += @"        		        let currentUri = $(this).attr('href').replace(/^\.\//, '');" + "\n";
                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"        		        if (currentUri.match(/^https?:/)) {" + "\n";
                    htmlTemplate1 += @"        		        	// 外部リンク" + "\n";
                    htmlTemplate1 += @"        		        } else if (currentUri.match(/^#[A-Z]{3}[0-9]{5}$/)) {" + "\n";
                    htmlTemplate1 += @"        		        	// 内部リンク" + "\n";
                    htmlTemplate1 += @"        		        } else if (!areDirectoriesEqual(currentUri) && checkFileName(currentUri)) {" + "\n";
                    htmlTemplate1 += @"        		    	    // 外部参照" + "\n";
                    htmlTemplate1 += @"        		    	    $(this).attr('href', changeFileNameWithHash(currentUri));" + "\n";
                    htmlTemplate1 += @"        		        } else " + "\n";
                    htmlTemplate1 += @"" + "\n";

                    htmlTemplate1 += @"        		    if (currentUri.split('.html#')[0] == currentUri.split('.html#')[1]) {" + "\n";
                    htmlTemplate1 += @"        		        for (i = 0; i < Object.keys(refPage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        		            if (Object.keys(refPage)[i] == currentUri.split('.html#')[1]) {" + "\n";
                    htmlTemplate1 += @"        		                var key = Object.keys(refPage)[i];" + "\n";
                    htmlTemplate1 += @"        		                var expectUrl = refPage[key][0] + "".html#"" + currentUri.split("".html#"")[1];" + "\n";
                    htmlTemplate1 += @"        		                $(this).attr('href', expectUrl);" + "\n";
                    htmlTemplate1 += @"        		                break;" + "\n";
                    htmlTemplate1 += @"        		            }" + "\n";
                    htmlTemplate1 += @"        		        }" + "\n";
                    htmlTemplate1 += @"        		      } else {" + "\n";
                    htmlTemplate1 += @"        		      if (currentUri?.indexOf(""."") > 0 || currentUri?.indexOf(""_Ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        		    let subDestinationId = currentUri.split('.')[0];" + "\n";
                    htmlTemplate1 += @"        		    let destinationId = currentUri.split('#')[1] == undefined ? subDestinationId : currentUri.split('#')[1];" + "\n";
                    htmlTemplate1 += @"        		    let temp = mergePage;" + "\n";
                    htmlTemplate1 += @"        		    for (i = 0; i < Object.keys(mergePage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        			    let startId = Object.keys(temp).find(key => temp[key] === subDestinationId);" + "\n";
                    htmlTemplate1 += @"        				    if (startId == undefined) break;" + "\n";
                    htmlTemplate1 += @"        				    subDestinationId = startId;" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        		    let href = """"" + "\n";
                    htmlTemplate1 += @"        		    if (subDestinationId == destinationId && subDestinationId?.indexOf(""_Ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        		    	for (i = 0; i < Object.keys(refPage).length; i++) {" + "\n";
                    htmlTemplate1 += @"        		    		if (Object.keys(refPage)[i] == subDestinationId) {" + "\n";
                    htmlTemplate1 += @"        		    			var key = Object.keys(refPage)[i];" + "\n";
                    htmlTemplate1 += @"        		    			href = refPage[key][0] + "".html#"" + destinationId;" + "\n";
                    htmlTemplate1 += @"        		    			break;" + "\n";
                    htmlTemplate1 += @"        		    		}" + "\n";
                    htmlTemplate1 += @"        		    	}" + "\n";
                    htmlTemplate1 += @"        		    } else {" + "\n";
                    htmlTemplate1 += @"        		    	href = subDestinationId + '.html#' + destinationId;" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        		    $(this).attr('href', href);" + "\n";
                    htmlTemplate1 += @"        		    $(this).attr('onclick', ""anchorElement(href.split('#Ref')[0])"");" + "\n";
                    htmlTemplate1 += @"        		   }" + "\n";
                    htmlTemplate1 += @"        		   }" + "\n";
                    htmlTemplate1 += @"        	    }" + "\n";
                    htmlTemplate1 += @"         });" + "\n";
                    htmlTemplate1 += @"             $(this).find('.ref').each(function () {" + "\n";
                    htmlTemplate1 += @"        	    var name = $(this).attr(""name"");" + "\n";
                    htmlTemplate1 += @"        	    if (name?.indexOf(""_ref"") > -1) {" + "\n";
                    htmlTemplate1 += @"        		    name = name.replace(""_ref"", """");" + "\n";
                    htmlTemplate1 += @"        		    for (let key in refPage) {" + "\n";
                    htmlTemplate1 += @"        					if (key == name) {" + "\n";
                    htmlTemplate1 += @"        						var replaceStr = refPage[key][1];" + "\n";
                    htmlTemplate1 += @"        						if ($("".ref[name="" + key + ""]"").length > 0) {" + "\n";
                    htmlTemplate1 += @"        							$(this).html($(this).html().replace(replaceStr, ""<a href='#"" + key + ""'>"" + replaceStr + ""</a>""));" + "\n";
                    htmlTemplate1 += @"        						} else {" + "\n";
                    htmlTemplate1 += @"        						    let expectUrl = findFirstPageInMerge(refPage[key][0]) + "".html#"" + key;" + "\n";
                    htmlTemplate1 += @"        							$(this).attr(""href"", expectUrl);" + "\n";
                    htmlTemplate1 += @"        						}" + "\n";
                    htmlTemplate1 += @"        						break;" + "\n";
                    htmlTemplate1 += @"        					}" + "\n";
                    htmlTemplate1 += @"        		    }" + "\n";
                    htmlTemplate1 += @"        	    }" + "\n";
                    htmlTemplate1 += @"         });" + "\n";
                    htmlTemplate1 += @"		    });" + "\n";
                    htmlTemplate1 += @"		" + "\n";
                    htmlTemplate1 += @"		" + "\n";
                    htmlTemplate1 += @"		if(text.indexOf(""#"")>0){" + "\n";
                    htmlTemplate1 += @"			window.location.href=text;" + "\n";
                    htmlTemplate1 += @"		}" + "\n";
                    htmlTemplate1 += @"		var di = $('body');" + "\n";
                    htmlTemplate1 += @"     di.html(""<div></div><main>"" + di.html() + ""</main>"");" + "\n";
                    htmlTemplate1 += @"		}" + "\n";
                    htmlTemplate1 += @"		/*" + "\n";
                    htmlTemplate1 += @"		$('a').each(function(){" + "\n";
                    htmlTemplate1 += @"			if($(this).attr('href') !== undefined){" + "\n";
                    htmlTemplate1 += @"				var test = $(this).attr('href').match(/[A-Z]{3}[0-9]{5}[.]html/);" + "\n";
                    htmlTemplate1 += @"				if(test != null){" + "\n";
                    htmlTemplate1 += @"					test = test[0].replace("".html"","""");" + "\n";
                    htmlTemplate1 += @"					var lastPageInMerge = findFirstPageInMerge(test);" + "\n";
                    htmlTemplate1 += @"					if(lastPageInMerge!=test){" + "\n";
                    htmlTemplate1 += @"						$(this).attr('href',lastPageInMerge + "".html#"" + test);" + "\n";
                    htmlTemplate1 += @"					}" + "\n";
                    htmlTemplate1 += @"				}				" + "\n";
                    htmlTemplate1 += @"			}" + "\n";
                    htmlTemplate1 += @"		});*/" + "\n";

                    htmlTemplate1 += @"	}" + "\n";
                    htmlTemplate1 += @"});" + "\n";

                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"function findFirstPageInMerge(page){" + "\n";
                    htmlTemplate1 += @"	var found=false;" + "\n";
                    htmlTemplate1 += @"	for (let key in mergePage) {" + "\n";
                    htmlTemplate1 += @"		if(mergePage[key]==page){" + "\n";
                    htmlTemplate1 += @"			found = true;" + "\n";
                    htmlTemplate1 += @"			return findFirstPageInMerge(key);" + "\n";
                    htmlTemplate1 += @"			break;" + "\n";
                    htmlTemplate1 += @"		}" + "\n";
                    htmlTemplate1 += @"	}" + "\n";
                    htmlTemplate1 += @"	if(!found){" + "\n";
                    htmlTemplate1 += @"		return page;" + "\n";
                    htmlTemplate1 += @"	}" + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"" + "\n";
                    htmlTemplate1 += @"function anchorElement(url) {" + "\n";
                    htmlTemplate1 += @" if (window.location.href.indexOf("".html#"") > -1){" + "\n";
                    htmlTemplate1 += @"     window.location.href = url;" + "\n";
                    htmlTemplate1 += @" }" + "\n";
                    htmlTemplate1 += @"}" + "\n";

                    htmlTemplate1 += @"function loadNextPage(id) {" + "\n";
                    htmlTemplate1 += @"	if (mergePage[id] !== undefined) {" + "\n";
                    htmlTemplate1 += @"		url.push({" + "\n";
                    htmlTemplate1 += @"			api: $.ajax({" + "\n";
                    htmlTemplate1 += @"			url: mergePage[id] + "".html""" + "\n";
                    htmlTemplate1 += @"				})," + "\n";
                    htmlTemplate1 += @"			page:  mergePage[id]" + "\n";
                    htmlTemplate1 += @"				});" + "\n";
                    htmlTemplate1 += @"		loadNextPage(mergePage[id])" + "\n";
                    htmlTemplate1 += @"	} " + "\n";
                    htmlTemplate1 += @"}" + "\n";
                    htmlTemplate1 += @"</script>" + "\n";
                    htmlTemplate1 += @" <meta name=""topic-breadcrumbs"" content="""" />" + "\n";
                    htmlTemplate1 += @"</head>" + "\n";
                    htmlTemplate1 += @"<body style=""text-justify-trim: punctuation;"">" + "\n";

                    string htmlTemplate2 = "";
                    htmlTemplate2 += @"</body>" + "\n";
                    htmlTemplate2 += @"</html>" + "\n";

                    string searchJs = "";
                    searchJs += @"var searchWords = $('♪');" + "\n";
                    searchJs += @"var wide = Array(""０"",""１"",""２"",""３"",""４"",""５"",""６"",""７"",""８"",""９"",""Ａ"",""Ｂ"",""Ｃ"",""Ｄ"",""Ｅ"",""Ｆ"",""Ｇ"",""Ｈ"",""Ｉ"",""Ｊ"",""Ｋ"",""Ｌ"",""Ｍ"",""Ｎ"",""Ｏ"",""Ｐ"",""Ｑ"",""Ｒ"",""Ｓ"",""Ｔ"",""Ｕ"",""Ｖ"",""Ｗ"",""Ｘ"",""Ｙ"",""Ｚ"",""ａ"",""ｂ"",""ｃ"",""ｄ"",""ｅ"",""ｆ"",""ｇ"",""ｈ"",""ｉ"",""ｊ"",""ｋ"",""ｌ"",""ｍ"",""ｎ"",""ｏ"",""ｐ"",""ｑ"",""ｒ"",""ｓ"",""ｔ"",""ｕ"",""ｖ"",""ｗ"",""ｘ"",""ｙ"",""ｚ"",""ガ"",""ギ"",""グ"",""ゲ"",""ゴ"",""ザ"",""ジ"",""ズ"",""ゼ"",""ゾ"",""ダ"",""ヂ"",""ヅ"",""デ"",""ド"",""バ"",""ビ"",""ブ"",""ベ"",""ボ"",""パ"",""ピ"",""プ"",""ペ"",""ポ"",""。"",""「"",""」"",""、"",""ヲ"",""ァ"",""ィ"",""ゥ"",""ェ"",""ォ"",""ャ"",""ュ"",""ョ"",""ッ"",""ー"",""ア"",""イ"",""ウ"",""エ"",""オ"",""カ"",""キ"",""ク"",""ケ"",""コ"",""サ"",""シ"",""ス"",""セ"",""ソ"",""タ"",""チ"",""ツ"",""テ"",""ト"",""ナ"",""ニ"",""ヌ"",""ネ"",""ノ"",""ハ"",""ヒ"",""フ"",""ヘ"",""ホ"",""マ"",""ミ"",""ム"",""メ"",""モ"",""ヤ"",""ユ"",""ヨ"",""ラ"",""リ"",""ル"",""レ"",""ロ"",""ワ"",""ン"");" + "\n";
                    searchJs += @"var narrow = Array(""0"",""1"",""2"",""3"",""4"",""5"",""6"",""7"",""8"",""9"",""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k"",""l"",""m"",""n"",""o"",""p"",""q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z"",""a"",""b"",""c"",""d"",""e"",""f"",""g"",""h"",""i"",""j"",""k"",""l"",""m"",""n"",""o"",""p"",""q"",""r"",""s"",""t"",""u"",""v"",""w"",""x"",""y"",""z"",""ｶﾞ"",""ｷﾞ"",""ｸﾞ"",""ｹﾞ"",""ｺﾞ"",""ｻﾞ"",""ｼﾞ"",""ｽﾞ"",""ｾﾞ"",""ｿﾞ"",""ﾀﾞ"",""ﾁﾞ"",""ﾂﾞ"",""ﾃﾞ"",""ﾄﾞ"",""ﾊﾞ"",""ﾋﾞ"",""ﾌﾞ"",""ﾍﾞ"",""ﾎﾞ"",""ﾊﾟ"",""ﾋﾟ"",""ﾌﾟ"",""ﾍﾟ"",""ﾎﾟ"",""｡"",""｢"",""｣"",""､"",""ｦ"",""ｧ"",""ｨ"",""ｩ"",""ｪ"",""ｫ"",""ｬ"",""ｭ"",""ｮ"",""ｯ"",""ｰ"",""ｱ"",""ｲ"",""ｳ"",""ｴ"",""ｵ"",""ｶ"",""ｷ"",""ｸ"",""ｹ"",""ｺ"",""ｻ"",""ｼ"",""ｽ"",""ｾ"",""ｿ"",""ﾀ"",""ﾁ"",""ﾂ"",""ﾃ"",""ﾄ"",""ﾅ"",""ﾆ"",""ﾇ"",""ﾈ"",""ﾉ"",""ﾊ"",""ﾋ"",""ﾌ"",""ﾍ"",""ﾎ"",""ﾏ"",""ﾐ"",""ﾑ"",""ﾒ"",""ﾓ"",""ﾔ"",""ﾕ"",""ﾖ"",""ﾗ"",""ﾘ"",""ﾙ"",""ﾚ"",""ﾛ"",""ﾜ"",""ﾝ"");" + "\n";
                    searchJs += @"var hilight = Array(""(?:０|0)"",""(?:１|1)"",""(?:２|2)"",""(?:３|3)"",""(?:４|4)"",""(?:５|5)"",""(?:６|6)"",""(?:７|7)"",""(?:８|8)"",""(?:９|9)"",""(?:Ａ|A|ａ|a)"",""(?:Ｂ|B|ｂ|b)"",""(?:Ｃ|C|ｃ|c)"",""(?:Ｄ|D|ｄ|d)"",""(?:Ｅ|E|ｅ|e)"",""(?:Ｆ|F|ｆ|f)"",""(?:Ｇ|G|ｇ|g)"",""(?:Ｈ|H|ｈ|h)"",""(?:Ｉ|I|ｉ|i)"",""(?:Ｊ|J|ｊ|j)"",""(?:Ｋ|K|ｋ|k)"",""(?:Ｌ|L|ｌ|l)"",""(?:Ｍ|M|ｍ|m)"",""(?:Ｎ|N|ｎ|n)"",""(?:Ｏ|O|ｏ|o)"",""(?:Ｐ|P|ｐ|p)"",""(?:Ｑ|Q|ｑ|q)"",""(?:Ｒ|R|ｒ|r)"",""(?:Ｓ|S|ｓ|s)"",""(?:Ｔ|T|ｔ|t)"",""(?:Ｕ|U|ｕ|u)"",""(?:Ｖ|V|ｖ|v)"",""(?:Ｗ|W|ｗ|w)"",""(?:Ｘ|X|ｘ|x)"",""(?:Ｙ|Y|ｙ|y)"",""(?:Ｚ|Z|ｚ|z)"",""(?:ガ|ｶﾞ)"",""(?:ギ|ｷﾞ)"",""(?:グ|ｸﾞ)"",""(?:ゲ|ｹﾞ)"",""(?:ゴ|ｺﾞ)"",""(?:ザ|ｻﾞ)"",""(?:ジ|ｼﾞ)"",""(?:ズ|ｽﾞ)"",""(?:ゼ|ｾﾞ)"",""(?:ゾ|ｿﾞ)"",""(?:ダ|ﾀﾞ)"",""(?:ヂ|ﾁﾞ)"",""(?:ヅ|ﾂﾞ)"",""(?:デ|ﾃﾞ)"",""(?:ド|ﾄﾞ)"",""(?:バ|ﾊﾞ)"",""(?:ビ|ﾋﾞ)"",""(?:ブ|ﾌﾞ)"",""(?:ベ|ﾍﾞ)"",""(?:ボ|ﾎﾞ)"",""(?:パ|ﾊﾟ)"",""(?:ピ|ﾋﾟ)"",""(?:プ|ﾌﾟ)"",""(?:ペ|ﾍﾟ)"",""(?:ポ|ﾎﾟ)"",""(?:。|｡)"",""(?:「|｢)"",""(?:」|｣)"",""(?:、|､)"",""(?:ヲ|ｦ)"",""(?:ァ|ｧ)"",""(?:ィ|ｨ)"",""(?:ゥ|ｩ)"",""(?:ェ|ｪ)"",""(?:ォ|ｫ)"",""(?:ャ|ｬ)"",""(?:ュ|ｭ)"",""(?:ョ|ｮ)"",""(?:ッ|ｯ)"",""(?:ー|ｰ)"",""(?:ア|ｱ)"",""(?:イ|ｲ)"",""(?:ウ|ｳ)"",""(?:エ|ｴ)"",""(?:オ|ｵ)"",""(?:カ|ｶ)"",""(?:キ|ｷ)"",""(?:ク|ｸ)"",""(?:ケ|ｹ)"",""(?:コ|ｺ)"",""(?:サ|ｻ)"",""(?:シ|ｼ)"",""(?:ス|ｽ)"",""(?:セ|ｾ)"",""(?:ソ|ｿ)"",""(?:タ|ﾀ)"",""(?:チ|ﾁ)"",""(?:ツ|ﾂ)"",""(?:テ|ﾃ)"",""(?:ト|ﾄ)"",""(?:ナ|ﾅ)"",""(?:ニ|ﾆ)"",""(?:ヌ|ﾇ)"",""(?:ネ|ﾈ)"",""(?:ノ|ﾉ)"",""(?:ハ|ﾊ)"",""(?:ヒ|ﾋ)"",""(?:フ|ﾌ)"",""(?:ヘ|ﾍ)"",""(?:ホ|ﾎ)"",""(?:マ|ﾏ)"",""(?:ミ|ﾐ)"",""(?:ム|ﾑ)"",""(?:メ|ﾒ)"",""(?:モ|ﾓ)"",""(?:ヤ|ﾔ)"",""(?:ユ|ﾕ)"",""(?:ヨ|ﾖ)"",""(?:ラ|ﾗ)"",""(?:リ|ﾘ)"",""(?:ル|ﾙ)"",""(?:レ|ﾚ)"",""(?:ロ|ﾛ)"",""(?:ワ|ﾜ)"",""(?:ン|ﾝ)"");" + "\n";
                    searchJs += @"function selectorEscape(val){" + "\n";
                    searchJs += @"  return val.replace(/[-\/\\^$*+?.()|[\]{}\!]/g, '\\$&');" + "\n";
                    searchJs += @"}" + "\n\n";
                    searchJs += @"$(function(){" + "\n";
                    searchJs += @"  $(document).on(""click"", ""ul.toc li.book"", function() {" + "\n";
                    //searchJs += @"    $(this).children(""a[href='#']"").each(function(){" + "\n";
                    //searchJs += @"      $(this).attr(""href"", ""javascript:void 0;"");" + "\n";
                    //searchJs += @"    });" + "\n";
                    searchJs += @"    if($(this).children(""a[href='#'],a[href='javascript:void 0;']"").length == 0)" + "\n";
                    searchJs += @"    {" + "\n";
                    searchJs += @"      $(this).children(""a"").each(function(){" + "\n";
                    searchJs += @"        location.href=location.href.replace(location.hash,"""")+""#t=""+$(this).attr(""href"");" + "\n";
                    searchJs += @"      });" + "\n";
                    searchJs += @"    }" + "\n";
                    searchJs += @"  });" + "\n";
                    searchJs += @"  $("".wSearchField"").each(function() {" + "\n";
                    searchJs += @"    $(this).off();" + "\n";
                    searchJs += @"  });" + "\n";
                    searchJs += @"  $(document).on(""keyup"", "".wSearchField"", function(){" + "\n";
                    searchJs += @"    if($(this).val() == """")" + "\n";
                    searchJs += @"    {" + "\n";
                    searchJs += @"      $("".wSearchResultItemsBlock"").html("""");" + "\n";
                    searchJs += @"      $("".wSearchResultsEnd"").addClass(""rh-hide"");" + "\n";
                    searchJs += @"      $("".wSearchResultsEnd"").attr(""hidden"", """");" + "\n";
                    searchJs += @"      $(""#searchMsg"").html(""2つ以上の語句を入力して検索する場合は、スペース（空白）で区切ります。"");" + "\n";
                    searchJs += @"      $(""iframe.topic"").contents().find("".keyword"").each(function() {" + "\n";
                    searchJs += @"        for(var i = 0; i < $(this)[0].childNodes.length; i ++)" + "\n";
                    searchJs += @"        {" + "\n";
                    searchJs += @"          this.parentNode.insertBefore($(this)[0].childNodes[i], this)" + "\n";
                    searchJs += @"        }" + "\n";
                    searchJs += @"        $(this).remove();" + "\n";
                    searchJs += @"      });" + "\n";
                    searchJs += @"    }" + "\n";
                    searchJs += @"    else" + "\n";
                    searchJs += @"    {" + "\n";
                    searchJs += @"      $(""#searchMsg"").html("""");" + "\n";
                    searchJs += @"      var searchWordTmp = $(this).val().replace(/(.*?)(?:　| )+(.*?)/g, ""$1 $2"").trim().toLowerCase();" + "\n";
                    searchJs += @"      for(i = 0; i < wide.length; i ++)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        searchWordTmp = searchWordTmp.split(wide[i]).join(narrow[i]);" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"      var searchWord = searchWordTmp.split("" "");" + "\n";
                    searchJs += @"      var searchQuery = """";" + "\n";
                    searchJs += @"      for(i = 0; i < searchWord.length; i ++)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        searchQuery += "":contains("" + searchWord[i] + "")"";" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"      " + "\n";
                    searchJs += @"      var findItems = searchWords.find("".search_word""+searchQuery);" + "\n";
                    searchJs += @"      if(findItems.length != 0)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        $("".wSearchResultsEnd"").removeClass(""rh-hide"");" + "\n";
                    searchJs += @"        $("".wSearchResultsEnd"").removeAttr(""hidden"");" + "\n";
                    searchJs += @"        $("".wSearchResultItemsBlock"").html("""");" + "\n";
                    searchJs += @"        findItems.each(function() {" + "\n";
                    searchJs += @"          var displayText = $(this).parent().find("".displayText"").text();" + "\n";
                    searchJs += @"          $("".wSearchResultItemsBlock"").append($(""<div class='wSearchResultItem'><a class='nolink' href='./""+$(this).parent().attr(""id"")+"".html'><div class='wSearchResultTitle'>""+$(this).parent().find("".search_title"").html()+""</div></a><div class='wSearchContent'><span class='wSearchContext'>""+displayText+""</span></div></div>""));" + "\n";
                    searchJs += @"        });" + "\n";
                    searchJs += @"        $(""iframe.topic"").contents().find("".keyword"").each(function() {" + "\n";
                    searchJs += @"          for(var i = 0; i < $(this)[0].childNodes.length; i ++)" + "\n";
                    searchJs += @"          {" + "\n";
                    searchJs += @"            this.parentNode.insertBefore($(this)[0].childNodes[i], this)" + "\n";
                    searchJs += @"          }" + "\n";
                    searchJs += @"          $(this).remove();" + "\n";
                    searchJs += @"        });" + "\n";
                    searchJs += @"        for(i = 0; i < searchWord.length; i ++)" + "\n";
                    searchJs += @"        {" + "\n";
                    searchJs += @"        	searchWord[i] = selectorEscape(searchWord[i].replace("">"", ""&gt;"").replace(""<"", ""&lt;""));" + "\n";
                    searchJs += @"        }" + "\n";
                    searchJs += @"        " + "\n";
                    searchJs += @"        var hilightWord = searchWord.join(""|"");" + "\n";
                    searchJs += @"        for(i = 0; i < hilight.length; i ++)" + "\n";
                    searchJs += @"        {" + "\n";
                    searchJs += @"	        var reg = new RegExp(hilight[i], ""gm"");" + "\n";
                    searchJs += @"        	hilightWord = hilightWord.replace(reg, hilight[i]);" + "\n";
                    searchJs += @"        }" + "\n";
                    searchJs += @"        var reg = new RegExp(""(""+hilightWord+"")(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        var regnbsp = new RegExp(""&nbsp;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        var reggt = new RegExp(""&gt;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        var reglt = new RegExp(""&lt;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        var regquot = new RegExp(""&quot;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        var regamp = new RegExp(""&amp;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        $(""iframe.topic"").contents().find(""body"").html($(""iframe.topic"").contents().find(""body"").html().replace(regnbsp, ""　"").replace(reggt, "">"").replace(reglt, ""<"").replace(regquot, '""').replace(regamp, ""&"").replace(reg, ""<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>""));" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"      else" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        $(""iframe.topic"").contents().find("".keyword"").each(function() {" + "\n";
                    searchJs += @"          for(var i = 0; i < $(this)[0].childNodes.length; i ++)" + "\n";
                    searchJs += @"          {" + "\n";
                    searchJs += @"            this.parentNode.insertBefore($(this)[0].childNodes[i], this)" + "\n";
                    searchJs += @"          }" + "\n";
                    searchJs += @"          $(this).remove();" + "\n";
                    searchJs += @"        });" + "\n";
                    searchJs += @"        $("".wSearchResultsEnd"").addClass(""rh-hide"");" + "\n";
                    searchJs += @"        $("".wSearchResultsEnd"").attr(""hidden"", """");" + "\n";
                    searchJs += @"        $("".wSearchResultItemsBlock"").html("""");" + "\n";
                    searchJs += @"        displayText = ""検索条件に一致するトピックはありません。"";" + "\n";
                    searchJs += @"        $("".wSearchResultItemsBlock"").append($(""<div class='wSearchResultItem'><div class='wSearchContent'><span class='wSearchContext'>""+displayText+""</span></div></div>""));" + "\n";
                    searchJs += @"        //this.parentNode.insertBefore($(this)[0].childNodes[0], this);" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"    }" + "\n";
                    searchJs += @"  });" + "\n";
                    searchJs += @"  $(""iframe.topic"").on(""load"", function(){" + "\n";
                    searchJs += @"    if($("".search-input"", document).is("":not(.rh-hide)"") && ($("".wSearchField"", document).val() != """"))" + "\n";
                    searchJs += @"    {" + "\n";
                    searchJs += @"      var searchWordTmp = $("".wSearchField"", document).val().split(""　"").join("" "").trim();" + "\n";
                    searchJs += @"      searchWordTmp = searchWordTmp.split(""  "").join("" "");" + "\n";
                    searchJs += @"      for(i = 0; i < wide.length; i ++)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        searchWordTmp = searchWordTmp.replace(wide[i], narrow[i]);" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"      var searchWord = searchWordTmp.split("" "");" + "\n";
                    searchJs += @"      for(i = 0; i < searchWord.length; i ++)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        searchWord[i] = selectorEscape(searchWord[i].replace("">"", ""&gt;"").replace(""<"", ""&lt;""));" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"      var hilightWord = searchWord.join(""|"");" + "\n";
                    searchJs += @"      for(i = 0; i < hilight.length; i ++)" + "\n";
                    searchJs += @"      {" + "\n";
                    searchJs += @"        var reg = new RegExp(hilight[i], ""gm"");" + "\n";
                    searchJs += @"        hilightWord = hilightWord.replace(reg, hilight[i]);" + "\n";
                    searchJs += @"      }" + "\n";
                    searchJs += @"" + "\n";
                    searchJs += @"      var reg = new RegExp(""(""+hilightWord+"")(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"      var regnbsp = new RegExp(""&nbsp;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"      var reggt = new RegExp(""&gt;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"      var reglt = new RegExp(""&lt;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"      var regquot = new RegExp(""&quot;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"      var regamp = new RegExp(""&amp;(?=[^<>]*<)"", ""gm"");" + "\n";
                    searchJs += @"        $(""iframe.topic"").contents().find(""body"").html($(""iframe.topic"").contents().find(""body"").html().replace(regnbsp, ""　"").replace(reggt, "">"").replace(reglt, ""<"").replace(regquot, '""').replace(regamp, ""&"").replace(reg, ""<font class='keyword' style='color:rgb(0, 0, 0); background-color:rgb(252, 255, 0);'>$1</font>""));" + "\n";
                    searchJs += @"    }" + "\n";
                    searchJs += @"  });" + "\n";
                    searchJs += @"});" + "\n";
                    #endregion
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
                                //else
                                //{
                                objBodyCurrent.AppendChild(objBody.CreateElement("p"));
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("class", "Heading3");
                                ((System.Xml.XmlElement)objBodyCurrent.LastChild).SetAttribute("id", Regex.Replace(setid, "^.*?♯(.*?)$", "$1"));

                                foreach (System.Xml.XmlNode childItem in childs.ChildNodes)
                                {
                                    innerNode(styleName, objBodyCurrent.LastChild, childItem);
                                }
                                //}
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

            WordAddIn1.Globals.ThisAddIn.Application.ActiveWindow.View.Type = defaultView;
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

            //ActiveDocumentのパスは、「WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.Path」で取得できます。
            //index.htmlが出力されるとして、「WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.Path + @"\index.html"」に
            //出力されるindex.htmlのパスという想定で、以下に出力後のHTMLをブラウザで閲覧するか否かの
            //メッセージボックス表示のコードを書いています。


            //DialogResult selectMess = MessageBox.Show(WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.Path + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？", "HTML出力成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            //if (selectMess == DialogResult.Yes)
            //{
            //    try
            //    {
            //        Process.Start(WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.Path + @"\index.html");
            //    }
            //    catch
            //    {
            //        MessageBox.Show("HTMLの出力に失敗しました。", "HTML出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //    }
            //}



            /*以下は、次期対応変更履歴保存用コードの一部です。
            var activeDoc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument as Microsoft.Office.Interop.Word.Document;
            Word.Selection ws = WordAddIn1.Globals.ThisAddIn.Application.Selection;
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
            WordAddIn1.Globals.ThisAddIn.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }
    }
}
