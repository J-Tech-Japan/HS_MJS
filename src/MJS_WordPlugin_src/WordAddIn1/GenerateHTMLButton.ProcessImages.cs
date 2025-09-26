// GenerateHTMLButton.ProcessImages.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 最初のセクション内の図形に対して、図形グループをすべて解除
        private void UngroupAllShapesInFirstSection(Document docCopy, Word.Application application)
        {
            bool repeatUngroup = true;
            while (repeatUngroup)
            {
                repeatUngroup = false;
                foreach (Shape ws in docCopy.Shapes)
                {
                    ws.Select();
                    if (application.Selection.Information[WdInformation.wdActiveEndSectionNumber] == 1)
                    {
                        if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            ws.Ungroup();
                            repeatUngroup = true;
                        }
                    }
                }
            }
        }

        private void ProcessCanvasAndPictureShapesInFirstSection(Document docCopy, Word.Application application, ref string subTitle, ref int biCount, string strOutFileName, (string rootPath, string exportDir, string exportDirPath, string tmpHtmlPath, string tmpFolderForImagesSavedBySaveAs2Method, string logPath, string indexHtmlPath, string docTitle, string docid, string zipDirPath, string headerDir, string docFullName, string docName) paths)
        {
            foreach (Shape ws in docCopy.Shapes)
            {
                ws.Select();
                if (application.Selection.Information[WdInformation.wdActiveEndSectionNumber] == 1)
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
                            if (!Directory.Exists(Path.Combine(paths.rootPath, "tmpcoverpic")))
                                Directory.CreateDirectory(Path.Combine(paths.rootPath, "tmpcoverpic"));

                            strOutFileName = Path.Combine(paths.rootPath, "tmpcoverpic");
                            byte[] vData = (byte[])application.Selection.EnhMetaFileBits;
                            if (vData != null && vData.Length > 0)
                            {
                                using (var ms = new MemoryStream(vData))
                                using (var temp = Image.FromStream(ms))
                                {
                                    float aspectTemp = (float)temp.Width / (float)temp.Height;
                                    if (aspectTemp > 2.683 || aspectTemp < 2.681)
                                    {
                                        biCount++;
                                        temp.Save(Path.Combine(strOutFileName, biCount + ".png"), ImageFormat.Png);
                                    }
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
        }

        private void ConvertPictureShapesToInlineInFirstSection(Document docCopy, Word.Application application)
        {
            foreach (Shape ws in docCopy.Shapes)
            {
                ws.Select();
                if (application.Selection.Information[WdInformation.wdActiveEndSectionNumber] == 1)
                {
                    if (ws.Type == Microsoft.Office.Core.MsoShapeType.msoPicture)
                    {
                        ws.ConvertToInlineShape();
                    }
                }
            }
        }

        private void ExtractProductLogosPattern1Pattern2(Document docCopy, Word.Application application, string strOutFileName, List<List<string>> productSubLogoGroups)
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
                            application.Selection.CopyAsPicture();

                            using (Image img = Clipboard.GetImage())
                            {
                                if (img != null)
                                {
                                    img.Save(Path.Combine(strOutFileName, "product_logo_main.png"), ImageFormat.Png);
                                }
                            }
                            Clipboard.Clear();

                            break; //get first product main logo only
                        }
                    }
                    catch (Exception)
                    {
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
                            application.Selection.CopyAsPicture();

                            using (Image img = Clipboard.GetImage())
                            {
                                if (img != null)
                                {
                                    productSubLogoCount++;
                                    string subLogoFileName = string.Format("product_logo_sub{0}.png", productSubLogoCount);
                                    img.Save(Path.Combine(strOutFileName, subLogoFileName), ImageFormat.Png);
                                    productSubLogoFileNames.Add(subLogoFileName);
                                }
                            }
                            Clipboard.Clear();

                            if (productSubLogoCount == 3)
                            {
                                break; //get first 3 sub logos only
                            }
                        }

                        productSubLogoGroups.Add(productSubLogoFileNames);
                    }
                    catch (Exception)
                    {
                        //log.WriteLine("Error when extracting [MJS_製品ロゴ（サブ）]: " + ex.ToString());
                    }
                }
            }
        }

        /// <summary>
        /// Pattern3用の製品ロゴ（メイン・サブ）を抽出します
        /// </summary>
        /// <param name="docCopy">ドキュメントのコピー</param>
        /// <param name="application">Wordアプリケーション</param>
        /// <param name="strOutFileName">出力先フォルダパス</param>
        /// <param name="productSubLogoGroups">製品サブロゴグループのリスト</param>
        /// <param name="log">ログライター</param>
        private void ExtractProductLogosPattern3(Document docCopy, Word.Application application, string strOutFileName, List<List<string>> productSubLogoGroups)
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
                            application.Selection.CopyAsPicture();

                            using (Image img = Clipboard.GetImage())
                            {
                                if (img != null)
                                {
                                    img.Save(Path.Combine(strOutFileName, "product_logo_main.png"), ImageFormat.Png);
                                }
                            }
                            Clipboard.Clear();

                            break; //get first product main logo only
                        }
                    }
                    catch (Exception)
                    {
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
                            application.Selection.CopyAsPicture();

                            using (Image img = Clipboard.GetImage())
                            {
                                if (img != null)
                                {
                                    productSubLogoCount++;
                                    string subLogoFileName = string.Format("product_logo_sub{0}.png", productSubLogoCount);
                                    img.Save(Path.Combine(strOutFileName, subLogoFileName), ImageFormat.Png);
                                    productSubLogoFileNames.Add(subLogoFileName);
                                }
                            }
                            Clipboard.Clear();

                            if (productSubLogoCount == 3)
                            {
                                break; //get first 3 sub logos only
                            }
                        }

                        productSubLogoGroups.Add(productSubLogoFileNames);
                    }
                    catch (Exception)
                    {
                    }
                }
            }
        }

        /// <summary>
        /// デフォルトパターン用のインライン図形から画像を抽出します
        /// </summary>
        /// <param name="docCopy">ドキュメントのコピー</param>
        /// <param name="strOutFileName">出力先フォルダパス</param>
        /// <param name="biCount">画像カウンター（参照渡し）</param>
        private void ExtractInlineShapesDefaultPattern(Document docCopy, string strOutFileName, ref int biCount)
        {
            foreach (InlineShape wis in docCopy.Sections[1].Range.InlineShapes)
            {
                byte[] vData = (byte[])wis.Range.EnhMetaFileBits;

                if (vData != null && vData.Length > 0)
                {
                    using (var ms = new MemoryStream(vData))
                    using (var temp = Image.FromStream(ms))
                    {
                        float aspectTemp = (float)temp.Width / (float)temp.Height;
                        if ((float)temp.Height < 360) continue;
                        if (aspectTemp > 12.225 && aspectTemp < 12.226) continue;
                        if (aspectTemp > 2.681 && aspectTemp < 2.683) continue;
                        biCount++;
                        temp.Save(Path.Combine(strOutFileName, biCount + ".png"), ImageFormat.Png);
                    }
                }
            }
        }
    }
}
