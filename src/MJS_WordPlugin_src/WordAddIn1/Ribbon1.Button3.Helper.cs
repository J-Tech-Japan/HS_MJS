using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        // 表紙に関連する段落を収集
        public void CollectCoverParagraphs(Document docCopy, ref string manualTitle, ref string manualSubTitle, ref string manualVersion,
                                      ref string manualTitleCenter, ref string manualSubTitleCenter, ref string manualVersionCenter,
                                      ref bool coverExist)
        {
            foreach (Paragraph wp in docCopy.Sections[1].Range.Paragraphs)
            {
                if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualTitle += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualSubTitle += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualVersion += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルタイトル（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualTitleCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルサブタイトル（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualSubTitleCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
                }
                else if (wp.get_Style().NameLocal == "MJS_マニュアルバージョン（中央）")
                {
                    if (!string.IsNullOrEmpty(wp.Range.Text.Trim()) && wp.Range.Text.Trim() != "/")
                    {
                        manualVersionCenter += wp.Range.Text + "<br/>";
                        coverExist = true;
                    }
                    continue;
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
            ref string trademarkRight,
            ref bool isTradeMarksDetected,
            ref bool isRightDetected)
        {
            foreach (Paragraph wp in docCopy.Sections[lastSectionIdx].Range.Paragraphs)
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
        }

        // 表紙選択ダイアログを表示し、選択されたテンプレートに応じてフラグを設定
        public bool HandleCoverSelection(loader load, out bool isEasyCloud, out bool isEdgeTracker, out bool isPattern1, out bool isPattern2)
        {
            isEasyCloud = false;
            isEdgeTracker = false;
            isPattern1 = false;
            isPattern2 = false;

            CoverSelectionForm coverSelectionForm = new CoverSelectionForm();
            load.Visible = false;
            coverSelectionForm.ShowDialog();

            if (coverSelectionForm.DialogResult != DialogResult.OK)
            {
                load.Close();
                load.Dispose();
                return false;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.None)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("表紙のパターンをを選択してください。");
                return false;
            }

            if (coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern3)
            {
                load.Close();
                load.Dispose();
                MessageBox.Show("[汎用パターン3]テンプレートはまもなく登場します。");
                return false;
            }

            isEasyCloud = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EasyCloud;
            isEdgeTracker = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.EdgeTracker;
            isPattern1 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern1;
            isPattern2 = coverSelectionForm.SelectedCoverTemplate == CoverSelectionForm.CoverTemplateEnum.GeneralPattern2;

            return true;
        }

        // キャンバスに関連する図形のプロパティを調整
        public void AdjustCanvasShapes(Document docCopy)
        {
            foreach (Shape docS in docCopy.Shapes)
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

                    docS.LockAspectRatio = Microsoft.Office.Core.MsoTriState.msoFalse;
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
        }
    }
}
