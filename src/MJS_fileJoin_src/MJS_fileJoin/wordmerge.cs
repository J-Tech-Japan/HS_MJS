using System.Collections.Generic;
using System.Reflection;
using Word = Microsoft.Office.Interop.Word;
using System.Windows.Forms;
using MJS_fileJoin;
using System.Text.RegularExpressions;
using System.Linq;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        public void Merge(string strOrgDoc, string[] arrCopies, string strOutDoc, MainForm form, bool check1, bool check2, bool check3, object nothing)
        {
            Word.Application objApp = null;
            MainForm fm = form;
            object objMissing = Missing.Value;
            object objFalse = false;
            object objTarget = Word.WdMergeTarget.wdMergeTargetSelected;
            object objUseFormatFrom = Word.WdUseFormattingFrom.wdFormattingFromSelected;
            object objDetectFromChanges = Word.WdRevisionType.wdNoRevision;

            int chapCnt;

            try
            {
                objApp = new Word.Application();
                objApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;
                objApp.Options.CheckGrammarAsYouType = false;
                objApp.Options.CheckGrammarWithSpelling = false;
                objApp.Options.CheckSpellingAsYouType = false;
                objApp.Options.ShowReadabilityStatistics = false;
                object objOrgDoc = strOrgDoc;
                objApp.Visible = false;
                Word.Document objDocLast = null;

                objDocLast = objApp.Documents.Open(
                  ref objOrgDoc,    //FileName
                  ref objMissing,   //ConfirmVersions
                  ref objMissing,   //ReadOnly
                  ref objMissing,   //AddToRecentFiles
                  ref objMissing,   //PasswordDocument
                  ref objMissing,   //PasswordTemplate
                  ref objMissing,   //Revert
                  ref objMissing,   //WritePasswordDocument
                  ref objMissing,   //WritePasswordTemplate
                  ref objMissing,   //Format
                  ref objMissing,   //Enconding
                  ref objMissing,   //Visible
                  ref objMissing,   //OpenAndRepair
                  ref objMissing,   //DocumentDirection
                  ref objMissing,   //NoEncodingDialog
                  ref objMissing    //XMLTransform
                  );

                chapCnt = objDocLast.Sections.Count;
                Dictionary<int, string> dic1 = new Dictionary<int, string>();
                Dictionary<int, int> dic2 = new Dictionary<int, int>();
                dic1.Add(0, Regex.Replace(strOrgDoc, @"^.*?@([^\\]*?)\\.*?$", "$1"));
                dic2.Add(0, 0);
                fm.progressBar1.Maximum = arrCopies.Length + 1;
                fm.progressBar1.Value = 1;
                int chapCntLast = 0;

                foreach (string strCopy in arrCopies)
                {
                    objApp.Selection.EndKey(Word.WdUnits.wdStory);
                    objApp.Selection.HomeKey(Word.WdUnits.wdStory);
                    objApp.Selection.EndKey(Word.WdUnits.wdStory);
                    Application.DoEvents();

                    objApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);
                    Application.DoEvents();

                    chapCntLast = objDocLast.Sections.Count;
                    objApp.Selection.InsertFile(strCopy, ref objMissing, objMissing, objMissing, objMissing);

                    fm.progressBar1.Increment(1);
                }

                object objOutDoc = strOutDoc;

                if (!check3)
                {
                    fm.label10.Text = "重複箇所削除中...";
                    fm.progressBar1.Maximum = 11;
                    fm.progressBar1.Value = 1;


                    string[] lsStyleName = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次", "奥付タイトル", "索引見出し" };
                    foreach (string styleName in lsStyleName)
                    {
                        object styleObject = styleName;
                        for (int i = chapCnt + 1; i <= chapCntLast; i++)
                        {
                            Word.Range wr = objDocLast.Sections[i].Range;
                            wr.Find.ClearFormatting();
                            wr.Find.set_Style(ref styleObject);
                            //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                            wr.Find.Execute();
                            if (wr.Find.Found)
                            {
                                objDocLast.Sections[i].Range.Delete();
                                i--;
                                chapCntLast--;
                            }
                        }
                        fm.progressBar1.Increment(1);
                    }

                    bool last = false;
                    string[] indexItems = { "索引見出し" };
                    foreach (string styleName in indexItems)
                    {
                        object styleObject = styleName;
                        int allChap = objDocLast.Sections.Count;
                        for (int i = allChap; i > chapCntLast; i--)
                        {
                            Word.Range wr = objDocLast.Sections[i].Range;
                            wr.Find.ClearFormatting();
                            wr.Find.set_Style(ref styleObject);
                            //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                            wr.Find.Execute();
                            if (wr.Find.Found)
                            {
                                last = true;
                                break;
                            }
                        }
                        fm.progressBar1.Increment(1);
                    }

                    fm.label10.Text = "章扉章節項番号修正中...";
                    fm.progressBar1.Maximum = objDocLast.Sections.Count;
                    fm.progressBar1.Value = 1;

                    string[] lastItems = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次" };
                    foreach (string styleName in lastItems)
                    {
                        object styleObject = styleName;
                        int allChap = objDocLast.Sections.Count;
                        for (int i = allChap; i > chapCntLast; i--)
                        {
                            try
                            {
                                Word.Range wr = objDocLast.Sections[i].Range;
                                wr.Find.ClearFormatting();
                                wr.Find.set_Style(ref styleObject);
                                //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                                wr.Find.Execute();
                                if (wr.Find.Found)
                                {
                                    if (last)
                                        last = false;
                                    else
                                    {
                                        objDocLast.Sections[i].Range.Delete();
                                        i--;
                                        chapCntLast--;
                                    }
                                }
                            }
                            catch
                            {
                                break;
                            }
                        }
                    }

                    string[] chapFrontItems = { "MJS_章扉-タイトル" };
                    foreach (string styleName in chapFrontItems)
                    {
                        object styleObject = styleName;
                        object shouTobira = "MJS_章扉-目次1";
                        int allChap = objDocLast.Sections.Count;
                        for (int i = 1; i <= allChap; i++)
                        {
                            try
                            {
                                Word.Range wr = objDocLast.Sections[i].Range;
                                wr.Find.ClearFormatting();
                                wr.Find.set_Style(ref styleObject);
                                //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                                wr.Find.Execute();
                                if (wr.Find.Found)
                                {
                                    int shou = 0;
                                    for (int p = 1; p <= objDocLast.Sections[i].Range.Paragraphs.Count - 1; p++)
                                    {
                                        if (objDocLast.Sections[i].Range.Paragraphs[p].get_Style().NameLocal.Trim() == "MJS_章扉-タイトル")
                                            shou = objDocLast.Sections[i].Range.Paragraphs[p].Range.ListFormat.ListValue;
                                        else if (objDocLast.Sections[i].Range.Paragraphs[p].get_Style().NameLocal.Trim().Contains("MJS_章扉-目次1"))
                                        {
                                            objDocLast.Sections[i].Range.Paragraphs[p].Range.Text = Regex.Replace(objDocLast.Sections[i].Range.Paragraphs[p].Range.Text, @"^\d+?", shou.ToString());
                                            objDocLast.Sections[i].Range.Paragraphs[p].Range.set_Style(ref shouTobira);
                                        }
                                    }
                                }
                            }
                            catch
                            { }
                            fm.progressBar1.Increment(1);
                        }
                    }

                    List<string> styleNames = new List<string>();
                    styleNames.Add("MJS_章扉-タイトル");
                    styleNames.Add("見出し 1,MJS_見出し 1");
                    styleNames.Add("見出し 2,MJS_見出し 2");
                    styleNames.Add("見出し 3,MJS_見出し 3");

                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].NumberFormat = "第%1章";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].NumberPosition = objApp.MillimetersToPoints(0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].TextPosition = objApp.MillimetersToPoints(5.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingNone;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].ResetOnHigher = 0;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].StartAt = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Font.Bold = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Font.Italic = 0;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Font.Color = Word.WdColor.wdColorAutomatic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Font.Size = 60;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].Font.Name = "メイリオ";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1].LinkedStyle = "MJS_章扉-タイトル";

                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].NumberFormat = "%1.%2";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].NumberPosition = objApp.MillimetersToPoints(1.5F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].TextPosition = objApp.MillimetersToPoints(20.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].TabPosition = objApp.MillimetersToPoints(20.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].ResetOnHigher = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].StartAt = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Font.Bold = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Font.Italic = 0;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Font.Color = Word.WdColor.wdColorAutomatic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Font.Size = 16;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].Font.Name = "メイリオ";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2].LinkedStyle = "見出し 1";

                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].NumberFormat = "%1.%2.%3";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].NumberPosition = objApp.MillimetersToPoints(0.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].TextPosition = objApp.MillimetersToPoints(20.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].TabPosition = objApp.MillimetersToPoints(20.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].ResetOnHigher = 2;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].StartAt = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Font.Bold = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Font.Italic = 0;
                    //Color mycolor = Color.FromArgb(31, 73, 125);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Font.Color = (Word.WdColor)(mycolor.R + 0x100 * mycolor.G + 0x10000 * mycolor.B);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Font.Size = 14;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].Font.Name = "メイリオ";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3].LinkedStyle = "見出し 2";

                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].NumberFormat = "%1.%2.%3.%4";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].NumberPosition = objApp.MillimetersToPoints(7.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].TextPosition = objApp.MillimetersToPoints(28.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].TabPosition = objApp.MillimetersToPoints(28.0F);
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].ResetOnHigher = 3;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].StartAt = 1;
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].Font.Name = "メイリオ";
                    //objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4].LinkedStyle = "見出し 3";

                    SetOutlineNumberingFormat(objApp, objDocLast, fm);

                    int first = 0;
                    int second = 0;
                    int third = 0;
                    int fourth = 0;

                    for (int i = 1; i <= objDocLast.ListParagraphs.Count; i++)
                    {
                        fm.progressBar1.Increment(1);
                        if (!Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"第.*?章") && !Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"\d\.\d")) continue;
                        if (Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"第.*?章"))
                        {
                            first++;
                            second = 0;
                            third = 0;
                            fourth = 0;
                            if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != first)
                                objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                        }
                        else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 2)
                        {
                            second++;
                            third = 0;
                            fourth = 0;
                            if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != second)
                                objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                        }
                        else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 3)
                        {
                            third++;
                            fourth = 0;
                            if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != third)
                                objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                        }
                        else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 4)
                        {
                            fourth++;
                            if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != fourth)
                                objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                        }
                    }

                    fm.label10.Text = "索引更新中...";
                    object styleObject3 = "索引見出し";
                    bool bl3 = false;
                    int secCnt3 = objDocLast.Sections.Count;

                    for (int i = secCnt3; i > 0; i--)
                    {
                        Word.Range wr = objDocLast.Sections[i].Range;
                        wr.Find.ClearFormatting();
                        wr.Find.set_Style(ref styleObject3);
                        wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
                        wr.Find.Execute();
                        if (wr.Find.Found)
                        {
                            if (bl3)
                            {
                                objDocLast.Sections[i].Range.Delete();
                                i--;
                            }
                            else
                                bl3 = true;
                        }
                    }

                    fm.label10.Text = "目次更新中...";
                    fm.progressBar1.Value = 0;
                    fm.progressBar1.Maximum = 1;

                    if (objDocLast.TablesOfContents.Count >= 1)
                        objDocLast.TablesOfContents[1].Update();
                    fm.progressBar1.Value = 1;

                    fm.label10.Text = "索引更新中...";
                    fm.progressBar1.Value = 0;
                    fm.progressBar1.Maximum = 1;
                    if (objDocLast.Indexes.Count >= 1)
                        objDocLast.TablesOfContents[1].Update();
                    fm.progressBar1.Value = 1;
                }

                fm.label10.Text = "ハイパーリンク更新中...";
                List<string> ls = new List<string>();

                foreach (Word.Bookmark wb in objDocLast.Bookmarks)
                    ls.Add(wb.Name);

                fm.progressBar1.Value = 0;
                fm.progressBar1.Maximum = objDocLast.Fields.Count;

                foreach (Word.Field wf in objDocLast.Fields)
                {
                    fm.progressBar1.Increment(1);
                    if (wf.Type == Word.WdFieldType.wdFieldHyperlink)
                    {
                        if (wf.Code.Text.Contains("\"http")) continue;
                        string text = "";
                        if (!wf.Code.Text.Contains(@"\l"))
                            text = Regex.Replace(wf.Code.Text, @".*?""([^""]*?)"".*?", "$1");
                        else
                        {
                            if (!Regex.IsMatch(wf.Code.Text, @".*?""[^""]*?"".*?""[^""]*?"".*?")) continue;
                            text = Regex.Replace(wf.Code.Text, @".*?""([^""]*?)"".*?""([^""]*?)"".*?", "$1#$2");
                        }

                        string[] subtext = text.Split('\\');
                        text = subtext[subtext.Count() - 1];
                        subtext = text.Split('/');
                        text = subtext[subtext.Count() - 1];
                        if (ls.Contains(text.Replace(".html", "").Replace("#", "♯").Trim()))
                        {
                            wf.Code.Text = @"HYPERLINK \l """ + text.Replace(".html", "").Replace("#", "♯").Trim() + @"""";
                            wf.Update();
                        }
                        else
                            wf.Unlink();
                    }
                }

                foreach (Word.Hyperlink wh in objDocLast.Hyperlinks)
                {
                    if (Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}") ||
                        Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}♯\w{3}\d{5}"))
                        wh.TextToDisplay = Regex.Replace(wh.TextToDisplay, @".*(\d+\.)+\d+[\s　\t]", "");
                }

                if (check2)
                {
                    fm.label10.Text = "PDF出力中...";

                    objDocLast.ExportAsFixedFormat(
                        strOutDoc.Replace(".doc", ".pdf"),
                        Word.WdExportFormat.wdExportFormatPDF,
                        false,
                        Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                        Word.WdExportRange.wdExportAllDocument,
                        1,
                        1,
                        Word.WdExportItem.wdExportDocumentContent,
                        false,
                        true,
                        Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                        false,
                        true,
                        false
                        );
                }

                if (check1)
                {
                    fm.label10.Text = "Word保存中...";
                    objDocLast.SaveAs(
                      ref objOutDoc,      //FileName
                      ref objMissing,     //FileFormat
                      ref objMissing,     //LockComments
                      ref objMissing,     //PassWord     
                      ref objMissing,     //AddToRecentFiles
                      ref objMissing,     //WritePassword
                      ref objMissing,     //ReadOnlyRecommended
                      ref objMissing,     //EmbedTrueTypeFonts
                      ref objMissing,     //SaveNativePictureFormat
                      ref objMissing,     //SaveFormsData
                      ref objMissing,     //SaveAsAOCELetter,
                      ref objMissing,     //Encoding
                      ref objMissing,     //InsertLineBreaks
                      ref objMissing,     //AllowSubstitutions
                      ref objMissing,     //LineEnding
                      ref objMissing      //AddBiDiMarks
                      );
                }

                foreach (Word.Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }

            }
            finally
            {
                objApp.Quit(
                  ref objMissing,     //SaveChanges
                  ref objMissing,     //OriginalFormat
                  ref objMissing      //RoutDocument
                  );
                objApp = null;
            }
        }

        public void Merge(string strOrgDoc, List<string> strCopyFolder, string strOutDoc, MainForm fm, bool check1, bool check2, bool check3)
        {
            MainForm form = fm;
            string[] arrFiles = strCopyFolder.ToArray();
            Merge(strOrgDoc, arrFiles, strOutDoc, form, check1, check2, check3, null);
        }
    }
}
