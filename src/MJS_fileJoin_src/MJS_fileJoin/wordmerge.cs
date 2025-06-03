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
            Word.Document objDocLast = null;

            int chapCnt;

            try
            {
                (objApp,objDocLast) = CreateAndOpenWordDocument(strOrgDoc, objMissing);

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
                    // スタイル名で重複しているセクションを削除する
                    RemoveDuplicateSectionsByStyle(objDocLast, fm, chapCnt, ref chapCntLast);

                    // 末尾に索引セクションが存在するかチェックする
                    bool last = CheckIndexSection(objDocLast, fm, chapCntLast);

                    // 末尾の不要なセクション（スタイル指定）を削除
                    RemoveLastSectionsByStyle(objDocLast, fm, chapCntLast, ref last);

                    // 章扉の章番号を修正する
                    UpdateChapterFrontNumbers(objDocLast, fm);

                    // アウトライン番号（章・節番号）を再適用する
                    ApplyOutlineNumbering(objApp, objDocLast, fm);

                    // 目次・索引を更新する
                    UpdateIndexAndToc(objDocLast, fm);
                }

                // ハイパーリンクの更新処理
                UpdateHyperlinks(objDocLast, fm);

                //fm.label10.Text = "ハイパーリンク更新中...";
                //List<string> ls = new List<string>();

                //foreach (Word.Bookmark wb in objDocLast.Bookmarks)
                //    ls.Add(wb.Name);

                //fm.progressBar1.Value = 0;
                //fm.progressBar1.Maximum = objDocLast.Fields.Count;

                //foreach (Word.Field wf in objDocLast.Fields)
                //{
                //    fm.progressBar1.Increment(1);
                //    if (wf.Type == Word.WdFieldType.wdFieldHyperlink)
                //    {
                //        if (wf.Code.Text.Contains("\"http")) continue;
                //        string text = "";
                //        if (!wf.Code.Text.Contains(@"\l"))
                //            text = Regex.Replace(wf.Code.Text, @".*?""([^""]*?)"".*?", "$1");
                //        else
                //        {
                //            if (!Regex.IsMatch(wf.Code.Text, @".*?""[^""]*?"".*?""[^""]*?"".*?")) continue;
                //            text = Regex.Replace(wf.Code.Text, @".*?""([^""]*?)"".*?""([^""]*?)"".*?", "$1#$2");
                //        }

                //        string[] subtext = text.Split('\\');
                //        text = subtext[subtext.Count() - 1];
                //        subtext = text.Split('/');
                //        text = subtext[subtext.Count() - 1];
                //        if (ls.Contains(text.Replace(".html", "").Replace("#", "♯").Trim()))
                //        {
                //            wf.Code.Text = @"HYPERLINK \l """ + text.Replace(".html", "").Replace("#", "♯").Trim() + @"""";
                //            wf.Update();
                //        }
                //        else
                //            wf.Unlink();
                //    }
                //}

                //foreach (Word.Hyperlink wh in objDocLast.Hyperlinks)
                //{
                //    if (Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}") ||
                //        Regex.IsMatch(wh.Name.Trim(), @"^\w{3}\d{5}♯\w{3}\d{5}"))
                //        wh.TextToDisplay = Regex.Replace(wh.TextToDisplay, @".*(\d+\.)+\d+[\s　\t]", "");
                //}

                //if (check2)
                //{
                //    fm.label10.Text = "PDF出力中...";

                //    objDocLast.ExportAsFixedFormat(
                //        strOutDoc.Replace(".doc", ".pdf"),
                //        Word.WdExportFormat.wdExportFormatPDF,
                //        false,
                //        Word.WdExportOptimizeFor.wdExportOptimizeForPrint,
                //        Word.WdExportRange.wdExportAllDocument,
                //        1,
                //        1,
                //        Word.WdExportItem.wdExportDocumentContent,
                //        false,
                //        true,
                //        Word.WdExportCreateBookmarks.wdExportCreateHeadingBookmarks,
                //        false,
                //        true,
                //        false
                //        );
                //}

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

        private void RemoveDuplicateSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCnt, ref int chapCntLast)
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

                    if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                    {
                        wr.Find.set_Style(ref styleObject);
                    }
                    else
                    {
                        // スタイルがなければスキップやログ出力
                    }

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
        }

        private bool CheckIndexSection(Word.Document objDocLast, MainForm fm, int chapCntLast)
        {
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
                    if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                    {
                        wr.Find.set_Style(ref styleObject);
                    }
                    else
                    {
                        // スタイルがなければスキップやログ出力
                    }
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
            return last;
        }

        private void RemoveLastSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCntLast, ref bool last)
        {
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
                        if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                        {
                            wr.Find.set_Style(ref styleObject);
                        }
                        else
                        {
                            // スタイルがなければスキップやログ出力
                        }
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
        }

        private void UpdateChapterFrontNumbers(Word.Document objDocLast, MainForm fm)
        {
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
                        if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                        {
                            wr.Find.set_Style(ref styleObject);
                        }
                        else
                        {
                            // スタイルがなければスキップやログ出力
                        }
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
        }

        private void ApplyOutlineNumbering(Word.Application objApp, Word.Document objDocLast, MainForm fm)
        {
            List<string> styleNames = new List<string>();
            styleNames.Add("MJS_章扉-タイトル");
            styleNames.Add("見出し 1,MJS_見出し 1");
            styleNames.Add("見出し 2,MJS_見出し 2");
            styleNames.Add("見出し 3,MJS_見出し 3");

            // スタイルのアウトライン番号を設定
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
        }

        private void UpdateIndexAndToc(Word.Document objDocLast, MainForm fm)
        {
            fm.label10.Text = "索引更新中...";
            object styleObject3 = "索引見出し";
            bool bl3 = false;
            int secCnt3 = objDocLast.Sections.Count;

            for (int i = secCnt3; i > 0; i--)
            {
                Word.Range wr = objDocLast.Sections[i].Range;
                wr.Find.ClearFormatting();
                if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == "索引見出し"))
                {
                    wr.Find.set_Style(ref styleObject3);
                }
                else
                {
                    // スタイルがなければスキップやログ出力
                }
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
    }
}
