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
                (objApp,objDocLast) = CreateWordDocument(strOrgDoc, objMissing);

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

        // Merge をラップするメソッド
        // List<string> 型のフォルダ（またはファイル）リストを string[] 配列に変換し、
        // その他の引数とともに Merge に渡す
        public void MergeFromFolders(string strOrgDoc, List<string> strCopyFolder, string strOutDoc, MainForm fm, bool check1, bool check2, bool check3)
        {
            MainForm form = fm;
            string[] arrFiles = strCopyFolder.ToArray();
            Merge(strOrgDoc, arrFiles, strOutDoc, form, check1, check2, check3, null);
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
    }
}
