using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Windows.Forms;
using MJS_fileJoin;
using System;
using System.Text.RegularExpressions;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        public void Merge(string strOrgDoc, string[] arrCopies, string strOutDoc, MainForm form, bool check1, bool check2, bool check3, object nothing)
        {
            Word.Application objApp = null;
            Word.Document objDocLast = null;
            object objMissing = Type.Missing;
            object objFalse = false;
            int chapCnt;

            try
            {
                InitializeWordAndOpenDocument(strOrgDoc, ref objApp, ref objDocLast);

                chapCnt = objDocLast.Sections.Count;
                Dictionary<int, string> dic1 = new Dictionary<int, string>();
                Dictionary<int, int> dic2 = new Dictionary<int, int>();
                dic1.Add(0, Regex.Replace(strOrgDoc, @"^.*?@([^\\]*?)\\.*?$", "$1"));
                dic2.Add(0, 0);
                form.progressBar1.Maximum = arrCopies.Length + 1;
                form.progressBar1.Value = 1;
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

                    form.progressBar1.Increment(1);
                }

                object objOutDoc = strOutDoc;

                if (!check3)
                {
                    string[] lsStyleName = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次", "奥付タイトル", "索引見出し" };

                    // 指定したスタイル名のセクションを削除する
                    RemoveSectionsInRangeByStyle(objDocLast, lsStyleName, chapCnt, ref chapCntLast, form);

                    bool last = false;
                    string[] indexItems = { "索引見出し" };

                    // 指定したスタイル名が見つかったらlastフラグをtrueにして進捗バーを進める
                    SetLastFlagIfStyleFound(objDocLast, indexItems, ref last, chapCntLast, form);

                    string[] lastItems = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次" };
                    
                    // 後方からlastフラグ付きで削除（最後のセクションは削除しない）
                    RemoveSectionsFromEndByStyleWithLastFlag(objDocLast, lastItems, ref chapCntLast, ref last, form);

                    // 章扉の項番号を修正
                    UpdateChapterFrontNumbers(objDocLast, form);

                    List<string> styleNames = new List<string>();
                    styleNames.Add("MJS_章扉-タイトル");
                    styleNames.Add("見出し 1,MJS_見出し 1");
                    styleNames.Add("見出し 2,MJS_見出し 2");
                    styleNames.Add("見出し 3,MJS_見出し 3");

                    Color mycolor = Color.FromArgb(31, 73, 125);

                    // Wordのアウトライン番号書式を設定
                    SetOutlineNumberingFormat(objApp, mycolor);

                    // 段落のアウトライン番号を修正
                    FixOutlineNumbering(objDocLast, objApp, form);

                    // 指定したスタイル名のセクションを後方から1つだけ残して削除
                    RemoveSectionsByStyleKeepLast(objDocLast, "索引見出し", form);

                    // 目次と索引の更新処理
                    UpdateTocAndIndex(objDocLast, form);
                }

                // 「HYPERLINK _Ref...」形式のフィールドを「REF ... \h」形式に変換
                List<string> targetStyles = new List<string> { "MJS_参照先" };
                ConvertHyperlinkToRef(objDocLast, targetStyles);

                // ハイパーリンクの更新
                UpdateHyperlinks(objDocLast, form);

                if (check2)
                {
                    form.label10.Text = "PDF出力中...";

                    // PDFとして出力
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
                    form.label10.Text = "Word保存中...";

                    // Wordとして保存
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
        // List<string> 型のフォルダ（またはファイル）リストを配列に変換し、
        // 他の引数とともに Merge に渡す
        public void MergeFromFolders(string strOrgDoc, List<string> strCopyFolder, string strOutDoc, MainForm fm, bool check1, bool check2, bool check3)
        {
            MainForm form = fm;
            string[] arrFiles = strCopyFolder.ToArray();
            Merge(strOrgDoc, arrFiles, strOutDoc, form, check1, check2, check3, null);
        }
    }
}




