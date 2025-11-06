using System.Collections.Generic;
using Word = Microsoft.Office.Interop.Word;
using System.Drawing;
using System.Windows.Forms;
using MJS_fileJoin;
using System;
using System.Text.RegularExpressions;
using System.Diagnostics;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        public void Merge(string strOrgDoc, string[] arrCopies, string strOutDoc, MainForm form, bool check1, bool check3, object nothing)
        {
            // 全体の処理時間計測開始
            Stopwatch totalStopwatch = Stopwatch.StartNew();
            Trace.WriteLine("=================================================");
            Trace.WriteLine("=== Word結合処理 開始 ===");
            Trace.WriteLine($"元ドキュメント: {strOrgDoc}");
            Trace.WriteLine($"結合ファイル数: {arrCopies.Length}");
            Trace.WriteLine($"出力先: {strOutDoc}");
            Trace.WriteLine($"Word保存: {check1}, スキップモード: {check3}");
            Trace.WriteLine("=================================================");
            
            Word.Application objApp = null;
            Word.Document objDocLast = null;
            object objMissing = Type.Missing;
            object objFalse = false;
            int chapCnt;

            try
            {
                // Word初期化と文書オープン
                Stopwatch stepStopwatch = Stopwatch.StartNew();
                InitializeWordAndOpenDocument(strOrgDoc, ref objApp, ref objDocLast);
                stepStopwatch.Stop();
                Trace.WriteLine($"[1] Word初期化と文書オープン: {stepStopwatch.ElapsedMilliseconds}ms");

                chapCnt = objDocLast.Sections.Count;
                Dictionary<int, string> dic1 = new Dictionary<int, string>();
                Dictionary<int, int> dic2 = new Dictionary<int, int>();
                dic1.Add(0, Regex.Replace(strOrgDoc, @"^.*?@([^\\]*?)\\.*?$", "$1"));
                dic2.Add(0, 0);
                form.progressBar1.Maximum = arrCopies.Length + 1;
                form.progressBar1.Value = 1;
                int chapCntLast = 0;

                // ファイル結合処理
                stepStopwatch = Stopwatch.StartNew();
                int fileIndex = 0;
                foreach (string strCopy in arrCopies)
                {
                    Stopwatch fileStopwatch = Stopwatch.StartNew();
                    
                    objApp.Selection.EndKey(Word.WdUnits.wdStory);
                    objApp.Selection.HomeKey(Word.WdUnits.wdStory);
                    objApp.Selection.EndKey(Word.WdUnits.wdStory);

                    Application.DoEvents();

                    objApp.Selection.InsertBreak(Word.WdBreakType.wdSectionBreakNextPage);

                    Application.DoEvents();

                    chapCntLast = objDocLast.Sections.Count;
                    objApp.Selection.InsertFile(strCopy, ref objMissing, objMissing, objMissing, objMissing);

                    form.progressBar1.Increment(1);
                    
                    fileStopwatch.Stop();
                    fileIndex++;
                    if (fileIndex % 5 == 0 || fileStopwatch.ElapsedMilliseconds > 1000)
                    {
                        Trace.WriteLine($"  ファイル結合 {fileIndex}/{arrCopies.Length}: {fileStopwatch.ElapsedMilliseconds}ms - {System.IO.Path.GetFileName(strCopy)}");
                    }
                }
                stepStopwatch.Stop();
                Trace.WriteLine($"[2] ファイル結合処理（全{arrCopies.Length}ファイル）: {stepStopwatch.ElapsedMilliseconds}ms ({stepStopwatch.Elapsed.TotalSeconds:F2}秒)");

                // ページ番号を通し番号に設定
                stepStopwatch = Stopwatch.StartNew();
                Utils.ResetPageNumbering(objDocLast, form);
                stepStopwatch.Stop();
                Trace.WriteLine($"[2.5] ページ番号通し番号設定: {stepStopwatch.ElapsedMilliseconds}ms");

                object objOutDoc = strOutDoc;

                if (!check3)
                {
                    Trace.WriteLine("--- 詳細処理開始 ---");
                    
                    // セクション削除処理1
                    stepStopwatch = Stopwatch.StartNew();
                    string[] lsStyleName = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次", "奥付タイトル", "索引見出し" };
                    RemoveSectionsInRangeByStyle(objDocLast, lsStyleName, chapCnt, ref chapCntLast, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[3] セクション削除（範囲指定）: {stepStopwatch.ElapsedMilliseconds}ms");

                    // 索引見出し検索
                    stepStopwatch = Stopwatch.StartNew();
                    bool last = false;
                    string[] indexItems = { "索引見出し" };
                    SetLastFlagIfStyleFound(objDocLast, indexItems, ref last, chapCntLast, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[4] 索引見出し検索: {stepStopwatch.ElapsedMilliseconds}ms");

                    // セクション削除処理2
                    stepStopwatch = Stopwatch.StartNew();
                    string[] lastItems = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次" };
                    RemoveSectionsFromEndByStyleWithLastFlag(objDocLast, lastItems, ref chapCntLast, ref last, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[5] セクション削除（後方から）: {stepStopwatch.ElapsedMilliseconds}ms");

                    // 章扉の項番号を修正
                    stepStopwatch = Stopwatch.StartNew();
                    UpdateChapterFrontNumbers(objDocLast, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[6] 章扉番号修正: {stepStopwatch.ElapsedMilliseconds}ms ({stepStopwatch.Elapsed.TotalSeconds:F2}秒)");

                    List<string> styleNames = new List<string>();
                    styleNames.Add("MJS_章扉-タイトル");
                    styleNames.Add("見出し 1,MJS_見出し 1");
                    styleNames.Add("見出し 2,MJS_見出し 2");
                    styleNames.Add("見出し 3,MJS_見出し 3");

                    Color mycolor = Color.FromArgb(31, 73, 125);

                    // アウトライン番号書式設定
                    stepStopwatch = Stopwatch.StartNew();
                    SetOutlineNumberingFormat(objApp, mycolor);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[7] アウトライン番号書式設定: {stepStopwatch.ElapsedMilliseconds}ms");

                    // アウトライン番号修正
                    stepStopwatch = Stopwatch.StartNew();
                    FixOutlineNumbering(objDocLast, objApp, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[8] アウトライン番号修正: {stepStopwatch.ElapsedMilliseconds}ms ({stepStopwatch.Elapsed.TotalSeconds:F2}秒)");

                    // 索引セクション削除
                    stepStopwatch = Stopwatch.StartNew();
                    RemoveSectionsByStyleKeepLast(objDocLast, "索引見出し", form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[9] 索引セクション削除: {stepStopwatch.ElapsedMilliseconds}ms");

                    // 目次と索引の更新
                    stepStopwatch = Stopwatch.StartNew();
                    UpdateTocAndIndex(objDocLast, form);
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[10] 目次・索引更新: {stepStopwatch.ElapsedMilliseconds}ms ({stepStopwatch.Elapsed.TotalSeconds:F2}秒)");
                }
                else
                {
                    Trace.WriteLine("--- スキップモード: 詳細処理をスキップ ---");
                }

                // ハイパーリンク変換
                stepStopwatch = Stopwatch.StartNew();
                List<string> targetStyles = new List<string> { "MJS_参照先" };
                ConvertHyperlinkToRef(objDocLast, targetStyles);
                stepStopwatch.Stop();
                Trace.WriteLine($"[11] ハイパーリンク変換: {stepStopwatch.ElapsedMilliseconds}ms");

                // ハイパーリンク更新
                stepStopwatch = Stopwatch.StartNew();
                UpdateHyperlinks(objDocLast, form);
                stepStopwatch.Stop();
                Trace.WriteLine($"[12] ハイパーリンク更新: {stepStopwatch.ElapsedMilliseconds}ms");

                if (check1)
                {
                    form.label10.Text = "Word保存中...";

                    // Word保存
                    stepStopwatch = Stopwatch.StartNew();
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
                    stepStopwatch.Stop();
                    Trace.WriteLine($"[13] Word保存: {stepStopwatch.ElapsedMilliseconds}ms ({stepStopwatch.Elapsed.TotalSeconds:F2}秒)");
                }
                else
                {
                    Trace.WriteLine("[13] Word保存: スキップ");
                }

                // 文書クローズ
                stepStopwatch = Stopwatch.StartNew();
                foreach (Word.Document objDocument in objApp.Documents)
                {
                    objDocument.Close(
                      ref objFalse,     //SaveChanges
                      ref objMissing,   //OriginalFormat
                      ref objMissing    //RouteDocument
                      );
                }
                stepStopwatch.Stop();
                Trace.WriteLine($"[14] 文書クローズ: {stepStopwatch.ElapsedMilliseconds}ms");

            }
            catch (Exception ex)
            {
                Trace.WriteLine($"!!! エラー発生 !!!");
                Trace.WriteLine($"エラーメッセージ: {ex.Message}");
                Trace.WriteLine($"スタックトレース: {ex.StackTrace}");
                throw;
            }
            finally
            {
                Stopwatch cleanupStopwatch = Stopwatch.StartNew();
                if (objApp != null)
                {
                    objApp.Quit(
                        ref objMissing,
                        ref objMissing,
                        ref objMissing
                    );
                    objApp = null;
                }
                cleanupStopwatch.Stop();
                Trace.WriteLine($"[15] Word終了処理: {cleanupStopwatch.ElapsedMilliseconds}ms");
                
                totalStopwatch.Stop();
                Trace.WriteLine("=================================================");
                Trace.WriteLine($"=== Word結合処理 完了 ===");
                Trace.WriteLine($"総処理時間: {totalStopwatch.ElapsedMilliseconds}ms ({totalStopwatch.Elapsed.TotalMinutes:F2}分)");
                Trace.WriteLine("=================================================");
            }
        }

        // Merge をラップするメソッド
        // List<string> 型のフォルダ（またはファイル）リストを配列に変換し、
        // 他の引数とともに Merge に渡す
        public void MergeFromFolders(string strOrgDoc, List<string> strCopyFolder, string strOutDoc, MainForm fm, bool check1, bool check3)
        {
            MainForm form = fm;
            string[] arrFiles = strCopyFolder.ToArray();
            Merge(strOrgDoc, arrFiles, strOutDoc, form, check1, check3, null);
        }
    }
}




