using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool makeBookInfo(loader load, StreamWriter swLog = null)
        {
            Globals.ThisAddIn.Application.ScreenUpdating = false;
            Word.Document thisDocument = Globals.ThisAddIn.Application.ActiveDocument;

            if (!CheckFileName(thisDocument, load))
            {
                return false;
            }

            var selection = Globals.ThisAddIn.Application.Selection;
            int selStart = selection.Start;
            int selEnd = selection.End;
            string rootPath = thisDocument.Path;
            string docName = thisDocument.Name;
            string headerDir = "headerFile";
            string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");
            string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");

            selection.EndKey(Word.WdUnits.wdStory);
            Application.DoEvents();

            selection.HomeKey(Word.WdUnits.wdStory);
            Application.DoEvents();

            if (selection.Type == Word.WdSelectionType.wdSelectionInlineShape ||
                selection.Type == Word.WdSelectionType.wdSelectionShape)
            {
                selection.MoveLeft(Word.WdUnits.wdCharacter);
            }

            bookInfoDef = "";
            Word.Document Doc = Globals.ThisAddIn.Application.ActiveDocument;

            // 書誌情報番号・書誌情報番号最大値
            int bibNum = 0;
            int bibMaxNum = 0;

            bool checkBL = false;

            if (!CheckAndLoadHeaderFile(Doc, load, ref oldInfo, ref newInfo, ref checkResult, ref bibNum, ref bibMaxNum, ref bookInfoDef, ref checkBL))
            {
                return false;
            }

            bookInfoDic.Clear();

            StreamWriter log = swLog;

            if (swLog == null)
            {
                log = new StreamWriter(rootPath + "\\log.txt", false, Encoding.UTF8);
            }

            try
            {
                
                if (!EnsureBookInfoDef(thisDocument, ref bookInfoDef, log, rootPath))
                {
                    return false;
                }

                Dictionary<string, string> oldBookInfoDic = new Dictionary<string, string>();
                HashSet<string> ls = new HashSet<string>();

                if (!Directory.Exists(rootPath + "\\" + headerDir))
                {
                    Directory.CreateDirectory(rootPath + "\\" + headerDir);
                }

                // 既存のブックマークを削除
                DeleteNestedBookmarks(thisDocument);

                // 既存のブックマークのうち、書誌情報番号に該当しないものを削除
                // 入れ子構造やコレクションの動的変化により、1回のループでは全ての不要なブックマークを削除できない場合がある
                // 2回繰り返すことで、削除漏れを防ぎ、確実に不要なブックマークを全て削除する

                DeleteUnmatchedNestedBookmarks(thisDocument, docid, bookInfoDef);
                DeleteUnmatchedNestedBookmarks(thisDocument, docid, bookInfoDef);

                // ブックマークリストと最大番号を更新
                UpdateBookmarkListAndMaxNum(thisDocument, ls, ref bibMaxNum);

                maxNo = bibMaxNum;

                int splitCount = 1;

                string lv1styleName = "";
                string lv2styleName = "";
                string lv3styleName = "";

                int lv1count = 0;
                int lv2count = 0;
                int lv3count = 0;

                bool breakFlg = false;

                if (!bookInfoDic.ContainsKey(docid + "00000"))
                {
                    bookInfoDic.Add(docid + "00000", "表紙");
                }

                log.WriteLine("書誌情報リスト作成開始");
                string upperClassID = "";
                string previousSetId = "";
                bool isMerge = false;
                Dictionary<string, string> mergeSetId = new Dictionary<string, string>();
                title4Collection = new Dictionary<string, string[]>();
                headerCollection = new Dictionary<string, string[]>();

                foreach (Word.Section tgtSect in thisDocument.Sections)
                {
                    foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                    {
                        string styleName = tgtPara.get_Style().NameLocal;

                        AddReferenceBookmarks(tgtPara, styleName);

                        isMerge = false;

                        try
                        {
                            string styleCharacterName = tgtPara.Range.CharacterStyle.NameLocal;
                            if (styleCharacterName.Equals("MJS_見出し結合用"))
                            {
                                isMerge = true;
                            }
                        }
                        catch (Exception) { }

                        // タイトル4コレクションに追加
                        AddTitle4CollectionIfHeading(tgtPara, styleName, upperClassID, title4Collection);

                        if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                        string innerText = tgtPara.Range.Text.Trim();

                        if (tgtPara.Range.Text.Trim() == "") continue;

                        if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                        {
                            breakFlg = true;
                            break;
                        }

                        if (Regex.IsMatch(styleName, @"章[　 ]*扉.*タイトル"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";

                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = setid;

                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            lv1count++;
                            lv2styleName = "";
                            lv2count = 0;
                            lv3styleName = "";
                            lv3count = 0;

                            lv1styleName = styleName;
                        }
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$"))
                        {
                            Application.DoEvents();
                            if (!Regex.IsMatch(innerText, @"目\s*次\s*$"))
                            {
                                // 行末尾を選択状態にする
                                tgtPara.Range.Select();
                                Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                                sel.EndKey(Word.WdUnits.wdLine);

                                string setid = "";
                                foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                                {
                                    if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                    {
                                        setid = bm.Name;
                                        upperClassID = bm.Name;

                                        // 行末尾にブックマークを追加する
                                        sel.Bookmarks.Add(setid);

                                        break;
                                    }
                                }

                                if (setid == "")
                                {
                                    bibMaxNum++;
                                    splitCount = bibMaxNum;
                                    ls.Add(splitCount.ToString("000"));
                                    setid = docid + bookInfoDef + splitCount.ToString("000");
                                    upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }
                                    previousSetId = setid;
                                }
                                else if (!bookInfoDic.ContainsKey(setid))
                                {
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }
                                    previousSetId = setid;
                                }

                                if ((lv1styleName == "") || (lv1styleName == styleName) || Regex.IsMatch(lv1styleName, @"(見出し|Heading)\s*[２2]"))
                                {
                                    lv1count++;
                                    lv2styleName = "";
                                    lv2count = 0;
                                    lv3styleName = "";
                                    lv3count = 0;

                                    lv1styleName = styleName;
                                }
                                else
                                {
                                    lv2count++;
                                    lv3styleName = "";
                                    lv3count = 0;

                                    lv2styleName = styleName;
                                }
                            }
                        }
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);

                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            if ((lv1styleName == "") || (lv1styleName == styleName))
                            {
                                lv1count++;
                                lv2styleName = "";
                                lv2count = 0;
                                lv3styleName = "";
                                lv3count = 0;

                                lv1styleName = styleName;
                            }
                            else if ((lv2styleName == "") || (lv2styleName == styleName))
                            {
                                lv2count++;
                                lv3styleName = "";
                                lv3count = 0;

                                lv2styleName = styleName;
                            }
                            else
                            {
                                lv3count++;
                                lv3styleName = styleName;
                            }
                        }
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "♯" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(♯.*?)$", "$1");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "＃" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(＃.*?)$", "$1");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000");
                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));

                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            if ((lv1styleName == "") || (lv1styleName == styleName))
                            {
                                lv1count++;
                                lv2styleName = "";
                                lv2count = 0;
                                lv3styleName = "";
                                lv3count = 0;

                                lv1styleName = styleName;
                            }
                            else if ((lv2styleName == "") || (lv2styleName == styleName))
                            {
                                lv2count++;
                                lv3styleName = "";
                                lv3count = 0;
                                lv2styleName = styleName;
                            }
                            else if ((lv3styleName == "") || (lv3styleName == styleName))
                            {
                                lv3count++;
                                lv3styleName = styleName;
                            }
                            else
                            {
                                continue;
                            }
                        }
                    }

                    if (breakFlg) break;
                }

                // SOURCELINK変更==========================================================================START

                if (checkBL || oldInfo.Count == 0)
                {
                    WriteBookInfoToFile(rootPath, headerDir, docid, bookInfoDic, mergeSetId, MakeHeaderLine);

                    thisDocument.Save();

                    log.WriteLine("書誌情報リスト作成終了");
                }
                else
                {
                    // 書誌情報（新）
                    AddHeadingInfoFromBookInfoDic(bookInfoDic, mergeSetId, newInfo);

                    // 新旧比較処理
                    int ret = CheckDocInfo(oldInfo, newInfo, out checkResult);

                    // 処理結果が0:正常の場合
                    if (ret == 0)
                    {
                        using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                        {
                            foreach (HeadingInfo info in newInfo)
                            {
                                MakeHeaderLine(docinfo, mergeSetId, info.num, info.title, info.id);
                            }
                        }

                        thisDocument.Save();

                        log.WriteLine("書誌情報リスト作成終了");
                    }
                    else if (ret == 1)
                    {
                        // 処理結果が1:異常の場合
                        // 書誌情報比較チェック画面を表示する
                        load.Visible = false;
                        CheckForm checkForm = new CheckForm(this);
                        DialogResult returnCode = checkForm.ShowDialog();

                        if (returnCode != DialogResult.OK)
                        {

                            if (swLog == null)
                            {
                                log.Close();
                            }

                            return false;
                        }
                        else
                        {
                            if (blHTMLPublish)
                                load.Visible = true;

                            // 新.IDをドキュメントに反映する
                            foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();

                            // 比較結果（checkResult）に基づいたブックマークを追加
                            AddBookmarksFromCheckResult(thisDocument, checkResult, ref breakFlg);

                            // 書誌情報をファイルに書き込む
                            WriteCheckInfoToFile(rootPath, headerDir, docid, checkResult, mergeSetId, MakeHeaderLine);

                            thisDocument.Save();

                            log.WriteLine("書誌情報リスト作成終了");
                        }
                    }
                }

                // SOURCELINK変更==========================================================================END

                if (swLog == null)
                {
                    log.Close();
                    File.Delete(rootPath + "\\log.txt");
                }
                blHTMLPublish = false;
                return true;
            }
            catch (Exception ex)
            {
                return HandleMakeBookInfoException(ex, log, swLog, load, (Button)button4, ref blHTMLPublish);
            }
            finally
            {
                Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
                Application.DoEvents();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
            }
        }
    }
}
