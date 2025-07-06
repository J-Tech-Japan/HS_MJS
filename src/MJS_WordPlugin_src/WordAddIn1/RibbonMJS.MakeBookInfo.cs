using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
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

            //int selStart = Globals.ThisAddIn.Application.Selection.Start;
            //int selEnd = Globals.ThisAddIn.Application.Selection.End;
            //Globals.ThisAddIn.Application.Selection.EndKey(Word.WdUnits.wdStory);
            //Application.DoEvents();
            //Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
            //Application.DoEvents();

            //if (Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionInlineShape ||
            //    Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionShape)
            //    Globals.ThisAddIn.Application.Selection.MoveLeft(Word.WdUnits.wdCharacter);

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
                //if (bookInfoDef == "")
                //{

                //    foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();
                //    using (BookInfo bi = new BookInfo())
                //    {
                //        if (bi.ShowDialog() == DialogResult.OK)
                //        {
                //            bookInfoDef = bi.tbxDefaultValue.Text;
                //        }
                //        else
                //        {
                //            log.Close();
                //            if (File.Exists(rootPath + "\\log.txt")) File.Delete(rootPath + "\\log.txt");
                //            button4.Enabled = true;
                //            return false;
                //        }
                //    }
                //}

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

                //foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                //{
                //    try
                //    {
                //        for (int w = 1; w < wb.Range.Bookmarks.Count; w++)
                //        {
                //            wb.Range.Bookmarks[w].Delete();
                //        }
                //    }
                //    catch (Exception e)
                //    {
                //        Console.WriteLine(e);
                //    }
                //}

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

                        //if (styleName.Equals("MJS_参照先"))
                        //{
                        //    foreach (Word.Field fld in tgtPara.Range.Fields)
                        //    {
                        //        if (fld.Type == Word.WdFieldType.wdFieldRef)
                        //        {
                        //            string bookmarkName = fld.Code.Text.Split(new char[] { ' ' })[2] + "_ref";
                        //            tgtPara.Range.Bookmarks.Add(bookmarkName);
                        //            fld.Code.Text = "HYPERLINK " + fld.Code.Text.Split(new char[] { ' ' })[2];
                        //        }
                        //    }
                        //}

                        //AddReferenceBookmarks(tgtPara, styleName);

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


                        if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[４4](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[５5](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            tgtPara.Range.Bookmarks.ShowHidden = true;

                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (!title4Collection.ContainsKey(bm.Name))
                                {
                                    if (bm.Name.IndexOf("_Ref") == 0)
                                    {
                                        title4Collection.Add(bm.Name, new string[] { upperClassID, tgtPara.Range.Text.Replace("\r", "").Replace("\n", "").Replace("\"", "\"\"") });
                                    }
                                }
                            }
                            tgtPara.Range.Bookmarks.ShowHidden = false;
                        }
                        

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
                                    //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                    //while (ls.Contains(splitCount.ToString("000")))
                                    //{
                                    //    splitCount++;
                                    //}
                                    bibMaxNum++;
                                    splitCount = bibMaxNum;
                                    ls.Add(splitCount.ToString("000"));
                                    setid = docid + bookInfoDef + splitCount.ToString("000");
                                    upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                    //tgtPara.Range.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                    //splitCount++;
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
                                //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                //while (ls.Contains(splitCount.ToString("000")))
                                //{
                                //    splitCount++;
                                //}
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                //tgtPara.Range.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                //splitCount++;
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
                                //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                //while (ls.Contains(splitCount.ToString("000")))
                                //{
                                //    splitCount++;
                                //}
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000");
                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));

                                //tgtPara.Range.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                //splitCount++;
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
                    using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                    {

                        foreach (string key in bookInfoDic.Keys)
                        {
                            string[] secText = new string[2];
                            if (bookInfoDic[key].Contains("♪"))
                            {
                                secText[0] = Regex.Replace(bookInfoDic[key], "^(.*?)♪.*?$", "$1");
                                secText[1] = Regex.Replace(bookInfoDic[key], "^.*?♪(.*?)$", "$1");
                            }
                            else
                                secText[1] = bookInfoDic[key];
                            HeadingInfo headingInfo = new HeadingInfo();
                            if (string.IsNullOrEmpty(secText[0]))
                            {
                                headingInfo.num = "";
                            }
                            else
                            {
                                headingInfo.num = secText[0];
                            }
                            if (string.IsNullOrEmpty(secText[1]))
                            {
                                headingInfo.title = "";
                            }
                            else
                            {
                                headingInfo.title = secText[1];
                            }
                            headingInfo.id = key.Replace("♯", "#");

                            if (mergeSetId.ContainsKey(headingInfo.id))
                            {
                                headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                                MakeHeaderLine(docinfo, mergeSetId, headingInfo.num, headingInfo.title, headingInfo.id);
                            }
                            else
                            {
                                docinfo.WriteLine(secText[0] + "\t" + secText[1] + "\t" + key.Replace("♯", "#") + "\t");
                            }
                        }
                    }

                    thisDocument.Save();

                    log.WriteLine("書誌情報リスト作成終了");
                }
                else
                {
                    // 書誌情報（新）
                    foreach (string key in bookInfoDic.Keys)
                    {

                        string[] secText = new string[2];
                        if (bookInfoDic[key].Contains("♪"))
                        {
                            secText[0] = Regex.Replace(bookInfoDic[key], "^(.*?)♪.*?$", "$1");
                            secText[1] = Regex.Replace(bookInfoDic[key], "^.*?♪(.*?)$", "$1");
                        }
                        else
                            secText[1] = bookInfoDic[key];

                        HeadingInfo headingInfo = new HeadingInfo();
                        if (string.IsNullOrEmpty(secText[0]))
                        {
                            headingInfo.num = "";
                        }
                        else
                        {
                            headingInfo.num = secText[0];
                        }
                        if (string.IsNullOrEmpty(secText[1]))
                        {
                            headingInfo.title = "";
                        }
                        else
                        {
                            headingInfo.title = secText[1];
                        }
                        if (key.Contains("＃"))
                        {
                            headingInfo.id = key.Replace("＃", "#");
                        }
                        else
                        {
                            headingInfo.id = key.Replace("♯", "#");

                        }

                        if (mergeSetId.ContainsKey(headingInfo.id))
                        {
                            headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                        }

                        newInfo.Add(headingInfo);
                    }

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
                                //docinfo.WriteLine(info.num + "\t" + info.title + "\t" + info.id + "\t" + (mergeSetId.ContainsKey(info.id) ? mergeSetId[info.id]:""));
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

                            foreach (Word.Section tgtSect in thisDocument.Sections)
                            {
                                foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                                {
                                    string styleName = tgtPara.get_Style().NameLocal;

                                    if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                                    string innerText = tgtPara.Range.Text.Trim();

                                    if (tgtPara.Range.Text.Trim() == "") continue;

                                    if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                                    {
                                        breakFlg = true;
                                        break;
                                    }

                                    if (Regex.IsMatch(styleName, @"章[　 ]*扉.*タイトル")
                                        || (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$") && !Regex.IsMatch(innerText, @"目\s*次\s*$"))
                                        || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                                        || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                                    {
                                        Application.DoEvents();

                                        // 行末尾を選択状態にする
                                        tgtPara.Range.Select();
                                        Word.Selection sel = Globals.ThisAddIn.Application.Selection;
                                        sel.EndKey(Word.WdUnits.wdLine);

                                        string num = Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "");
                                        string title = tgtPara.Range.Text.Trim();

                                        CheckInfo info = checkResult.Where(p => ((string.IsNullOrEmpty(p.new_num) && string.IsNullOrEmpty(num)) || p.new_num.Equals(num))
                                            && p.new_title.Equals(title)).FirstOrDefault();

                                        if (info != null)
                                        {
                                            // 行末尾にブックマークを追加する
                                            sel.Bookmarks.Add(info.new_id_show.Split(new char[] { '(' })[0].Trim().Replace("#", "♯"));
                                        }
                                    }
                                }

                                if (breakFlg) break;
                            }

                            using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                            {
                                foreach (CheckInfo info in checkResult)
                                {
                                    if (string.IsNullOrEmpty(info.new_id))
                                    {
                                        continue;
                                    }
                                    MakeHeaderLine(docinfo, mergeSetId, info.new_num, info.new_title, info.new_id_show.Split(new char[] { '(' })[0].Trim());
                                    //docinfo.WriteLine(info.new_num + "\t" + info.new_title + "\t" + info.new_id_show + "\t" + (mergeSetId.ContainsKey(info.new_id_show) ? mergeSetId[info.new_id_show] : ""));
                                }
                            }

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
                StackTrace stackTrace = new StackTrace(ex, true);

                log.WriteLine(ex.Message);
                log.WriteLine(ex.HelpLink);
                log.WriteLine(ex.Source);
                log.WriteLine(ex.StackTrace);
                log.WriteLine(ex.TargetSite);

                if (swLog == null)
                {
                    log.Close();
                }
                load.Visible = false;
                MessageBox.Show("エラーが発生しました");

                button4.Enabled = true;
                blHTMLPublish = false;
                return false;
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
