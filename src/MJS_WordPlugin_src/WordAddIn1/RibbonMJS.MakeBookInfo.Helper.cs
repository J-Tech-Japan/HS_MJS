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
        private bool CheckFileName(Word.Document doc, loader load)
        {
            if (!Regex.IsMatch(doc.Name, FileNamePattern))
            {
                load.Visible = false;
                MessageBox.Show(
                    "開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」",
                    "ファイル命名規則エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error
                );
                Application.DoEvents();
                Globals.ThisAddIn.Application.ScreenUpdating = true;
                return false;
            }
            return true;
        }

        private bool CheckAndLoadHeaderFile(
            Word.Document Doc,
            loader load,
            ref List<HeadingInfo> oldInfo,
            ref List<HeadingInfo> newInfo,
            ref List<CheckInfo> checkResult,
            ref int bibNum,
            ref int bibMaxNum,
            ref string bookInfoDef,
            ref bool checkBL)
        {
            string headerFilePath = Path.Combine(
                Path.GetDirectoryName(Doc.FullName),
                "headerFile",
                Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + ".txt"
            );
            if (File.Exists(headerFilePath))
            {
                try
                {
                    using (Stream stream = new FileStream(headerFilePath, FileMode.Open))
                    {
                    }
                }
                catch
                {
                    load.Visible = false;
                    MessageBox.Show(headerFilePath + "が開かれています。\r\nファイルを閉じてから書誌情報出力を実行してください。",
                        "ファイルエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.DoEvents();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return false;
                }

                // SOURCELINK追加==========================================================================START
                // 書誌情報（旧）
                oldInfo = new List<HeadingInfo>();
                // 書誌情報（新）
                newInfo = new List<HeadingInfo>();
                // 比較結果
                checkResult = new List<CheckInfo>();
                // SOURCELINK追加==========================================================================END

                using (StreamReader sr = new StreamReader(headerFilePath, System.Text.Encoding.Default))
                {
                    // 書誌情報番号の最大値取得
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

                        // SOURCELINK追加==========================================================================START
                        string[] info = strBuffer.Split('\t');

                        HeadingInfo headingInfo = new HeadingInfo();
                        headingInfo.num = info[0];
                        headingInfo.title = info[1];
                        if (info.Length == 4)
                        {
                            headingInfo.mergeto = info[3];
                        }
                        headingInfo.id = info[2];

                        oldInfo.Add(headingInfo);

                        // SOURCELINK追加==========================================================================END

                        bibNum = int.Parse(info[2].Substring(info[2].Length - 3, 3));
                        if (bibMaxNum < bibNum)
                        {
                            bibMaxNum = bibNum;
                        }
                    }
                }

                foreach (Word.Bookmark bm in Doc.Bookmarks)
                {
                    if (Regex.IsMatch(bm.Name, "^" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1")))
                    {
                        bookInfoDef = Regex.Replace(bm.Name, "^.{3}(.{2}).*$", "$1");
                        break;
                    }
                }

                button4.Enabled = true;
                button2.Enabled = true;
                button5.Enabled = true;
            }
            else
            {
                button3.Enabled = false;
                button2.Enabled = false;
                button5.Enabled = false;
                checkBL = true;
            }
            return true;
        }

        private bool EnsureBookInfoDef(Word.Document thisDocument, ref string bookInfoDef, StreamWriter log, string rootPath)
        {
            if (bookInfoDef == "")
            {
                foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();
                using (BookInfo bi = new BookInfo())
                {
                    if (bi.ShowDialog() == DialogResult.OK)
                    {
                        bookInfoDef = bi.tbxDefaultValue.Text;
                    }
                    else
                    {
                        log.Close();
                        string logPath = Path.Combine(rootPath, "log.txt");
                        if (File.Exists(logPath)) File.Delete(logPath);
                        button4.Enabled = true;
                        return false;
                    }
                }
            }
            return true;
        }

        private void DeleteNestedBookmarks(Word.Document document)
        {
            foreach (Word.Bookmark wb in document.Bookmarks)
            {
                try
                {
                    for (int w = 1; w < wb.Range.Bookmarks.Count; w++)
                    {
                        wb.Range.Bookmarks[w].Delete();
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e);
                }
            }
        }

        private void DeleteUnmatchedNestedBookmarks(Word.Document document, string docid, string bookInfoDef)
        {
            foreach (Word.Bookmark wb in document.Bookmarks)
            {
                foreach (Word.Bookmark wbInWb in wb.Range.Bookmarks)
                {
                    if (!Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}♯" + docid + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}＃" + docid + bookInfoDef + @"\d{3}$"))
                    {
                        wbInWb.Delete();
                    }
                }
            }
        }

        private void UpdateBookmarkListAndMaxNum(Word.Document document, HashSet<string> ls, ref int bibMaxNum)
        {
            foreach (Word.Bookmark wb in document.Bookmarks)
            {
                if (!ls.Contains(wb.Name.Substring(wb.Name.Length - 3, 3)))
                    ls.Add(wb.Name.Substring(wb.Name.Length - 3, 3));
                else
                    wb.Delete();
            }
            if (ls.Count != 0)
            {
                string maxResult = ls.Max(val => val);
                if (int.Parse(maxResult) > bibMaxNum) bibMaxNum = int.Parse(maxResult);
            }
        }

        // REFフィールドをHYPERLINKに変換し、ブックマークを追加するメソッド
        // 元原稿に変更を加えないように、GenerateHTMLButton.csで呼び出す
        private void AddReferenceBookmarks(Word.Paragraph tgtPara, string styleName)
        {
            if (styleName.Equals("MJS_参照先"))
            {
                foreach (Word.Field fld in tgtPara.Range.Fields)
                {
                    if (fld.Type == Word.WdFieldType.wdFieldRef)
                    {
                        string[] codeParts = fld.Code.Text.Split(new char[] { ' ' });
                        if (codeParts.Length > 2)
                        {
                            string bookmarkName = codeParts[2] + "_ref";
                            tgtPara.Range.Bookmarks.Add(bookmarkName);
                            fld.Code.Text = "HYPERLINK " + codeParts[2];
                        }
                    }
                }
            }
        }



        // Word文書内の段落（Word.Paragraph）が特定の「見出し」スタイルである場合に、
        // その段落に含まれるブックマークのうち、名前が"_Ref"で始まるものを title4Collection コレクションに追加
        private void AddTitle4CollectionIfHeading(Word.Paragraph tgtPara, string styleName, string upperClassID, Dictionary<string, string[]> title4Collection)
        {
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
        }


        private void WriteBookInfoToFile(
            string rootPath,
            string headerDir,
            string docid,
            Dictionary<string, string> bookInfoDic,
            Dictionary<string, string> mergeSetId,
            Action<StreamWriter, Dictionary<string, string>, string, string, string> MakeHeaderLine)
        {
            string filePath = Path.Combine(rootPath, headerDir, docid + ".txt");
            using (StreamWriter docinfo = new StreamWriter(filePath, false, Encoding.UTF8))
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
                    {
                        secText[1] = bookInfoDic[key];
                    }
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
        }

        private void AddHeadingInfoFromBookInfoDic(
            Dictionary<string, string> bookInfoDic,
            Dictionary<string, string> mergeSetId,
            List<HeadingInfo> newInfo)
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
                {
                    secText[1] = bookInfoDic[key];
                }

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
        }

        // Word ドキュメント内の各段落を走査し、
        // 特定のスタイルや条件に一致する段落の行末に、
        // 比較結果（checkResult）に基づいたブックマークを追加
        private void AddBookmarksFromCheckResult(Word.Document thisDocument, List<CheckInfo> checkResult, ref bool breakFlg)
        {
            foreach (Word.Section tgtSect in thisDocument.Sections)
            {
                foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                {
                    string styleName = tgtPara.get_Style().NameLocal;

                    if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                    string innerText = tgtPara.Range.Text.Trim();

                    if (tgtPara.Range.Text.Trim() == "") continue;

                    if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") &&
                        (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
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

                        CheckInfo info = checkResult
                            .Where(p => ((string.IsNullOrEmpty(p.new_num) && string.IsNullOrEmpty(num)) || p.new_num.Equals(num))
                                && p.new_title.Equals(title))
                            .FirstOrDefault();

                        if (info != null)
                        {
                            // 行末尾にブックマークを追加する
                            sel.Bookmarks.Add(info.new_id_show.Split(new char[] { '(' })[0].Trim().Replace("#", "♯"));
                        }
                    }
                }

                if (breakFlg) break;
            }
        }

        // 書誌情報をファイルに書き込む
        private void WriteCheckInfoToFile(
            string rootPath,
            string headerDir,
            string docid,
            IEnumerable<CheckInfo> checkResult,
            Dictionary<string, string> mergeSetId,
            Action<StreamWriter, Dictionary<string, string>, string, string, string> MakeHeaderLine)
        {
            string filePath = Path.Combine(rootPath, headerDir, docid + ".txt");
            using (StreamWriter docinfo = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                foreach (CheckInfo info in checkResult)
                {
                    if (string.IsNullOrEmpty(info.new_id))
                    {
                        continue;
                    }
                    MakeHeaderLine(docinfo, mergeSetId, info.new_num, info.new_title, info.new_id_show.Split(new char[] { '(' })[0].Trim());
                }
            }
        }

        private bool HandleMakeBookInfoException(
            Exception ex,
            StreamWriter log,
            StreamWriter swLog,
            loader load,
            Button button4,
            ref bool blHTMLPublish)
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


    }
}
