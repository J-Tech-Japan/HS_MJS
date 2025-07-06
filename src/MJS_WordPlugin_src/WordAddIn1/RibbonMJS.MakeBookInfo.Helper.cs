using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using DocumentFormat.OpenXml.Wordprocessing;
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
    }
}
