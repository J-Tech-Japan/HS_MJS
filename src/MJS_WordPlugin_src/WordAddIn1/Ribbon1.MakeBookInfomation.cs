﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Diagnostics;
using System.Security.Cryptography.X509Certificates;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private bool makeBookInfo(loader load, StreamWriter swLog = null)
        {
            // 画面更新を無効化して処理を高速化
            WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = false;
            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;

            // 命名規則に違反している場合
            if (!Regex.IsMatch(thisDocument.Name, @"^[A-Z]{3}(_[^_]*?){2}\.docx*$"))
            {
                // エラーメッセージを表示して処理を終了
                load.Visible = false;
                MessageBox.Show("開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」", "ファイル命名規則エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.DoEvents();
                WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
                return false;
            }

            // 現在の選択範囲の開始位置と終了位置を保存
            int selStart = WordAddIn1.Globals.ThisAddIn.Application.Selection.Start;
            int selEnd = WordAddIn1.Globals.ThisAddIn.Application.Selection.End;

            // ドキュメント全体を選択
            WordAddIn1.Globals.ThisAddIn.Application.Selection.EndKey(Word.WdUnits.wdStory);
            Application.DoEvents();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
            Application.DoEvents();


            // 選択範囲が図形の場合、カーソルを左に移動
            if (WordAddIn1.Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionInlineShape ||
                WordAddIn1.Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionShape)
                WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveLeft(Word.WdUnits.wdCharacter);

            // 書誌情報の初期化
            bookInfoDef = "";
            Word.Document Doc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;

            // 書誌情報番号とその最大値を初期化
            int bibNum = 0;  // 現在の書誌情報番号
            int bibMaxNum = 0;  // 書誌情報番号の最大値

            // チェックフラグの初期化
            bool checkBL = false;

            // 指定されたヘッダーファイルが存在するか確認
            if (File.Exists(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt"))
            {
                // ヘッダーファイルを開けるか確認（他のプロセスでロックされていないかチェック）
                try
                {
                    using (Stream stream = new FileStream(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt", FileMode.Open))
                    {
                    }
                }
                catch
                {
                    // ファイルが開かれている場合、エラーメッセージを表示して処理を終了
                    load.Visible = false;
                    MessageBox.Show(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt" + "が開かれています。\r\nファイルを閉じてから書誌情報出力を実行してください。",
                        "ファイルエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.DoEvents();
                    WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return false;
                }

                // SOURCELINK追加==========================================================================START
                oldInfo = new List<HeadingInfo>();  // 旧書誌情報
                newInfo = new List<HeadingInfo>();  // 新書誌情報
                checkResult = new List<CheckInfo>();  // 比較結果

                // SOURCELINK追加==========================================================================END

                // ヘッダーファイルを読み込み、書誌情報番号の最大値を取得
                using (StreamReader sr = new StreamReader(
                    Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt", System.Text.Encoding.Default))
                {
                    // 書誌情報番号の最大値取得
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

                        // SOURCELINK追加==========================================================================START

                        // ヘッダーファイルの内容を分割して書誌情報を作成
                        string[] info = strBuffer.Split('\t');

                        HeadingInfo headingInfo = new HeadingInfo();
                        headingInfo.num = info[0];  // 書誌番号
                        headingInfo.title = info[1];  // タイトル

                        if (info.Length == 4)
                        {
                            headingInfo.mergeto = info[3];  // 結合先情報
                        }

                        headingInfo.id = info[2];  // ID
                        oldInfo.Add(headingInfo);  // 旧書誌情報リストに追加

                        // SOURCELINK追加==========================================================================END

                        // 書誌情報番号の最大値を取得
                        bibNum = int.Parse(info[2].Substring(info[2].Length - 3, 3));
                        if (bibMaxNum < bibNum)
                        {
                            bibMaxNum = bibNum;
                        }
                    }
                }

                // ドキュメント内のブックマークを確認し、書誌情報のデフォルト値を取得
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
                // ヘッダーファイルが存在しない場合、ボタンを無効化してチェックフラグを設定
                button3.Enabled = false;
                button2.Enabled = false;
                button5.Enabled = false;
                checkBL = true;
            }

            // ドキュメントの保存先ディレクトリのパスを取得
            string rootPath = thisDocument.Path;

            // ドキュメントのファイル名を取得
            string docName = thisDocument.Name;

            // ヘッダーファイルが格納されるディレクトリ名を指定
            string headerDir = "headerFile";

            // ドキュメント名の先頭3文字を抽出してドキュメントIDとして使用
            string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");

            // ドキュメント名からタイトルを抽出
            string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");
            
            bookInfoDic.Clear();

            // ログ出力用のStreamWriterを設定（引数で渡されたものを使用、なければ後で新規作成）
            StreamWriter log = swLog;

            // ログファイルが指定されていない場合、新規作成
            if (swLog == null)
            {
                // "log.txt" ファイルを作成し、UTF-8 で書き込みを行う
                log = new StreamWriter(rootPath + "\\log.txt", false, Encoding.UTF8);
            }

            try
            {
                // 書誌情報のデフォルト値が空の場合、ユーザーに入力を求める
                if (bookInfoDef == "")
                {
                    // ドキュメント内のすべてのブックマークを削除
                    foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();

                    // 書誌情報入力フォームを表示
                    using (bookInfo bi = new bookInfo())
                    {
                        if (bi.ShowDialog() == DialogResult.OK)
                        {
                            // ユーザーが入力したデフォルト値を取得
                            bookInfoDef = bi.tbxDefaultValue.Text;
                        }
                        else
                        {
                            // キャンセルされた場合、ログを閉じてファイルを削除して処理を終了
                            log.Close();
                            if (File.Exists(rootPath + "\\log.txt")) File.Delete(rootPath + "\\log.txt");
                            button4.Enabled = true;
                            return false;
                        }
                    }
                }

                // 旧書誌情報を格納する辞書と一時的なセットを初期化
                Dictionary<string, string> oldBookInfoDic = new Dictionary<string, string>();
                HashSet<string> ls = new HashSet<string>();

                // ヘッダーファイルのディレクトリが存在しない場合、新規作成
                if (!Directory.Exists(rootPath + "\\" + headerDir))
                {
                    Directory.CreateDirectory(rootPath + "\\" + headerDir);
                }

                // ドキュメント内のすべてのブックマークを確認し、ネストされたブックマークを削除
                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
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
                        // エラーが発生した場合、例外をログに出力
                        Console.WriteLine(e);
                    }
                }

                // ブックマークの名前が指定された形式に一致しない場合削除
                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                {
                    foreach (Word.Bookmark wbInWb in wb.Range.Bookmarks)
                    {
                        if (!Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") &&
                            !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}♯" + docid + bookInfoDef + @"\d{3}$") &&
                            !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") &&
                            !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}＃" + docid + bookInfoDef + @"\d{3}$"))
                            wbInWb.Delete();
                    }
                }

                // 重複するブックマークを削除し、一意の名前をセットに追加
                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                {
                    if (!ls.Contains(wb.Name.Substring(wb.Name.Length - 3, 3)))
                        ls.Add(wb.Name.Substring(wb.Name.Length - 3, 3));
                    else
                        wb.Delete();
                }


                // ブックマーク名の最大値を取得し、書誌情報番号の最大値を更新
                if (ls.Count != 0)
                {
                    string maxResult = ls.Max(val => val);
                    if (int.Parse(maxResult) > bibMaxNum) bibMaxNum = int.Parse(maxResult);
                }

                // 書誌情報番号の最大値を設定
                maxNo = bibMaxNum;

                // 分割カウントとスタイル名、カウントを初期化
                int splitCount = 1;
                string lv1styleName = "";
                string lv2styleName = "";
                string lv3styleName = "";
                int lv1count = 0;
                int lv2count = 0;
                int lv3count = 0;

                // 処理を中断するフラグを初期化
                bool breakFlg = false;

                // 書誌情報辞書に「表紙」のエントリが存在しない場合
                if (!bookInfoDic.ContainsKey(docid + "00000"))
                {
                    // ドキュメントIDに「00000」を付加して「表紙」として登録
                    bookInfoDic.Add(docid + "00000", "表紙");
                }

                // ログに書誌情報リスト作成の開始を記録
                log.WriteLine("書誌情報リスト作成開始");

                // 上位クラスIDを保持する変数を初期化
                string upperClassID = "";

                // 前回のセットIDを保持する変数を初期化
                string previousSetId = "";

                // 結合処理が必要かどうかを示すフラグを初期化
                bool isMerge = false;

                // 結合先情報を保持する辞書を初期化
                Dictionary<string, string> mergeSetId = new Dictionary<string, string>();

                // タイトル情報を保持するコレクションを初期化
                title4Collection = new Dictionary<string, string[]>();

                // ヘッダー情報を保持するコレクションを初期化
                headerCollection = new Dictionary<string, string[]>();


                // ドキュメント内のセクションをループ
                foreach (Word.Section tgtSect in thisDocument.Sections)
                {
                    // セクションの範囲内の段落をループ
                    foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                    {
                        // 段落のスタイル名を取得
                        string styleName = tgtPara.get_Style().NameLocal;

                        // 段落のスタイル名が「MJS_参照先」の場合
                        if (styleName.Equals("MJS_参照先"))
                        {
                            // 段落内のフィールドをループ
                            foreach (Word.Field fld in tgtPara.Range.Fields)
                            {
                                // フィールドが参照フィールドの場合
                                if (fld.Type == Word.WdFieldType.wdFieldRef)
                                {
                                    // フィールドコードからブックマーク名を生成し、"_ref"を付加
                                    string bookmarkName = fld.Code.Text.Split(new char[] { ' ' })[2] + "_ref";

                                    // ブックマークを段落範囲に追加
                                    tgtPara.Range.Bookmarks.Add(bookmarkName);

                                    // フィールドコードをハイパーリンク形式に変更
                                    fld.Code.Text = "HYPERLINK " + fld.Code.Text.Split(new char[] { ' ' })[2];
                                }
                            }
                        }

                        // 結合処理フラグを初期化
                        isMerge = false;

                        try
                        {
                            // 段落の文字スタイル名を取得
                            string styleCharacterName = tgtPara.Range.CharacterStyle.NameLocal;

                            // スタイルが「MJS_見出し結合用」の場合
                            if (styleCharacterName.Equals("MJS_見出し結合用"))
                            {
                                isMerge = true;
                            }
                        }
                        catch (Exception) { }

                        // スタイル名が特定の「見出し」形式に一致する場合
                        if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[４4](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[５5](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            // 隠しブックマークを表示
                            tgtPara.Range.Bookmarks.ShowHidden = true;

                            // 段落内のすべてのブックマークをループ処理
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                // ブックマーク名が"タイトルコレクションに未登録の場合
                                if (!title4Collection.ContainsKey(bm.Name))
                                {
                                    // ブックマーク名が"_Ref"で始まる場合
                                    if (bm.Name.IndexOf("_Ref") == 0)
                                    {
                                        // タイトルコレクションにブックマーク名と関連情報を追加
                                        title4Collection.Add(bm.Name, new string[] { upperClassID, tgtPara.Range.Text.Replace("\r", "").Replace("\n", "").Replace("\"", "\"\"") });
                                    }
                                }
                            }

                            // 隠しブックマークを非表示に戻す
                            tgtPara.Range.Bookmarks.ShowHidden = false;
                        }

                        // スタイル名が「章 扉 タイトル」に一致しない、かつ「見出し」を含まない場合は次の段落へ
                        if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                        // 段落のテキストをトリムして取得
                        string innerText = tgtPara.Range.Text.Trim();

                        // 段落のテキストが空の場合は次の段落へ
                        if (tgtPara.Range.Text.Trim() == "") continue;

                        // 段落のテキストが「索引」に一致し、スタイルが「章 扉 タイトル」または「見出し1」に一致する場合
                        if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                        {
                            // 処理を中断するフラグを設定し、ループを終了
                            breakFlg = true;
                            break;
                        }

                        // スタイル名が「章 扉 タイトル」に一致する場合
                        if (Regex.IsMatch(styleName, @"章[　 ]*扉.*タイトル"))
                        {
                            // 他のイベントを処理してUIを更新
                            Application.DoEvents();

                            // 段落の行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            // ブックマークIDを初期化
                            string setid = "";

                            // 段落内のすべてのブックマークをループ処理
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                // ブックマーク名が「docid + bookInfoDef + 3桁の数字」の形式に一致する場合
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    // ブックマークIDを設定し、上位クラスIDとして保持
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            // ブックマークIDが空の場合、新しいIDを生成
                            if (setid == "")
                            {
                                // 書誌情報番号の最大値をインクリメント
                                bibMaxNum++;
                                splitCount = bibMaxNum;

                                // 一意の番号をリストに追加
                                ls.Add(splitCount.ToString("000"));

                                // 新しいブックマークIDを生成し、 上位クラスIDとして設定
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = setid;

                                // 行末尾に新しいブックマークを追加
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                // 書誌情報辞書に新しいエントリを追加
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                // 結合処理が必要な場合、結合先情報を追加
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            // 既存のブックマークIDが書誌情報辞書に存在しない場合
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                // 書誌情報辞書に新しいエントリを追加
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                // 結合処理が必要な場合、結合先情報を追加
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            // レベル1のカウントをインクリメント
                            lv1count++;

                            // レベル2とレベル3のスタイル名とカウントを初期化
                            lv2styleName = "";
                            lv2count = 0;
                            lv3styleName = "";
                            lv3count = 0;

                            // レベル1のスタイル名を現在のスタイル名に設定
                            lv1styleName = styleName;
                        }

                        // スタイル名が「見出し1」に一致する場合
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$"))
                        {
                            // 他のイベントを処理してUIを更新
                            Application.DoEvents();

                            // 段落のテキストが「目次」に一致しない場合
                            if (!Regex.IsMatch(innerText, @"目\s*次\s*$"))
                            {
                                // 段落の行末尾を選択状態にする
                                tgtPara.Range.Select();
                                Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                                sel.EndKey(Word.WdUnits.wdLine);

                                string setid = "";

                                // 段落内のすべてのブックマークをループ処理
                                foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                                {
                                    // ブックマーク名が「docid + bookInfoDef + 3桁の数字」の形式に一致する場合
                                    if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                    {
                                        // ブックマークIDを設定し、上位クラスIDとして保持
                                        setid = bm.Name;
                                        upperClassID = bm.Name;

                                        // 行末尾にブックマークを追加
                                        sel.Bookmarks.Add(setid);

                                        break;
                                    }
                                }

                                // ブックマークIDが空の場合、新しいIDを生成
                                if (setid == "")
                                {
                                    // 書誌情報番号の最大値をインクリメント
                                    bibMaxNum++;
                                    splitCount = bibMaxNum;

                                    // 一意の番号をリストに追加
                                    ls.Add(splitCount.ToString("000"));

                                    // 新しいブックマークIDを生成し、 上位クラスIDとして設定
                                    setid = docid + bookInfoDef + splitCount.ToString("000");
                                    upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                    // 行末尾に新しいブックマークを追加
                                    sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                    // 書誌情報辞書に新しいエントリを追加
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                    // 結合処理が必要な場合、結合先情報を追加
                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }

                                    // 前回のセットIDを更新
                                    previousSetId = setid;
                                }

                                // 既存のブックマークIDが書誌情報辞書に存在しない場合
                                else if (!bookInfoDic.ContainsKey(setid))
                                {
                                    // 書誌情報辞書に新しいエントリを追加
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                    // 結合処理が必要な場合、結合先情報を追加
                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }

                                    // 前回のセットIDを更新
                                    previousSetId = setid;
                                }

                                // スタイル名が空、または現在のスタイル名と一致する場合、または「見出し2」に一致する場合
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

                        // スタイル名が「見出し2」に一致する場合
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])"))
                        {
                            Application.DoEvents();

                            // 段落の行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";

                            // 段落内のすべてのブックマークをループ処理
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                // ブックマーク名が「docid + bookInfoDef + 3桁の数字」の形式に一致する場合
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    // ブックマークIDを設定し、上位クラスIDとして保持
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                // 書誌情報番号の最大値をインクリメント
                                bibMaxNum++;
                                splitCount = bibMaxNum;

                                // 一意の番号をリストに追加
                                ls.Add(splitCount.ToString("000"));

                                // 新しいブックマークIDを生成し、 上位クラスIDとして設定
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                // 行末尾にブックマークを追加
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                // 書誌情報辞書に新しいエントリを追加
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());

                                // 結合処理が必要な場合、結合先情報を追加
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            // 既存のブックマークIDが書誌情報辞書に存在しない場合
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                // 書誌情報辞書に新しいエントリを追加
                                // ブックマークIDをキーとして、段落のリスト番号とテキストを結合
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

                        // スタイル名が「見出し3」に一致する場合
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            Application.DoEvents();

                            // 段落の行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";

                            // 段落内のブックマークをループ処理
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                // ブックマーク名が（docid + bookInfoDef + 3桁 + ♯ + docid + bookInfoDef + 3桁）に一致する場合
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "♯" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(♯.*?)$", "$1");

                                    // 行末尾にブックマークを追加
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }

                                
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "＃" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    // 上位クラスIDとブックマーク名を結合して、新しいIDを生成
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(＃.*?)$", "$1");

                                    // 行末尾にブックマークを追加
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            // ブックマークIDが空の場合、新しいIDを生成
                            if (setid == "")
                            {
                                bibMaxNum++;
                                splitCount = bibMaxNum;

                                // 一意の番号をリストに追加
                                ls.Add(splitCount.ToString("000"));

                                // 新しいブックマークIDを生成し、 上位クラスIDとして設定
                                setid = upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000");
                                
                                // 行末尾にブックマークを追加
                                sel.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));

                                // 書誌情報辞書に新しいエントリを追加
                                // キー: 新しいブックマークID、値: 段落のリスト番号とテキストを結合した文字列
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            // 既存のブックマークIDが書誌情報辞書に存在しない場合
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                // 書誌情報辞書に新しいエントリを追加
                                // キー: 既存のブックマークID、値: 段落のリスト番号とテキストを結合した文字列
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }

                                // 前回のセットIDを更新
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

                // チェックフラグが立っている、または旧書誌情報が空の場合
                if (checkBL || oldInfo.Count == 0)
                {
                    // 書誌情報を保存するためのファイルを作成
                    using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                    {
                        // 書誌情報辞書のすべてのキーをループ処理
                        foreach (string key in bookInfoDic.Keys)
                        {
                            // 書誌情報を分割して取得
                            string[] secText = new string[2];

                            // 書誌情報に「♪」が含まれている場合、分割して項番とタイトルを取得
                            if (bookInfoDic[key].Contains("♪"))
                            {
                                secText[0] = Regex.Replace(bookInfoDic[key], "^(.*?)♪.*?$", "$1");
                                secText[1] = Regex.Replace(bookInfoDic[key], "^.*?♪(.*?)$", "$1");
                            }
                            else
                                // 「♪」が含まれていない場合、タイトルのみを設定
                                secText[1] = bookInfoDic[key];

                            // 書誌情報を格納するクラスのインスタンスを作成
                            HeadingInfo headingInfo = new HeadingInfo();

                            // 項番が空の場合は空文字を設定、それ以外の場合はsecText[0]を設定
                            if (string.IsNullOrEmpty(secText[0]))
                            {
                                headingInfo.num = "";
                            }
                            else
                            {
                                headingInfo.num = secText[0];
                            }

                            // タイトルが空の場合は空文字を設定、それ以外の場合はsecText[1]を設定
                            if (string.IsNullOrEmpty(secText[1]))
                            {
                                headingInfo.title = "";
                            }
                            else
                            {
                                headingInfo.title = secText[1];
                            }

                            // IDを設定（特殊文字「♯」を「#」に置換）
                            headingInfo.id = key.Replace("♯", "#");

                            // 結合先情報が存在する場合
                            if (mergeSetId.ContainsKey(headingInfo.id))
                            {
                                // 結合先IDを取得し、headingInfo.mergetoに設定
                                headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];

                                // ヘッダー行を作成してファイルに書き込む
                                makeHeaderLine(docinfo, mergeSetId, headingInfo.num, headingInfo.title, headingInfo.id);
                            }
                            else
                            {
                                // 結合先情報がない場合、項番、タイトル、IDをタブ区切りでファイルに書き込む
                                docinfo.WriteLine(secText[0] + "\t" + secText[1] + "\t" + key.Replace("♯", "#") + "\t");
                            }
                        }
                    }

                    thisDocument.Save();

                    log.WriteLine("書誌情報リスト作成終了");
                }
                else
                {
                    // 書誌情報辞書のキーをループ処理
                    foreach (string key in bookInfoDic.Keys)
                    {
                        // 書誌情報を分割して取得
                        string[] secText = new string[2];

                        // 書誌情報に「♪」が含まれている場合、項番とタイトルを分割
                        if (bookInfoDic[key].Contains("♪"))
                        {
                            secText[0] = Regex.Replace(bookInfoDic[key], "^(.*?)♪.*?$", "$1");
                            secText[1] = Regex.Replace(bookInfoDic[key], "^.*?♪(.*?)$", "$1");
                        }
                        // 書誌情報に「♪」が含まれていない場合、タイトルのみを設定
                        else
                            secText[1] = bookInfoDic[key];

                        // 書誌情報を格納するクラスのインスタンスを作成
                        HeadingInfo headingInfo = new HeadingInfo();

                        // 項番が空の場合は空文字を設定、それ以外の場合はsecText[0]を設定
                        if (string.IsNullOrEmpty(secText[0]))
                        {
                            headingInfo.num = "";
                        }
                        else
                        {
                            headingInfo.num = secText[0];
                        }

                        // タイトルが空の場合は空文字を設定、それ以外の場合はsecText[1]を設定
                        if (string.IsNullOrEmpty(secText[1]))
                        {
                            headingInfo.title = "";
                        }
                        else
                        {
                            headingInfo.title = secText[1];
                        }

                        // 特殊文字「＃」が含まれている場合
                        // 「＃」を「#」に置換してIDを設定
                        if (key.Contains("＃"))
                        {
                            headingInfo.id = key.Replace("＃", "#");
                        }

                        // 特殊文字「＃」が含まれていない場合
                        else
                        {
                            // 「♯」を「#」に置換してIDを設定
                            headingInfo.id = key.Replace("♯", "#");
                        }

                        // 結合先情報が存在する場合
                        if (mergeSetId.ContainsKey(headingInfo.id))
                        {
                            // 結合先IDを取得し、headingInfo.mergetoに設定
                            headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                        }

                        // 新しい書誌情報をリストに追加
                        newInfo.Add(headingInfo);
                    }

                    // 新旧比較処理
                    int ret = checkDocInfo(oldInfo, newInfo, out checkResult);

                    // 処理結果が0:正常の場合
                    if (ret == 0)
                    {
                        // 処理結果が正常の場合、書誌情報を保存するためのファイルを作成
                        using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                        {
                            foreach (HeadingInfo info in newInfo)
                            {
                                makeHeaderLine(docinfo, mergeSetId, info.num, info.title, info.id);
                            }
                        }

                        thisDocument.Save();
                        log.WriteLine("書誌情報リスト作成終了");
                    }

                    // 処理結果が1（異常）の場合の処理
                    else if (ret == 1)
                    {
                        load.Visible = false;
                        CheckForm checkForm = new CheckForm(this);

                        // ダイアログを表示し、ユーザーの操作結果を取得
                        DialogResult returnCode = checkForm.ShowDialog();

                        // ユーザーが「OK」以外を選択した場合
                        if (returnCode != DialogResult.OK)
                        {
                            // ログファイルが指定されていない場合、ログを閉じる
                            if (swLog == null)
                            {
                                log.Close();
                            }

                            return false;
                        }
                        else
                        {
                            // HTML公開フラグが有効な場合、ロード画面を再表示
                            if (blHTMLPublish)
                                load.Visible = true;

                            // ドキュメント内のすべてのブックマークを削除
                            foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();

                            // ドキュメント内のセクションをループ処理
                            foreach (Word.Section tgtSect in thisDocument.Sections)
                            {
                                // セクション内の段落をループ処理
                                foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                                {
                                    // 段落のスタイル名を取得
                                    string styleName = tgtPara.get_Style().NameLocal;

                                    // スタイル名が「章 扉 タイトル」に一致しない、かつ「見出し」を含まない場合は次の段落へ
                                    if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                                    // 段落のテキストを取得
                                    string innerText = tgtPara.Range.Text.Trim();

                                    // 段落のテキストが空の場合は次の段落へ
                                    if (tgtPara.Range.Text.Trim() == "") continue;

                                    // 段落のテキストが「索引」に一致し、特定のスタイル名の場合、処理を中断
                                    if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                                    {
                                        // 処理中断フラグを設定
                                        breakFlg = true;
                                        break;
                                    }

                                    // スタイル名が特定の条件に一致する場合
                                    if (Regex.IsMatch(styleName, @"章[　 ]*扉.*タイトル")
                                        || (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$") && !Regex.IsMatch(innerText, @"目\s*次\s*$"))
                                        || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                                        || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                                    {
                                        Application.DoEvents();

                                        // 行末尾を選択状態にする
                                        tgtPara.Range.Select();
                                        Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                                        sel.EndKey(Word.WdUnits.wdLine);

                                        // 項番を取得
                                        string num = Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "");

                                        // タイトルを取得
                                        string title = tgtPara.Range.Text.Trim();

                                        // 比較結果リストから一致する情報を検索
                                        CheckInfo info = checkResult.Where(p => ((string.IsNullOrEmpty(p.new_num) && string.IsNullOrEmpty(num)) || p.new_num.Equals(num))
                                            && p.new_title.Equals(title)).FirstOrDefault();

                                        // 一致する情報が存在する場合、ブックマークを追加
                                        if (info != null)
                                        {
                                            // 行末尾にブックマークを追加する
                                            sel.Bookmarks.Add(info.new_id_show.Split(new char[] { '(' })[0].Trim().Replace("#", "♯"));
                                        }
                                    }
                                }

                                // 処理中断フラグが設定されている場合、セクションのループを終了
                                if (breakFlg) break;
                            }

                            // 書誌情報を保存するためのファイルを作成
                            using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                            {
                                // 比較結果リストをループ処理
                                foreach (CheckInfo info in checkResult)
                                {
                                    // 新しいIDが空の場合はスキップ
                                    if (string.IsNullOrEmpty(info.new_id))
                                    {
                                        continue;
                                    }
                                    // ヘッダー行を作成してファイルに書き込む
                                    makeHeaderLine(docinfo, mergeSetId, info.new_num, info.new_title, info.new_id_show.Split(new char[] { '(' })[0].Trim());
                                }
                            }

                            thisDocument.Save();
                            log.WriteLine("書誌情報リスト作成終了");
                        }
                    }
                }

                // SOURCELINK変更==========================================================================END

                // ログファイルが指定されていない場合、ログを閉じる
                if (swLog == null)
                {
                    log.Close();
                    File.Delete(rootPath + "\\log.txt");
                }

                // HTML公開フラグを無効化
                blHTMLPublish = false;

                // 処理が正常に終了したことを示す
                return true;

            }

            catch (Exception ex)
            {
                // スタックトレースを取得（例外の発生箇所を特定するための情報）
                StackTrace stackTrace = new StackTrace(ex, true);

                // ログに例外の詳細情報を記録
                log.WriteLine(ex.Message);  // 例外メッセージ
                log.WriteLine(ex.HelpLink);  // ヘルプリンク
                log.WriteLine(ex.Source);  // 例外の発生元
                log.WriteLine(ex.StackTrace);  // スタックトレース
                log.WriteLine(ex.TargetSite);  // 例外が発生したメソッド

                // ログファイルが指定されていない場合、ログを閉じる
                if (swLog == null)
                {
                    log.Close();
                }

                load.Visible = false;
                MessageBox.Show("エラーが発生しました");

                // ボタンを有効化して操作可能にする
                button4.Enabled = true;

                // HTML公開フラグを無効化
                blHTMLPublish = false;

                return false;
            }
            finally
            {
                // ドキュメントのカーソル位置を先頭に戻す
                WordAddIn1.Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
                
                Application.DoEvents();

                //  画面更新を再有効化
                WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
            }

            //WordAddIn1.Globals.ThisAddIn.Application.Selection.Start = selStart;
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.End = selEnd;
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveRight(Unit: Word.WdUnits.wdCharacter, Count: 1);
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveLeft(Unit: Word.WdUnits.wdCharacter, Count: 1);
        }
    }
}
