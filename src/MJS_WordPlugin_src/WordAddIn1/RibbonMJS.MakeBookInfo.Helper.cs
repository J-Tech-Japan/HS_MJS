using System;
using System.Collections.Generic;
using System.IO;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Linq;
using System.Diagnostics;
using System.Text;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        public void AddReferenceFieldBookmarks(Word.Paragraph paragraph)
        {
            foreach (Word.Field fld in paragraph.Range.Fields)
            {
                // フィールドが参照フィールドの場合
                if (fld.Type == Word.WdFieldType.wdFieldRef)
                {
                    // フィールドコードからブックマーク名を生成し、"_ref"を付加
                    string bookmarkName = fld.Code.Text.Split(new char[] { ' ' })[2] + "_ref";

                    // ブックマークを段落範囲に追加
                    paragraph.Range.Bookmarks.Add(bookmarkName);

                    // フィールドコードをハイパーリンク形式に変更
                    fld.Code.Text = "HYPERLINK " + fld.Code.Text.Split(new char[] { ' ' })[2];
                }
            }
        }

        // 見出しのブックマークをコレクションに追加
        public void AddBookmarksToTitleCollection(Word.Paragraph tgtPara, Dictionary<string, string[]> title4Collection, string upperClassID)
        {
            // 隠しブックマークを表示
            tgtPara.Range.Bookmarks.ShowHidden = true;

            // 段落内のすべてのブックマークをループ処理
            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
            {
                // ブックマーク名が"_Ref"で始まり、タイトルコレクションに未登録の場合
                if (bm.Name.StartsWith("_Ref") && !title4Collection.ContainsKey(bm.Name))
                {
                    // タイトルコレクションにブックマーク名と関連情報を追加
                    title4Collection.Add(bm.Name, new string[] { upperClassID, tgtPara.Range.Text.Replace("\r", "").Replace("\n", "").Replace("\"", "\"\"") });
                }
            }

            // 隠しブックマークを非表示に戻す
            tgtPara.Range.Bookmarks.ShowHidden = false;
        }

        // ドキュメント内のブックマークを削除
        public void DeleteAllBookmarks(Word.Document document)
        {
            if (document.Bookmarks.Count > 0)
            {
                foreach (Word.Bookmark bookmark in document.Bookmarks.Cast<Word.Bookmark>().ToList())
                {
                    bookmark.Delete();
                }
            }
        }

        // ネストされたブックマークを削除
        public void DeleteNestedBookmarks(Word.Document document)
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
                    // エラーが発生した場合、例外をログに出力
                    Console.WriteLine(e);
                }
            }
        }

        // 名前が指定された形式に一致しないブックマークを削除
        public void DeleteInvalidBookmarks(Word.Document document, string docID, string bookInfoDef)
        {
            foreach (Word.Bookmark wb in document.Bookmarks)
            {
                foreach (Word.Bookmark wbInWb in wb.Range.Bookmarks)
                {
                    if (!Regex.IsMatch(wbInWb.Name, @"^" + docID + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docID + bookInfoDef + @"\d{3}♯" + docID + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docID + bookInfoDef + @"\d{3}$") &&
                        !Regex.IsMatch(wbInWb.Name, @"^" + docID + bookInfoDef + @"\d{3}＃" + docID + bookInfoDef + @"\d{3}$"))
                    {
                        wbInWb.Delete();
                    }
                }
            }
        }

        // 重複するブックマークを削除し、一意の名前をセットに追加
        public void DeleteDuplicateBookmarks(Word.Document document, HashSet<string> uniqueNames)
        {
            foreach (Word.Bookmark wb in document.Bookmarks)
            {
                string bookmarkSuffix = wb.Name.Substring(wb.Name.Length - 3, 3);
                if (!uniqueNames.Contains(bookmarkSuffix))
                {
                    uniqueNames.Add(bookmarkSuffix);
                }
                else
                {
                    wb.Delete();
                }
            }
        }

        // 指定された段落内のブックマークを検索し、条件に一致するブックマークを設定
        private void SetBookmarkIfMatch(
            Word.Bookmarks bookmarks,
            string docID,
            string bookInfoDef,
            Word.Selection selection,
            ref string setid,
            ref string upperClassID)
        {
            foreach (Word.Bookmark bm in bookmarks)
            {
                // ブックマーク名が「docID + bookInfoDef + 3桁の数字」の形式に一致する場合
                if (Regex.IsMatch(bm.Name, "^" + docID + bookInfoDef + @"\d{3}$"))
                {
                    // ブックマークIDを設定し、上位クラスIDとして保持
                    setid = bm.Name;
                    upperClassID = bm.Name;

                    // 行末尾にブックマークを追加する
                    selection.Bookmarks.Add(setid);
                    break;
                }
            }
        }

        // 比較結果リストから一致する情報を検索し、ブックマークを追加
        public void AddBookmarkIfMatch(
            Word.Paragraph tgtPara,
            List<CheckInfo> checkResult,
            string docID,
            string bookInfoDef)
        {
            // 行末尾を選択状態にする
            tgtPara.Range.Select();
            Word.Selection sel = Globals.ThisAddIn.Application.Selection;
            sel.EndKey(Word.WdUnits.wdLine);

            // 項番とタイトルを取得
            string num = Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "");
            string title = tgtPara.Range.Text.Trim();

            // 比較結果リストから一致する情報を検索
            var info = checkResult.FirstOrDefault(p =>
                (string.IsNullOrEmpty(p.new_num) && string.IsNullOrEmpty(num)) || p.new_num == num && p.new_title == title);

            // 一致する情報が存在する場合、ブックマークを追加
            if (info != null)
            {
                string bookmarkName = info.new_id_show.Split('(')[0].Trim().Replace("#", "♯");
                sel.Bookmarks.Add(bookmarkName);
            }
        }


        // ヘッダー行を作成してファイルに書き込む
        private void CreateHeaderFile(string headerFilePath, List<CheckInfo> checkResult, Dictionary<string, string> mergeSetId)
        {
            using (StreamWriter docinfo = new StreamWriter(headerFilePath, false, Encoding.UTF8))
            {
                // 比較結果リストをループ処理
                foreach (CheckInfo info in checkResult)
                {
                    // 新しいIDが空の場合はスキップ
                    if (string.IsNullOrEmpty(info.new_id))
                    {
                        continue;
                    }

                    // 修正候補IDを取得
                    string newIdTrimmed = info.new_id_show.Split('(')[0].Trim();

                    // ヘッダー行を作成してファイルに書き込む
                    makeHeaderLine(docinfo, mergeSetId, info.new_num, info.new_title, newIdTrimmed);
                }
            }
        }

        public bool LogAndDisplayError(Exception ex, StreamWriter log, StreamWriter swLog, loader load)
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

            // ロード画面を非表示にする
            load.Visible = false;

            // エラーメッセージを表示
            MessageBox.Show(ErrMsg);

            // ボタンを有効化して操作可能にする
            button4.Enabled = true;

            // HTML公開フラグを無効化
            blHTMLPublish = false;

            return false;
        }

        // 書誌情報と結合先情報を基に、指定されたフォーマットでテキストファイルを作成
        public void WriteBookInfoToFile(
            string rootPath,
            string headerDir,
            string docID,
            Dictionary<string, string> bookInfoDic,
            Dictionary<string, string> mergeSetId)
        {
            string filePath = Path.Combine(rootPath, headerDir, $"{docID}.txt");
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

                    HeadingInfo headingInfo = new HeadingInfo
                    {
                        num = string.IsNullOrEmpty(secText[0]) ? "" : secText[0],
                        title = string.IsNullOrEmpty(secText[1]) ? "" : secText[1],
                        id = key.Replace("♯", "#")
                    };

                    if (mergeSetId.ContainsKey(headingInfo.id))
                    {
                        headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                        makeHeaderLine(docinfo, mergeSetId, headingInfo.num, headingInfo.title, headingInfo.id);
                    }
                    else
                    {
                        docinfo.WriteLine($"{secText[0]}\t{secText[1]}\t{headingInfo.id}\t");
                    }
                }
            }
        }

        // 正規表現を使ってデータを解析し、HeadingInfo オブジェクトを生成
        private void ParseBookInfo(
            Dictionary<string, string> bookInfoDic,
            Dictionary<string, string> mergeSetId,
            List<HeadingInfo> newInfo)
        {
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
                {
                    secText[1] = bookInfoDic[key];
                }

                // 書誌情報を格納するクラスのインスタンスを作成
                HeadingInfo headingInfo = new HeadingInfo
                {
                    num = string.IsNullOrEmpty(secText[0]) ? "" : secText[0],
                    title = string.IsNullOrEmpty(secText[1]) ? "" : secText[1],
                    id = key.Contains("＃") ? key.Replace("＃", "#") : key.Replace("♯", "#")
                };

                // 結合先情報が存在する場合
                if (mergeSetId.ContainsKey(headingInfo.id))
                {
                    // 結合先IDを取得し、headingInfo.mergetoに設定
                    headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                }

                // 新しい書誌情報をリストに追加
                newInfo.Add(headingInfo);
            }
        }

        public void ProcessParagraphsInSections(
            Word.Document document,
            List<CheckInfo> checkResult,
            string docID,
            string bookInfoDef,
            ref bool breakFlg)
        {
            foreach (Word.Section tgtSect in document.Sections)
            {
                foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                {
                    // 段落のスタイル名を取得
                    string styleName = tgtPara.get_Style().NameLocal;

                    // スタイル名が「章 扉 タイトル」に一致しない、かつ「見出し」を含まない場合は次の段落へ
                    if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                    // 段落のテキストを取得
                    string innerText = tgtPara.Range.Text.Trim();

                    // 段落のテキストが空の場合は次の段落へ
                    if (string.IsNullOrWhiteSpace(innerText)) continue;

                    // 段落のテキストが「索引」に一致し、特定のスタイル名の場合、処理を中断
                    if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") &&
                        (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                    {
                        breakFlg = true;
                        break;
                    }

                    // 統合された正規表現パターン
                    string pattern = @"章[　 ]*扉.*タイトル|見出し|Heading\s*[１1-３3](?!.*目\s*次|[・用])";

                    if (Regex.IsMatch(styleName, pattern))
                    {
                        // 比較結果リストから一致する情報を検索し、ブックマークを追加
                        AddBookmarkIfMatch(tgtPara, checkResult, docID, bookInfoDef);
                    }
                }

                // 処理中断フラグが設定されている場合、セクションのループを終了
                if (breakFlg) break;
            }
        }
    }
}
