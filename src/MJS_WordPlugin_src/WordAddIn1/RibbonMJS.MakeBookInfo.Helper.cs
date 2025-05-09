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
    }
}
