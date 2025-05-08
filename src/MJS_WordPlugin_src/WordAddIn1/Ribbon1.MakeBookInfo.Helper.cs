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
    public partial class Ribbon1
    {
        // ヘッダーファイルの確認と読み込み
        public bool CheckAndLoadHeaderFile(Word.Document Doc, loader load, int bibNum, int bibMaxNum)
        {
            string headerFilePath = Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt";

            // 指定されたヘッダーファイルが存在するか確認
            if (File.Exists(headerFilePath))
            {
                // ヘッダーファイルを開けるか確認（他のプロセスでロックされていないかチェック）
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

                oldInfo = new List<HeadingInfo>();  // 旧書誌情報
                newInfo = new List<HeadingInfo>();  // 新書誌情報
                checkResult = new List<CheckInfo>();  // 比較結果

                // ヘッダーファイルを読み込み、書誌情報番号の最大値を取得
                using (StreamReader sr = new StreamReader(headerFilePath, System.Text.Encoding.Default))
                {
                    // 書誌情報番号の最大値取得
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

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

                return true;
            }
            else
            {
                // ヘッダーファイルが存在しない場合
                return false;
            }
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
