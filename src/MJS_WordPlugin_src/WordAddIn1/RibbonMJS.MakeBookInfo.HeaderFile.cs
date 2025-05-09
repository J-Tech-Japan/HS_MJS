using System;
using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
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
    }
}
