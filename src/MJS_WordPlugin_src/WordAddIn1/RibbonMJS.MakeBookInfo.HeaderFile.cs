using System;
using System.Collections.Generic;
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
        // ヘッダーファイルの確認と読み込み
        public bool CheckAndLoadHeaderFile(Word.Document doc, loader load, int bibNum, int bibMaxNum)
        {
            string headerFilePath = Path.Combine(
                Path.GetDirectoryName(doc.FullName) ?? string.Empty,
                "headerFile",
                $"{Regex.Replace(doc.Name, "^(.{3}).+$", "$1")}.txt"
            );

            // ヘッダーファイルが存在するか確認
            if (!File.Exists(headerFilePath))
            {
                return false;
            }

            // ヘッダーファイルを開けるか確認
            if (!IsFileAccessible(headerFilePath, load))
            {
                return false;
            }

            oldInfo = new List<HeadingInfo>();
            newInfo = new List<HeadingInfo>();
            checkResult = new List<CheckInfo>();

            // ヘッダーファイルを読み込み、書誌情報番号の最大値を取得
            StreamReader sr = null;
            try
            {
                sr = new StreamReader(headerFilePath, Encoding.Default);
                while (sr.Peek() >= 0)
                {
                    string[] info = sr.ReadLine()?.Split('\t') ?? Array.Empty<string>();
                    if (info.Length < 3) continue;

                    var headingInfo = new HeadingInfo
                    {
                        num = info[0],
                        title = info[1],
                        id = info[2],
                        mergeto = info.Length == 4 ? info[3] : null
                    };

                    oldInfo.Add(headingInfo);

                    if (int.TryParse(info[2].Substring(info[2].Length - 3, 3), out int currentBibNum) && bibMaxNum < currentBibNum)
                    {
                        bibMaxNum = currentBibNum;
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ヘッダーファイルの読み込み中にエラーが発生しました: {ex.Message}",
                    "読み込みエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
            finally
            {
                sr?.Dispose();
            }

            // ドキュメント内のブックマークを確認し、書誌情報のデフォルト値を取得
            bookInfoDef = doc.Bookmarks
                .Cast<Word.Bookmark>()
                .FirstOrDefault(bm => Regex.IsMatch(bm.Name, $"^{Regex.Replace(doc.Name, "^(.{3}).+$", "$1")}"))?
                .Name.Substring(3, 2);

            return true;

            // ローカル関数: ファイルがアクセス可能か確認
            bool IsFileAccessible(string filePath, loader loaderInstance)
            {
                try
                {
                    using (var stream = new FileStream(filePath, FileMode.Open, FileAccess.Read, FileShare.None))
                    {
                        return true;
                    }
                }
                catch (IOException)
                {
                    loaderInstance.Visible = false;
                    MessageBox.Show($"{filePath}が開かれています。\r\nファイルを閉じてから書誌情報出力を実行してください。",
                        "ファイルエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.DoEvents();
                    Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return false;
                }
            }
        }
    }
}
