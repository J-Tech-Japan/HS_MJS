using System;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool blHTMLPublish = false;
        private string bookInfoDef = "";
        private Dictionary<string, string> bookInfoDic = new Dictionary<string, string>();
        private bool checkOK = false;

        // リボンロード時の初期化
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            try
            {
                Globals.ThisAddIn.Application.WindowSelectionChange -=
                    new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            }
            catch { /* 既に解除済みの場合は無視 */ }

            Globals.ThisAddIn.Application.DocumentChange +=
                new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);

            // アセンブリのバージョンを取得
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            string versionText = version.ToString(3); // "1.0.0" 形式で取得

            // labelVersion はリボンデザイナで追加したラベルの名前
            versionFileJoin.Label = $"\n\nバージョン\n{versionText}";

            // 設定ボタン（button11）の表示/非表示を設定から読み込んで制御
            // デフォルトでは表示（true）
            bool showSettingsButton = ApplicationSettings.GetShowSettingsButtonSetting();
            button11.Visible = showSettingsButton;
            
            // 設定ボタンを非表示にする場合、グループ全体も非表示にする
            // （グループ内に他のボタンがない場合）
            if (!showSettingsButton)
            {
                group5.Visible = false;
            }
        }

        // KeyValuePairのValueで比較
        private static int CompareKeyValuePair(KeyValuePair<string, float> x, KeyValuePair<string, float> y)
        {
            return x.Value.CompareTo(y.Value);
        }

        // マージ情報付きのhref生成
        private static string makeHrefWithMerge(Dictionary<string, string> mergeData, string id)
        {
            if (mergeData == null || string.IsNullOrEmpty(id))
                return id + ".html";
            return mergeData.ContainsKey(id)
                ? mergeData[id] + ".html#" + id
                : id + ".html";
        }

        // ドキュメント切替時の処理
        private void Application_DocumentChange()
        {
            bookInfoDef = string.Empty;

            Word.Document activeDoc = null;
            try
            {
                activeDoc = Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch
            {
                // ドキュメントが取得できない場合は何もしない
                return;
            }

            if (activeDoc == null || activeDoc.ActiveWindow == null)
                return;

            // ブックマーク表示オプションをON
            try
            {
                activeDoc.ActiveWindow.View.ShowBookmarks = true;
            }
            catch
            {
                // 例外発生時は無視
            }

            // ヘッダーファイルのパスを生成
            string docNamePrefix = Regex.Replace(activeDoc.Name, "^(.{3}).+$", "$1");
            string headerFilePath = Path.Combine(
                Path.GetDirectoryName(activeDoc.FullName) ?? "",
                "headerFile",
                docNamePrefix + ".txt"
            );

            // ヘッダーファイルの存在チェック
            if (File.Exists(headerFilePath))
            {
                foreach (Word.Bookmark bm in activeDoc.Bookmarks)
                {
                    if (Regex.IsMatch(bm.Name, "^" + docNamePrefix))
                    {
                        // ブックマーク名から2文字を抽出
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
                button4.Enabled = true;
                button3.Enabled = false;
                button5.Enabled = false;
                button2.Enabled = false;
            }
        }

        // 指定したディレクトリ（fromPath）配下の全ファイル・サブディレクトリを別ディレクトリ（toPath）へコピー
        private void copyDirectory(string fromPath, string toPath)
        {
            // コピー元ディレクトリの情報を取得
            DirectoryInfo sourceDirectory = new DirectoryInfo(fromPath);

            // コピー元ディレクトリ内の全ファイルを取得
            FileInfo[] files = sourceDirectory.GetFiles();

            if (!Directory.Exists(toPath))
            {
                Directory.CreateDirectory(toPath);
            }

            // 各ファイルをコピー先ディレクトリにコピー（同名ファイルは上書き）
            foreach (FileInfo file in files)
            {
                file.CopyTo(Path.Combine(toPath, file.Name), true);
            }

            DirectoryInfo[] sourceSubDirectories = sourceDirectory.GetDirectories();

            // 各サブディレクトリについて再帰的にコピー処理を実行
            foreach (DirectoryInfo dir in sourceSubDirectories)
            {
                if (!Directory.Exists(Path.Combine(toPath, dir.Name)))
                {
                    Directory.CreateDirectory(Path.Combine(toPath, dir.Name));
                }
                copyDirectory(dir.FullName, Path.Combine(toPath, dir.Name));
            }
        }

        // Wordの選択範囲が変更されたときに呼び出されるイベントハンドラ。
        // スタイルチェック後にドキュメントが変更された場合、再チェックを促す。
        private void Application_WindowSelectionChange(Word.Selection ws)
        {
            // スタイルチェック直後の一度だけは何もしない（フラグをリセットして終了）
            if (checkOK)
            {
                checkOK = false;
                return;
            }

            // このイベントハンドラを一時的に解除（多重呼び出し防止）
            Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            // スタイルチェックボタンが有効な場合、ドキュメント変更を通知し再チェックを促す
            if (button3.Enabled)
            {
                string ErrMsgDocumentChanged1 = "「スタイルチェック」クリック後に変更が加えられました。\r\n「HTML出力」を実行するためには\r\nもう一度「スタイルチェック」を実行してください。";
                string ErrMsgDocumentChanged2 = "ドキュメントが変更されました！";
                MessageBox.Show(ErrMsgDocumentChanged1, ErrMsgDocumentChanged2, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button3.Enabled = false;
                return;
            }
        }

        /// <summary>
        /// 設定ボタンのイベントハンドラー
        /// </summary>
        private void SettingsButton_Click(object sender, RibbonControlEventArgs e)
        {
            try
            {
                using (var settingsForm = new SettingsForm())
                {
                    if (settingsForm.ShowDialog() == DialogResult.OK)
                    {
                        MessageBox.Show(
                            "設定を保存しました。",
                            "設定完了",
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"設定画面の表示中にエラーが発生しました。{Environment.NewLine}{ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        public List<HeadingInfo> oldInfo;  // 書誌情報（旧）
        public List<HeadingInfo> newInfo;  // 書誌情報（新）
        public List<CheckInfo> checkResult;  // 比較結果
        public int? maxNo; // MAX番号保存用 

        public Dictionary<string, string[]> title4Collection = new Dictionary<string, string[]>();
        public Dictionary<string, string[]> headerCollection = new Dictionary<string, string[]>();
    }
}
