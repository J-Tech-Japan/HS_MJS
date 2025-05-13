using Microsoft.Office.Tools.Ribbon;

using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

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

        // ヘッダー行の出力
        private static void makeHeaderLine(StreamWriter docinfo, Dictionary<string, string> mergeSetId, string num, string title, string id)
        {
            string newId = id;
            if (mergeSetId != null && mergeSetId.ContainsKey(id))
            {
                // "♯"が含まれていれば最初の部分だけ取得
                if (mergeSetId[id].Contains("♯"))
                {
                    mergeSetId[id] = mergeSetId[id].Split(new char[] { '♯' })[0];
                }
                newId = mergeSetId[id] + "♯" + id;
            }
            docinfo.WriteLine($"{num}\t{title}\t{id}\t{(mergeSetId != null && mergeSetId.ContainsKey(id) ? "(" + mergeSetId[id] + ")" : "")}");
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

        //private void Application_DocumentBeforeSave(Word.Document Doc, ref bool SaveAsUI, ref bool Cancel)
        //{
        //    Word.Document activeDoc = Globals.ThisAddIn.Application.ActiveDocument;

        //    if (!File.Exists(activeDoc.Path + "\\" + Path.ChangeExtension(activeDoc.Name, ".h")))
        //            File.CreateText(activeDoc.Path + "\\" + Path.ChangeExtension(activeDoc.Name, ".h"));
        //    using (StreamReader sr = new StreamReader(activeDoc.Path + "\\" + Path.ChangeExtension(activeDoc.Name, ".h"), Encoding.UTF8))
        //    {

        //    }

        //    foreach (Word.Paragraph wp in activeDoc.Paragraphs)
        //    {
        //    }

        //}

        //private void toggleButton1_Click(object sender, RibbonControlEventArgs e)
        //{
        //    if (toggleButton1.Checked == true)
        //        button2.Enabled = true;
        //    else button2.Enabled = false;
        //    var activeDoc = Globals.ThisAddIn.Application.ActiveDocument as Microsoft.Office.Interop.Word.Document;
        //    Word.Selection ws = Globals.ThisAddIn.Application.Selection;

        //    if (toggleButton1.Checked)
        //    {
        //        activeDoc.TrackRevisions = true;
        //        activeDoc.ShowRevisions = false;
        //    }

        //    if (!toggleButton1.Checked)
        //    {
        //        activeDoc.TrackRevisions = false;
        //        activeDoc.ShowRevisions = true;
        //    }
        //}

        /*以下は、次期対応変更履歴保存用コードの一部です。
        private string cordConvert(int i)
        {
            string rireki = "";
            switch (i)
            {
                case 1:
                    rireki = "挿入";
                    break;
                case 2:
                    rireki = "削除";
                    break;
                case 3:
                    rireki = "プロパティの変更";
                    break;
                case 4:
                    rireki = "段落番号の変更";
                    break;
                case 5:
                    rireki = "フィールド表示の変更";
                    break;
                case 6:
                    rireki = "解決された競合";
                    break;
                case 7:
                    rireki = "競合";
                    break;
                case 8:
                    rireki = "スタイルの変更";
                    break;
                case 9:
                    rireki = "置換";
                    break;
                case 10:
                    rireki = "段落のプロパティの変更";
                    break;
                case 11:
                    rireki = "表のプロパティの変更";
                    break;
                case 12:
                    rireki = "セクションのプロパティの変更";
                    break;
                case 13:
                    rireki = "スタイル定義の変更";
                    break;
                case 14:
                    rireki = "内容の移動元";
                    break;
                case 15:
                    rireki = "内容の移動先";
                    break;
                case 16:
                    rireki = "表のセルの挿入";
                    break;
                case 17:
                    rireki = "表のセルの削除";
                    break;
                case 18:
                    rireki = "表のセルの結合";
                    break;
            }
            return rireki;
        }
        */

        //private void NowLoadingProc()
        //{
        //    alert f = new alert();
        //    try
        //    {
        //        f.ShowDialog();
        //        f.Dispose();
        //    }
        //    catch (ThreadAbortException)
        //    {
        //        f.Close();
        //    }
        //}

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
                MessageBox.Show(ErrMsgDocumentChanged1, ErrMsgDocumentChanged2, MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button3.Enabled = false;
                return;
            }
        }

        // 指定ノードのスタイル名をディクショナリから取得する。
        // class属性や子要素のclass属性も考慮する。
        private string getStyleName(Dictionary<string, string> styleName, System.Xml.XmlNode seekNode)
        {
            if (styleName == null || seekNode == null)
                return string.Empty;

            // 1.ノード自身のclass属性を確認
            var classAttr = seekNode.SelectSingleNode("@class");
            if (classAttr != null)
            {
                string key = seekNode.Name + "." + classAttr.InnerText;
                if (styleName.TryGetValue(key, out string value))
                    return value;
            }
            else
            {
                // class属性がなければノード名のみで検索
                if (styleName.TryGetValue(seekNode.Name, out string value))
                    return value;
            }

            // 2.子ノードでclass属性を持つものを探す
            var childWithClass = seekNode.SelectSingleNode("*[@class != '']");
            if (childWithClass != null)
            {
                var childClassAttr = childWithClass.SelectSingleNode("@class");
                if (childClassAttr != null)
                {
                    string key = childWithClass.Name + "." + childClassAttr.InnerText;
                    if (styleName.TryGetValue(key, out string value))
                        return value;
                }
            }

            // 3.子ノードで名前がh*（h1, h2, ...）のものを探す
            var headingNode = seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']");
            if (headingNode != null)
            {
                if (styleName.TryGetValue(headingNode.Name, out string value))
                    return value;
            }

            // どれにも該当しない場合は空文字列
            return string.Empty;
        }

        //private string getStyleName(Dictionary<string, string> styleName, System.Xml.XmlNode seekNode)
        //{
        //    string thisStyleName = "";

        //    if (seekNode.SelectSingleNode("@class") == null)
        //    {
        //        if (styleName.ContainsKey(seekNode.Name))
        //        {
        //            thisStyleName = styleName[seekNode.Name];
        //        }
        //    }
        //    else
        //    {
        //        if (styleName.ContainsKey(seekNode.Name + "." + seekNode.SelectSingleNode("@class").InnerText))
        //        {
        //            thisStyleName = styleName[seekNode.Name + "." + seekNode.SelectSingleNode("@class").InnerText];
        //        }
        //    }

        //    if ((thisStyleName == "") && (seekNode.SelectSingleNode("*[@class != '']") != null))
        //    {
        //        if (styleName.ContainsKey(seekNode.SelectSingleNode("*[@class != '']").Name + "." + seekNode.SelectSingleNode("*[@class != '']/@class").InnerText))
        //        {
        //            thisStyleName = styleName[seekNode.SelectSingleNode("*[@class != '']").Name + "." + seekNode.SelectSingleNode("*[@class != '']/@class").InnerText];
        //        }
        //    }
        //    else if ((thisStyleName == "") && (seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']") != null))
        //    {
        //        if (styleName.ContainsKey(seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name))
        //        {
        //            thisStyleName = styleName[seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name];
        //        }
        //    }
        //    return thisStyleName;
        //}

        public List<HeadingInfo> oldInfo;  // 書誌情報（旧）
        public List<HeadingInfo> newInfo;  // 書誌情報（新）
        public List<CheckInfo> checkResult;  // 比較結果
        public int? maxNo; // MAX番号保存用 

        public Dictionary<string, string[]> title4Collection = new Dictionary<string, string[]>();
        public Dictionary<string, string[]> headerCollection = new Dictionary<string, string[]>();
    }
}
