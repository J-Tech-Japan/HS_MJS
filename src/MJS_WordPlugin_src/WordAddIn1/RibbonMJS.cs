using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml.Linq;
using System.Drawing.Imaging;
using System.Threading.Tasks;
using System.Xml.Serialization;
using System.Diagnostics;
using System.Drawing;
using System.Xml;
using System.Threading;
using static System.Runtime.CompilerServices.RuntimeHelpers;
using System.Reflection.Emit;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private bool blHTMLPublish = false;
        private string bookInfoDef = "";
        private Dictionary<string, string> bookInfoDic = new Dictionary<string, string>();
        private bool checkOK = false;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = true;
            //Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions = false;
            Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //            Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
            Globals.ThisAddIn.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }

        static int CompareKeyValuePair(KeyValuePair<string, float> x, KeyValuePair<string, float> y)
        {
            return x.Value.CompareTo(y.Value);
        }

        static string makeHrefWithMerge(Dictionary<string, string> mergeData, string id)
        {
            if (mergeData.ContainsKey(id))
            {
                return mergeData[id] + ".html" + "#" + id;
            }
            else
            {
                return id + ".html";
            }
        }

        static void makeHeaderLine(StreamWriter docinfo, Dictionary<string, string> mergeSetId, string num, string title, string id)
        {
            string newId = id;
            // checked merge exiets
            if (mergeSetId.ContainsKey(id))
            {
                // check # exists
                if (mergeSetId[id].Contains("♯"))
                {
                    // get first #
                    mergeSetId[id] = mergeSetId[id].Split(new char[] { '♯' })[0];
                }

                newId = mergeSetId[id] + "♯" + id;
            }
            docinfo.WriteLine(num + "\t" + title + "\t" + id + "\t" + (mergeSetId.ContainsKey(id) ? "(" + mergeSetId[id] + ")" : ""));
        }

        private void Application_DocumentChange()
        {
            bookInfoDef = "";
            Word.Document Doc = Globals.ThisAddIn.Application.ActiveDocument;

            // ブックマーク表示オプションをオンにする
            Doc.ActiveWindow.View.ShowBookmarks = true;

            if (File.Exists(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt"))
            {
                foreach (Word.Bookmark bm in Doc.Bookmarks)
                {
                    if (Regex.IsMatch(bm.Name, "^" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1")))
                    {
                        bookInfoDef = Regex.Replace(bm.Name, "^.{3}(.{2}).*$", "$1");
                        break;
                    }
                }
                //bookInfoDef = Regex.Replace(Doc.Name, "^(.{3}).+$", "$1");
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
        private void NowLoadingProc()
        {
            alert f = new alert();
            try
            {
                f.ShowDialog();
                f.Dispose();
            }
            catch (ThreadAbortException)
            {
                f.Close();
            }
        }

        private void copyDirectory(string fromPath, string toPath)
        {
            DirectoryInfo di = new DirectoryInfo(fromPath);
            FileInfo[] files = di.GetFiles();

            if (!Directory.Exists(toPath))
            {
                Directory.CreateDirectory(toPath);
            }

            foreach (FileInfo file in files)
            {
                file.CopyTo(Path.Combine(toPath, file.Name), true);
            }

            DirectoryInfo[] dirs = di.GetDirectories();

            foreach (DirectoryInfo dir in dirs)
            {
                if (!Directory.Exists(Path.Combine(toPath, dir.Name)))
                {
                    Directory.CreateDirectory(Path.Combine(toPath, dir.Name));
                }
                copyDirectory(dir.FullName, Path.Combine(toPath, dir.Name));
            }
        }

        private void Application_WindowSelectionChange(Word.Selection ws)
        {
            if (checkOK)
            {
                checkOK = false;
                return;
            }

            Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //            Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
            if (button3.Enabled)
            {
                MessageBox.Show("「スタイルチェック」クリック後に変更が加えられました。\r\n「HTML出力」を実行するためには\r\nもう一度「スタイルチェック」を実行してください。", "ドキュメントが変更されました！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button3.Enabled = false;
                return;
            }
        }

        private string getStyleName(Dictionary<string, string> styleName, System.Xml.XmlNode seekNode)
        {
            string thisStyleName = "";

            if (seekNode.SelectSingleNode("@class") == null)
            {
                if (styleName.ContainsKey(seekNode.Name))
                {
                    thisStyleName = styleName[seekNode.Name];
                }
            }
            else
            {
                if (styleName.ContainsKey(seekNode.Name + "." + seekNode.SelectSingleNode("@class").InnerText))
                {
                    thisStyleName = styleName[seekNode.Name + "." + seekNode.SelectSingleNode("@class").InnerText];
                }
            }

            if ((thisStyleName == "") && (seekNode.SelectSingleNode("*[@class != '']") != null))
            {
                if (styleName.ContainsKey(seekNode.SelectSingleNode("*[@class != '']").Name + "." + seekNode.SelectSingleNode("*[@class != '']/@class").InnerText))
                {
                    thisStyleName = styleName[seekNode.SelectSingleNode("*[@class != '']").Name + "." + seekNode.SelectSingleNode("*[@class != '']/@class").InnerText];
                }
            }
            else if ((thisStyleName == "") && (seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']") != null))
            {
                if (styleName.ContainsKey(seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name))
                {
                    thisStyleName = styleName[seekNode.SelectSingleNode("*[translate(name(), '0123456789', '') = 'h']").Name];
                }
            }
            return thisStyleName;
        }

        // SOURCELINK追加==========================================================================START
        // 書誌情報（旧）
        public List<HeadingInfo> oldInfo;
        // 書誌情報（新）
        public List<HeadingInfo> newInfo;
        // 比較結果
        public List<CheckInfo> checkResult;
        // MAX番号保存用
        public int? maxNo;

        // Title 4 collection
        public Dictionary<string, string[]> title4Collection = new Dictionary<string, string[]>();
        public Dictionary<string, string[]> headerCollection = new Dictionary<string, string[]>();

        // SOURCELINK追加==========================================================================END


        // SOURCELINK追加==========================================================================START
        /// <summary>
        /// 新規比較処理
        /// </summary>
        /// <param name="oldInfos">書誌情報（旧）</param>
        /// <param name="newInfos">書誌情報（新）</param>
        /// <param name="checkResult">比較結果リスト</param>
        /// <returns>処理結果</returns>

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {

        }
        
        // SOURCELINK追加==========================================================================END
    }
}
