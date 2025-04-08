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
    public partial class Ribbon1
    {
        private bool blHTMLPublish = false;
        private string bookInfoDef = "";
        private Dictionary<string, string> bookInfoDic = new Dictionary<string, string>();
        private bool checkOK = false;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            //WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.TrackRevisions = true;
            //WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument.ShowRevisions = false;
            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
            WordAddIn1.Globals.ThisAddIn.Application.DocumentChange += new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
        }

        private void button1_Click(object sender, RibbonControlEventArgs e)
        {

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
            Word.Document Doc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;

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
        //    Word.Document activeDoc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;

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
        //    var activeDoc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument as Microsoft.Office.Interop.Word.Document;
        //    Word.Selection ws = WordAddIn1.Globals.ThisAddIn.Application.Selection;

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

            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            //            WordAddIn1.Globals.ThisAddIn.Application.WindowSelectionChange -= delegate (Word.Selection mySelection) { Application_WindowSelectionChange(); };
            if (button3.Enabled)
            {
                MessageBox.Show("「スタイルチェック」クリック後に変更が加えられました。\r\n「HTML出力」を実行するためには\r\nもう一度「スタイルチェック」を実行してください。", "ドキュメントが変更されました！", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
                button3.Enabled = false;
                return;
            }
        }

        private void innerNode(Dictionary<string, string> styleName, System.Xml.XmlNode objTargetNode, System.Xml.XmlNode seekNode)
        {
            string baseStyle = "";

            if (seekNode.NodeType == System.Xml.XmlNodeType.Text)
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, true));
            }
            else if (seekNode.NodeType == System.Xml.XmlNodeType.Element)
            {
                string thisStyleName = getStyleName(styleName, seekNode);

                if (seekNode.Name == "a")
                {
                    string refname = ((System.Xml.XmlElement)seekNode).GetAttribute("name");
                    if (refname.Contains("_Ref"))
                    {
                        objTargetNode.AppendChild(objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("span")));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("name", refname);
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "ref");
                    }
                }

                if ((seekNode.Name == "table") || (seekNode.Name == "tr") || (seekNode.Name == "td"))
                {
                    if (Regex.IsMatch(thisStyleName, "参照先"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));

                        baseStyle = "";
                        if (objTargetNode.LastChild.SelectNodes("@style").Count != 0)
                        {
                            baseStyle = ((System.Xml.XmlElement)objTargetNode.LastChild).GetAttribute("style");
                        }
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "text-align:right; font-size:90%;" + baseStyle);
                    }
                    else if (seekNode.Name == "table")
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        string thisStyle = ((System.Xml.XmlElement)objTargetNode.LastChild).GetAttribute("style");
                        if (Regex.IsMatch(seekNode.InnerXml, @"border-top:double[\s　]+#[^\s|　]+?[\s　]+4.5pt") && Regex.IsMatch(seekNode.InnerXml, @"border-bottom:double[\s　]+#[^\s|　]+?[\s　]+4.5pt"))
                        {
                            if (Regex.IsMatch(thisStyle, @"(?<![A-z\d-])width:"))
                            {
                                ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", Regex.Replace(thisStyle, @"(?<![A-z\d-])width:[^;]+;", ""));
                            }
                        }
                        else if (Regex.IsMatch(thisStyle, @"(?<![A-z\d-])width:"))
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", Regex.Replace(thisStyle, @"(?<![A-z\d-])(?<=width:)[\d\.]+",
                            Convert.ToString(Math.Round(decimal.Parse(Regex.Replace(thisStyle, @"^.*?width:([\d\.]+)\w+;.*?$", "$1")) * 1.15m, 1, MidpointRounding.AwayFromZero))));
                        }

                    }
                    else if (seekNode.Name == "td")
                    {
                        foreach (System.Xml.XmlNode childs in seekNode.SelectNodes(".//*[boolean(@class)]"))
                        {
                            if (Regex.IsMatch(getStyleName(styleName, childs), "コラム.*アイコン"))
                            {
                                ((System.Xml.XmlElement)seekNode).SetAttribute("width", "80");
                                break;
                            }

                            else if (Regex.IsMatch(getStyleName(styleName, childs), "コラム"))
                            {
                                ((System.Xml.XmlElement)seekNode).RemoveAttribute("width");
                                if (seekNode.SelectNodes("@style").Count != 0)
                                {
                                    ((System.Xml.XmlElement)seekNode).SetAttribute("style", Regex.Replace(((System.Xml.XmlElement)seekNode).GetAttribute("style"), "(?:^| )width:[^;]+;", ""));
                                }
                                break;
                            }
                            else if (childs.Name == "table")
                            {
                                XmlNode divNode = childs.OwnerDocument.CreateElement("div");
                                divNode.Attributes.Append(childs.OwnerDocument.CreateAttribute("class")).Value = "Q＆A";


                                foreach (System.Xml.XmlNode trNode in childs.SelectNodes(".//tr"))
                                {
                                    if (trNode.SelectNodes(".//p[@class='MJS-QA']").Count != 0)
                                    {
                                        XmlNode qBlockDivNode = divNode.OwnerDocument.CreateElement("div");
                                        foreach (System.Xml.XmlNode childNode in trNode.SelectNodes(".//td"))
                                        {
                                            XmlNodeList pNodes = childNode.SelectNodes(".//p");
                                            foreach (System.Xml.XmlNode pNode in pNodes)
                                            {
                                                if (pNode.SelectNodes("@class[. = 'MJS-QA']").Count == 0)
                                                {
                                                    qBlockDivNode.AppendChild(pNode);
                                                }
                                            }
                                            ((System.Xml.XmlElement)qBlockDivNode).SetAttribute("class", "MJS_qa_td_Qblock");
                                            divNode.AppendChild(qBlockDivNode);

                                        }
                                    }
                                    else if (trNode.SelectNodes(".//p[@class='MJSQAA']").Count != 0)
                                    {
                                        XmlNode aBlockDivNode = divNode.OwnerDocument.CreateElement("div");
                                        foreach (System.Xml.XmlNode childNode in trNode.SelectNodes(".//td"))
                                        {
                                            XmlNodeList pNodes = childNode.SelectNodes(".//p");
                                            foreach (System.Xml.XmlNode pNode in pNodes)
                                            {
                                                if (pNode.SelectNodes("@class[. = 'MsoNormal']").Count == 0)
                                                {
                                                    aBlockDivNode.AppendChild(pNode);
                                                }
                                            }
                                           ((System.Xml.XmlElement)aBlockDivNode).SetAttribute("class", "MJS_qa_td_Ablock");
                                            divNode.AppendChild(aBlockDivNode);

                                        }
                                    }
                                }
                                childs.ParentNode.ReplaceChild(divNode, childs);
                                break;
                            }
                        }
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                    }
                    else
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                    }
                }
                else if (Regex.IsMatch(seekNode.Name, @"^(?:[bui]|ins|h\d+)$", RegexOptions.IgnoreCase) ||
                        ((objTargetNode.SelectSingleNode("ancestor-or-self::*[starts-with(@class, 'Heading')]") != null) && (seekNode.Name == "p")) ||
                       (seekNode.SelectNodes("@class[. = 'msoIns']").Count != 0))
                {
                    foreach (System.Xml.XmlNode Child in seekNode.ChildNodes)
                    {
                        innerNode(styleName, objTargetNode, Child);
                    }
                    return;
                }
                else
                {
                    if (Regex.IsMatch(thisStyleName, "表見出し"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading_table");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'Heading_table']").Count == 0))
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading_table");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_画像（操作の流れ）"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_oflowPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順番号"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_oflow_stepNum");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順結果"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_oflow_stepResult");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_oflow_step");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_補足"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_oflow_sub");
                    }

                    else if (seekNode.SelectNodes("@class[. = 'MJS_qa_td_Qblock']").Count != 0)
                    {

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_td_Qblock");

                    }
                    else if (seekNode.SelectNodes("@class[. = 'MJS_qa_td_Ablock']").Count != 0)
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_td_Ablock");

                    }
                    else if (seekNode.SelectNodes("@class[. = 'Q＆A']").Count != 0)
                    {

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "Q＆A");

                    }

                    //' Ver - 2023.16.08 - VyNL - ↓ - 追加'
                    else if (Regex.IsMatch(thisStyleName, "Q＆A"))
                    {

                        if (thisStyleName == "MJS_Q＆A_Q")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_Q");

                        }
                        else if (Regex.IsMatch(thisStyleName, "アイコン"))
                        {
                            if (seekNode.SelectNodes(".//img").Count != 0)
                            {
                                foreach (System.Xml.XmlNode Icon in seekNode.SelectNodes(".//img"))
                                {
                                    ((System.Xml.XmlElement)Icon).SetAttribute("width", "80");
                                }
                            }

                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_icons");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        }
                        else if (thisStyleName == "MJS_Q＆A_A")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                            if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                            {
                                ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                            }

                            if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_qa_A']").Count == 0))
                            {
                                ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_A");
                            }

                        }
                        else if (thisStyleName == "MJS_Q＆A_A継続")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "padding-left: 40px;");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_A_cont");

                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");

                        }

                    }
                    //' Ver - 2023.16.08 - VyNL - ↑ - 追加'
                    // 3a
                    else if (Regex.IsMatch(thisStyleName, "リード文"))
                    {
                        if (Regex.IsMatch(thisStyleName, "リード文.*[1１2２]"))
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_leadSentence1_2");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        }
                        else if (Regex.IsMatch(thisStyleName, "リード文.*[3３4４]"))
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_leadSentence3_4");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_lead");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        }
                    }

                    else if (Regex.IsMatch(thisStyleName, "下線"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_ul");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "処理フロー"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (seekNode.SelectNodes("ancestor::*[@class = 'MJS_flow']").Count == 0)
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_flow");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "参照先"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_ref");
                    }
                    else if (Regex.IsMatch(thisStyleName, "選択肢等[2２]"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choice2");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "選択肢等"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choice");
                    }
                    else if (Regex.IsMatch(thisStyleName, "選択肢-説明等[2２]"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choiceDesc2");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "選択肢.*説明等"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_choiceDesc']").Count == 0))
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choiceDesc");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "箇条書き[2２]"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_listItem2']").Count == 0))
                        {
                            seekNode.InnerText = Regex.Replace(seekNode.InnerText, @"^\S{0,3}[ 　]+", "");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_listItem2");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "箇条書き"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_listItem']").Count == 0))
                        {
                            seekNode.InnerText = Regex.Replace(seekNode.InnerText, @"^\S{0,3}[ 　]+", "");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_listItem");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内-項目_センタリング"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem_center");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内-項目_右寄せ"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem_right");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*タイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*箇条"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_tdListItem']").Count == 0))
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdListItem");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*項目"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*本文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_tdText']").Count == 0))
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdText");
                        }
                        else
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "(?:コラム.*アイコン|事項.*アイコン用?)"))
                    {
                        if (seekNode.SelectNodes(".//img").Count != 0)
                        {
                            foreach (System.Xml.XmlNode Icon in seekNode.SelectNodes(".//img"))
                            {
                                ((System.Xml.XmlElement)Icon).SetAttribute("width", "80");
                            }
                        }

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_icons");
                    }
                    else if (Regex.IsMatch(thisStyleName, "コラム.*本文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (seekNode.SelectNodes("ancestor::*[@class = 'MJS_columnText']").Count == 0)
                        {
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnText");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "コラム.*タイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "見出し.*手順"))
                    {
                        if (seekNode.SelectNodes(".//img").Count != 0)
                        {
                            ((System.Xml.XmlElement)seekNode.SelectSingleNode(".//img")).SetAttribute("width", "35");
                            if (seekNode.SelectSingleNode(".//img").NextSibling.InnerText == "　")
                            {
                                seekNode.SelectSingleNode(".//img").NextSibling.InnerText = ((char)160).ToString();
                            }
                        }
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*本文内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_bodyPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*表内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*コラム内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        //((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "margin-left: 15mm;");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*手順内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_flowPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "メニュー[2２]"))
                    {
                        if (seekNode.Name == "p")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_Menu2");
                        }
                        else
                        {
                            foreach (System.Xml.XmlNode Child in seekNode.ChildNodes)
                            {
                                innerNode(styleName, objTargetNode, Child);
                            }
                            return;
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "メニュー"))
                    {
                        if (seekNode.Name == "p")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_menu");
                        }
                        else
                        {
                            foreach (System.Xml.XmlNode Child in seekNode.ChildNodes)
                            {
                                innerNode(styleName, objTargetNode, Child);
                            }
                            return;
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順結果"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processResult");
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順番号リセット用"))
                    {
                        return;
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_process");
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順補足"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processSuppl");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルタイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_title");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルサブタイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_subtitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルバージョン"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_version");
                    }
                    else if ((objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_process') or starts-with(@class, 'MJS_process')]").Count != 0) &&
                             (seekNode.SelectNodes("@style[contains(., 'color:#1F497D')]").Count != 0))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        string bold = "";
                        if (seekNode.SelectNodes("ancestor-or-self::b").Count != 0) bold = "font-weight:bold;";
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#1F497D;" + bold);
                    }

                    else if (Regex.IsMatch(thisStyleName, "ui"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if ((seekNode.Name == "p") && (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "") && (seekNode.SelectNodes(".//img").Count != 0))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "margin-left:2em;");
                    }
                    else if ((seekNode.Name == "span") && (seekNode.ChildNodes.Count == 0))
                    {
                        return;
                    }
                    else if (seekNode.SelectNodes("@class[. = 'MsoHyperlink']").Count != 0)
                    {
                        foreach (System.Xml.XmlNode Child in seekNode.ChildNodes)
                        {
                            innerNode(styleName, objTargetNode, Child);
                        }
                        return;
                    }

                    else if (seekNode.SelectNodes("@style[contains(translate(., ' ', ''), 'font-family:Wingdings')]").Count != 0)
                    {
                        if (Regex.IsMatch(seekNode.InnerText, @"\u009F"))
                        {
                            return;
                        }
                        else if (objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'Heading_table') or (@class = 'Heading4') or (@class = 'MJS_processResult')]").Count != 0)
                        {
                            return;
                        }

                        else if ((objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_process') or starts-with(@class, 'MJS_process')]").Count != 0) &&
                                Regex.IsMatch(((System.Xml.XmlElement)seekNode).GetAttribute("style"), @"(?<![A-z\d-])color:"))
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            string thisColor = Regex.Replace(((System.Xml.XmlElement)objTargetNode.LastChild).GetAttribute("style"), @"^.*(?<![A-z\d-])(color:[^;]+;).+$", "$1");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "font-family:Wingdings;" + thisColor);
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "font-family:Wingdings");
                        }
                    }// 3a
                    else if (seekNode.SelectNodes("@style[contains(translate(., ' ', ''), 'color:#246A98;font-weight:normal')]").Count != 0)
                    {
                        if (objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_qa_Q')]").Count != 0)
                        {
                            return;
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#246A98;font-weight:normal");
                        }
                    }
                    else if (seekNode.SelectNodes("@style[contains(translate(., ' ', ''), 'color:#8EAADB')]").Count != 0)
                    {
                        if (objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_qa_A')]").Count != 0)
                        {
                            return;
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#8EAADB");
                        }
                    }

                    else if (Regex.IsMatch(thisStyleName, "タブ見出し"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading4_tab");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表脚注"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tableFootnote");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }

                    else
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((System.Xml.XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                    }
                }

                foreach (System.Xml.XmlNode Child in seekNode.ChildNodes)
                {
                    innerNode(styleName, objTargetNode.LastChild, Child);
                }
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
        


        

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {

        }
        // SOURCELINK追加==========================================================================END
    }
}
