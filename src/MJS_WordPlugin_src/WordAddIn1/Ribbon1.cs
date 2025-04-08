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

        private bool makeBookInfo(loader load, StreamWriter swLog = null)
        {
            WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = false;
            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;

            // ファイル命名規則チェック
            if (!Regex.IsMatch(thisDocument.Name, @"^[A-Z]{3}(_[^_]*?){2}\.docx*$"))
            {
                load.Visible = false;
                MessageBox.Show("開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」", "ファイル命名規則エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Application.DoEvents();
                WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
                return false;
            }

            int selStart = WordAddIn1.Globals.ThisAddIn.Application.Selection.Start;
            int selEnd = WordAddIn1.Globals.ThisAddIn.Application.Selection.End;
            WordAddIn1.Globals.ThisAddIn.Application.Selection.EndKey(Word.WdUnits.wdStory);
            Application.DoEvents();
            WordAddIn1.Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
            Application.DoEvents();

            if (WordAddIn1.Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionInlineShape ||
                WordAddIn1.Globals.ThisAddIn.Application.Selection.Type == Word.WdSelectionType.wdSelectionShape)
                WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveLeft(Word.WdUnits.wdCharacter);

            bookInfoDef = "";
            Word.Document Doc = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            // 書誌情報番号
            int bibNum = 0;
            // 書誌情報番号最大値
            int bibMaxNum = 0;

            bool checkBL = false;

            if (File.Exists(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt"))
            {
                try
                {
                    using (Stream stream = new FileStream(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt", FileMode.Open))
                    {
                    }
                }
                catch
                {
                    load.Visible = false;
                    MessageBox.Show(Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt" + "が開かれています。\r\nファイルを閉じてから書誌情報出力を実行してください。",
                        "ファイルエラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    Application.DoEvents();
                    WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
                    return false;
                }

                // SOURCELINK追加==========================================================================START
                // 書誌情報（旧）
                oldInfo = new List<HeadingInfo>();
                // 書誌情報（新）
                newInfo = new List<HeadingInfo>();
                // 比較結果
                checkResult = new List<CheckInfo>();
                // SOURCELINK追加==========================================================================END

                using (StreamReader sr = new StreamReader(
                    Path.GetDirectoryName(Doc.FullName) + "\\headerFile\\" + Regex.Replace(Doc.Name, "^(.{3}).+$", "$1") + @".txt", System.Text.Encoding.Default))
                {
                    // 書誌情報番号の最大値取得
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

                        // SOURCELINK追加==========================================================================START
                        string[] info = strBuffer.Split('\t');

                        HeadingInfo headingInfo = new HeadingInfo();
                        headingInfo.num = info[0];
                        headingInfo.title = info[1];
                        if (info.Length == 4)
                        {
                            headingInfo.mergeto = info[3];
                        }
                        headingInfo.id = info[2];

                        oldInfo.Add(headingInfo);

                        // SOURCELINK追加==========================================================================END

                        bibNum = int.Parse(info[2].Substring(info[2].Length - 3, 3));
                        if (bibMaxNum < bibNum)
                        {
                            bibMaxNum = bibNum;
                        }
                    }
                }

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
                button3.Enabled = false;
                button2.Enabled = false;
                button5.Enabled = false;
                checkBL = true;
            }

            string rootPath = thisDocument.Path;
            string docName = thisDocument.Name;
            string headerDir = "headerFile";

            string docid = Regex.Replace(docName, "^(.{3}).+$", "$1");
            string docTitle = Regex.Replace(docName, @"^.{3}_?(.+?)(?:_.+)?\.[^\.]+$", "$1");
            bookInfoDic.Clear();

            //string headerFileName = docid + ".h";

            StreamWriter log = swLog;

            if (swLog == null)
            {
                log = new StreamWriter(rootPath + "\\log.txt", false, Encoding.UTF8);
            }

            try
            {
                if (bookInfoDef == "")
                {

                    foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();
                    using (bookInfo bi = new bookInfo())
                    {
                        if (bi.ShowDialog() == DialogResult.OK)
                        {
                            bookInfoDef = bi.tbxDefaultValue.Text;
                        }
                        else
                        {
                            log.Close();
                            if (File.Exists(rootPath + "\\log.txt")) File.Delete(rootPath + "\\log.txt");
                            button4.Enabled = true;
                            return false;
                        }
                    }
                }

                Dictionary<string, string> oldBookInfoDic = new Dictionary<string, string>();
                HashSet<string> ls = new HashSet<string>();

                if (!Directory.Exists(rootPath + "\\" + headerDir))
                {
                    Directory.CreateDirectory(rootPath + "\\" + headerDir);
                }
                //foreach (string docInfo in Directory.GetFiles(rootPath + "\\" + headerDir, "*.txt"))
                //{
                //    using (StreamReader sr = new StreamReader(docInfo))
                //    {
                //        while (!sr.EndOfStream)
                //        {
                //            string[] lineText = sr.ReadLine().Split('\t');

                //            if ((lineText.Length == 3) && Regex.IsMatch(lineText[2], @"^[A-Z]{3}\d+$") || Regex.IsMatch(lineText[2], @"^[A-Z]{3}\d+#[A-Z]{3}\d+$"))
                //            {
                //                oldBookInfoDic.Add(lineText[2], lineText[1]);
                //                try { ls.Add(lineText[2].Substring(lineText[2].Length - 3, 3)); }
                //                catch { }
                //            }
                //        }
                //    }
                //}

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
                        Console.WriteLine(e);
                    }
                }

                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                {
                    foreach (Word.Bookmark wbInWb in wb.Range.Bookmarks)
                    {
                        if (!Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") && !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}♯" + docid + bookInfoDef + @"\d{3}$") &&
                            !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") && !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}＃" + docid + bookInfoDef + @"\d{3}$"))
                            wbInWb.Delete();
                    }
                }

                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                {
                    foreach (Word.Bookmark wbInWb in wb.Range.Bookmarks)
                    {
                        if (!Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") && !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}♯" + docid + bookInfoDef + @"\d{3}$") &&
                            !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}$") && !Regex.IsMatch(wbInWb.Name, @"^" + docid + bookInfoDef + @"\d{3}＃" + docid + bookInfoDef + @"\d{3}$"))
                            wbInWb.Delete();
                    }
                }

                foreach (Word.Bookmark wb in thisDocument.Bookmarks)
                {
                    if (!ls.Contains(wb.Name.Substring(wb.Name.Length - 3, 3)))
                        ls.Add(wb.Name.Substring(wb.Name.Length - 3, 3));
                    else
                        wb.Delete();
                }
                if (ls.Count != 0)
                {
                    string maxResult = ls.Max(val => val);
                    if (int.Parse(maxResult) > bibMaxNum) bibMaxNum = int.Parse(maxResult);
                }

                maxNo = bibMaxNum;

                int splitCount = 1;

                string lv1styleName = "";
                string lv2styleName = "";
                string lv3styleName = "";

                int lv1count = 0;
                int lv2count = 0;
                int lv3count = 0;

                bool breakFlg = false;

                if (!bookInfoDic.ContainsKey(docid + "00000"))
                {
                    bookInfoDic.Add(docid + "00000", "表紙");
                }

                log.WriteLine("書誌情報リスト作成開始");
                string upperClassID = "";
                string previousSetId = "";
                bool isMerge = false;
                Dictionary<string, string> mergeSetId = new Dictionary<string, string>();
                title4Collection = new Dictionary<string, string[]>();
                headerCollection = new Dictionary<string, string[]>();

                foreach (Word.Section tgtSect in thisDocument.Sections)
                {
                    foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                    {
                        string styleName = tgtPara.get_Style().NameLocal;

                        if (styleName.Equals("MJS_参照先"))
                        {
                            foreach (Word.Field fld in tgtPara.Range.Fields)
                            {
                                if (fld.Type == Word.WdFieldType.wdFieldRef)
                                {
                                    string bookmarkName = fld.Code.Text.Split(new char[] { ' ' })[2] + "_ref";
                                    tgtPara.Range.Bookmarks.Add(bookmarkName);
                                    fld.Code.Text = "HYPERLINK " + fld.Code.Text.Split(new char[] { ' ' })[2];
                                }
                            }
                        }

                        isMerge = false;

                        try
                        {
                            string styleCharacterName = tgtPara.Range.CharacterStyle.NameLocal;
                            if (styleCharacterName.Equals("MJS_見出し結合用"))
                            {
                                isMerge = true;
                            }
                        }
                        catch (Exception) { }


                        if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[４4](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[５5](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                            || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            tgtPara.Range.Bookmarks.ShowHidden = true;

                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (!title4Collection.ContainsKey(bm.Name))
                                {
                                    if (bm.Name.IndexOf("_Ref") == 0)
                                    {
                                        title4Collection.Add(bm.Name, new string[] { upperClassID, tgtPara.Range.Text.Replace("\r", "").Replace("\n", "").Replace("\"", "\"\"") });
                                    }
                                }
                            }
                            tgtPara.Range.Bookmarks.ShowHidden = false;
                        }
                        //if (!Regex.IsMatch(styleName, @"(見出し|Heading)\s*[４4](?![・用])")
                        //    && (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?![・用])")
                        //    || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])")
                        //    || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])")))
                        //{
                        //    tgtPara.Range.Bookmarks.ShowHidden = true;

                        //    foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                        //    {
                        //        if (!headerCollection.ContainsKey(bm.Name))
                        //        {
                        //            if (bm.Name.IndexOf("_Ref") == 0)
                        //            {
                        //                headerCollection.Add(bm.Name, new string[] { upperClassID, tgtPara.Range.Text.Replace("\r", "").Replace("\n", "").Replace("\"", "\"\"") });
                        //            }
                        //        }
                        //    }
                        //    tgtPara.Range.Bookmarks.ShowHidden = false;
                        //}

                        if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                        string innerText = tgtPara.Range.Text.Trim();

                        if (tgtPara.Range.Text.Trim() == "") continue;

                        if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                        {
                            breakFlg = true;
                            break;
                        }

                        if (Regex.IsMatch(styleName, @"章[　 ]*扉.*タイトル"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                //while (ls.Contains(splitCount.ToString("000")))
                                //{
                                //    splitCount++;
                                //}
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = setid;

                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                //tgtPara.Range.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                //splitCount++;
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }

                            lv1count++;
                            lv2styleName = "";
                            lv2count = 0;
                            lv3styleName = "";
                            lv3count = 0;

                            lv1styleName = styleName;
                        }
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$"))
                        {
                            Application.DoEvents();
                            if (!Regex.IsMatch(innerText, @"目\s*次\s*$"))
                            {
                                // 行末尾を選択状態にする
                                tgtPara.Range.Select();
                                Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                                sel.EndKey(Word.WdUnits.wdLine);

                                string setid = "";
                                foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                                {
                                    if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                    {
                                        setid = bm.Name;
                                        upperClassID = bm.Name;

                                        // 行末尾にブックマークを追加する
                                        sel.Bookmarks.Add(setid);

                                        break;
                                    }
                                }

                                if (setid == "")
                                {
                                    //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                    //while (ls.Contains(splitCount.ToString("000")))
                                    //{
                                    //    splitCount++;
                                    //}
                                    bibMaxNum++;
                                    splitCount = bibMaxNum;
                                    ls.Add(splitCount.ToString("000"));
                                    setid = docid + bookInfoDef + splitCount.ToString("000");
                                    upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                    //tgtPara.Range.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                    //splitCount++;
                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }
                                    previousSetId = setid;
                                }
                                else if (!bookInfoDic.ContainsKey(setid))
                                {
                                    bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                    if (isMerge)
                                    {
                                        mergeSetId.Add(setid, previousSetId);
                                    }
                                    previousSetId = setid;
                                }

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
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[２2](?![・用])"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = bm.Name;
                                    upperClassID = bm.Name;

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);

                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                //while (ls.Contains(splitCount.ToString("000")))
                                //{
                                //    splitCount++;
                                //}
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = docid + bookInfoDef + splitCount.ToString("000");
                                upperClassID = docid + bookInfoDef + splitCount.ToString("000");

                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));

                                //tgtPara.Range.Bookmarks.Add(docid + bookInfoDef + splitCount.ToString("000"));
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                //splitCount++;
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
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
                            else
                            {
                                lv3count++;
                                lv3styleName = styleName;
                            }
                        }
                        else if (Regex.IsMatch(styleName, @"(見出し|Heading)\s*[３3](?![・用])"))
                        {
                            Application.DoEvents();

                            // 行末尾を選択状態にする
                            tgtPara.Range.Select();
                            Word.Selection sel = WordAddIn1.Globals.ThisAddIn.Application.Selection;
                            sel.EndKey(Word.WdUnits.wdLine);

                            string setid = "";
                            foreach (Word.Bookmark bm in tgtPara.Range.Bookmarks)
                            {
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "♯" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(♯.*?)$", "$1");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                                if (Regex.IsMatch(bm.Name, "^" + docid + bookInfoDef + @"\d{3}" + "＃" + docid + bookInfoDef + @"\d{3}$"))
                                {
                                    setid = upperClassID + Regex.Replace(bm.Name, @"^.*?(＃.*?)$", "$1");

                                    // 行末尾にブックマークを追加する
                                    sel.Bookmarks.Add(setid);
                                    break;
                                }
                            }

                            if (setid == "")
                            {
                                //while (bookInfoDic.ContainsKey(docid + bookInfoDef + splitCount.ToString("000")))
                                //while (ls.Contains(splitCount.ToString("000")))
                                //{
                                //    splitCount++;
                                //}
                                bibMaxNum++;
                                splitCount = bibMaxNum;
                                ls.Add(splitCount.ToString("000"));
                                setid = upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000");
                                // 行末尾にブックマークを追加する
                                sel.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));

                                //tgtPara.Range.Bookmarks.Add(upperClassID + "♯" + docid + bookInfoDef + splitCount.ToString("000"));
                                bookInfoDic.Add(setid, Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "") + "♪" + tgtPara.Range.Text.Trim());
                                //splitCount++;
                                if (isMerge)
                                {
                                    mergeSetId.Add(setid, previousSetId);
                                }
                                previousSetId = setid;
                            }
                            else if (!bookInfoDic.ContainsKey(setid))
                            {
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

                if (checkBL || oldInfo.Count == 0)
                {
                    using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
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
                                secText[1] = bookInfoDic[key];
                            HeadingInfo headingInfo = new HeadingInfo();
                            if (string.IsNullOrEmpty(secText[0]))
                            {
                                headingInfo.num = "";
                            }
                            else
                            {
                                headingInfo.num = secText[0];
                            }
                            if (string.IsNullOrEmpty(secText[1]))
                            {
                                headingInfo.title = "";
                            }
                            else
                            {
                                headingInfo.title = secText[1];
                            }
                            headingInfo.id = key.Replace("♯", "#");

                            if (mergeSetId.ContainsKey(headingInfo.id))
                            {
                                headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                                makeHeaderLine(docinfo, mergeSetId, headingInfo.num, headingInfo.title, headingInfo.id);
                            }
                            else
                            {
                                docinfo.WriteLine(secText[0] + "\t" + secText[1] + "\t" + key.Replace("♯", "#") + "\t");
                            }
                        }
                    }

                    thisDocument.Save();

                    log.WriteLine("書誌情報リスト作成終了");
                }
                else
                {
                    // 書誌情報（新）
                    foreach (string key in bookInfoDic.Keys)
                    {

                        string[] secText = new string[2];
                        if (bookInfoDic[key].Contains("♪"))
                        {
                            secText[0] = Regex.Replace(bookInfoDic[key], "^(.*?)♪.*?$", "$1");
                            secText[1] = Regex.Replace(bookInfoDic[key], "^.*?♪(.*?)$", "$1");
                        }
                        else
                            secText[1] = bookInfoDic[key];

                        HeadingInfo headingInfo = new HeadingInfo();
                        if (string.IsNullOrEmpty(secText[0]))
                        {
                            headingInfo.num = "";
                        }
                        else
                        {
                            headingInfo.num = secText[0];
                        }
                        if (string.IsNullOrEmpty(secText[1]))
                        {
                            headingInfo.title = "";
                        }
                        else
                        {
                            headingInfo.title = secText[1];
                        }
                        if (key.Contains("＃"))
                        {
                            headingInfo.id = key.Replace("＃", "#");
                        }
                        else
                        {
                            headingInfo.id = key.Replace("♯", "#");

                        }

                        if (mergeSetId.ContainsKey(headingInfo.id))
                        {
                            headingInfo.mergeto = mergeSetId[headingInfo.id].Split(new char[] { '♯', '#' })[0];
                        }

                        newInfo.Add(headingInfo);
                    }

                    // 新旧比較処理
                    int ret = checkDocInfo(oldInfo, newInfo, out checkResult);

                    // 処理結果が0:正常の場合
                    if (ret == 0)
                    {
                        using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                        {
                            foreach (HeadingInfo info in newInfo)
                            {
                                makeHeaderLine(docinfo, mergeSetId, info.num, info.title, info.id);
                                //docinfo.WriteLine(info.num + "\t" + info.title + "\t" + info.id + "\t" + (mergeSetId.ContainsKey(info.id) ? mergeSetId[info.id]:""));
                            }
                        }

                        thisDocument.Save();

                        log.WriteLine("書誌情報リスト作成終了");
                    }
                    else if (ret == 1)
                    {
                        // 処理結果が1:異常の場合
                        // 書誌情報比較チェック画面を表示する
                        load.Visible = false;
                        CheckForm checkForm = new CheckForm(this);
                        DialogResult returnCode = checkForm.ShowDialog();

                        if (returnCode != DialogResult.OK)
                        {

                            if (swLog == null)
                            {
                                log.Close();
                            }

                            return false;
                        }
                        else
                        {
                            if (blHTMLPublish)
                                load.Visible = true;
                            // 新.IDをドキュメントに反映する
                            foreach (Word.Bookmark wb in thisDocument.Bookmarks) wb.Delete();

                            foreach (Word.Section tgtSect in thisDocument.Sections)
                            {
                                foreach (Word.Paragraph tgtPara in tgtSect.Range.Paragraphs)
                                {
                                    string styleName = tgtPara.get_Style().NameLocal;

                                    if (!Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") && !styleName.Contains("見出し")) continue;

                                    string innerText = tgtPara.Range.Text.Trim();

                                    if (tgtPara.Range.Text.Trim() == "") continue;

                                    if (Regex.IsMatch(innerText, @"^[\s　]*索[\s　]*引[\s　]*$") && (Regex.IsMatch(styleName, "章[　 ]*扉.*タイトル") || Regex.IsMatch(styleName, @"(見出し|Heading)\s*[１1](?:[^・用]+|)$")))
                                    {
                                        breakFlg = true;
                                        break;
                                    }

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

                                        string num = Regex.Replace(tgtPara.Range.ListFormat.ListString, @"[^\.\d]", "");
                                        string title = tgtPara.Range.Text.Trim();

                                        CheckInfo info = checkResult.Where(p => ((string.IsNullOrEmpty(p.new_num) && string.IsNullOrEmpty(num)) || p.new_num.Equals(num))
                                            && p.new_title.Equals(title)).FirstOrDefault();

                                        if (info != null)
                                        {
                                            // 行末尾にブックマークを追加する
                                            sel.Bookmarks.Add(info.new_id_show.Split(new char[] { '(' })[0].Trim().Replace("#", "♯"));
                                        }
                                    }
                                }

                                if (breakFlg) break;
                            }

                            using (StreamWriter docinfo = new StreamWriter(rootPath + "\\" + headerDir + "\\" + docid + ".txt", false, Encoding.UTF8))
                            {
                                foreach (CheckInfo info in checkResult)
                                {
                                    if (string.IsNullOrEmpty(info.new_id))
                                    {
                                        continue;
                                    }
                                    makeHeaderLine(docinfo, mergeSetId, info.new_num, info.new_title, info.new_id_show.Split(new char[] { '(' })[0].Trim());
                                    //docinfo.WriteLine(info.new_num + "\t" + info.new_title + "\t" + info.new_id_show + "\t" + (mergeSetId.ContainsKey(info.new_id_show) ? mergeSetId[info.new_id_show] : ""));
                                }
                            }

                            thisDocument.Save();

                            log.WriteLine("書誌情報リスト作成終了");
                        }
                    }
                }

                // SOURCELINK変更==========================================================================END

                if (swLog == null)
                {
                    log.Close();
                    File.Delete(rootPath + "\\log.txt");
                }
                blHTMLPublish = false;
                return true;

            }
            catch (Exception ex)
            {
                StackTrace stackTrace = new StackTrace(ex, true);

                log.WriteLine(ex.Message);
                log.WriteLine(ex.HelpLink);
                log.WriteLine(ex.Source);
                log.WriteLine(ex.StackTrace);
                log.WriteLine(ex.TargetSite);

                if (swLog == null)
                {
                    log.Close();
                }
                load.Visible = false;
                MessageBox.Show("エラーが発生しました");

                button4.Enabled = true;
                blHTMLPublish = false;
                return false;
            }
            finally
            {
                WordAddIn1.Globals.ThisAddIn.Application.Selection.HomeKey(Word.WdUnits.wdStory);
                Application.DoEvents();
                WordAddIn1.Globals.ThisAddIn.Application.ScreenUpdating = true;
            }

            //WordAddIn1.Globals.ThisAddIn.Application.Selection.Start = selStart;
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.End = selEnd;
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveRight(Unit: Word.WdUnits.wdCharacter, Count: 1);
            //WordAddIn1.Globals.ThisAddIn.Application.Selection.MoveLeft(Unit: Word.WdUnits.wdCharacter, Count: 1);
        }


        // SOURCELINK追加==========================================================================START
        /// <summary>
        /// 新規比較処理
        /// </summary>
        /// <param name="oldInfos">書誌情報（旧）</param>
        /// <param name="newInfos">書誌情報（新）</param>
        /// <param name="checkResult">比較結果リスト</param>
        /// <returns>処理結果</returns>
        private int checkDocInfo(List<HeadingInfo> oldInfos, List<HeadingInfo> newInfos, out List<CheckInfo> checkResult)
        {
            // 比較結果リスト初期化する
            checkResult = new List<CheckInfo>();
            List<CheckInfo> syoriList = new List<CheckInfo>();
            List<CheckInfo> deleteList = new List<CheckInfo>();
            int returnCode = 0;

            // 一致判定と削除判定
            foreach (HeadingInfo oldInfo in oldInfos)
            {
                bool oldTitleExist = false;
                bool oldIdExist = false;

                foreach (HeadingInfo newInfo in newInfos)
                {
                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトルかつ書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.title.Equals(newInfo.title) && oldInfo.id.Equals(newInfo.id))
                    {
                        // 比較結果（一致）を作成する
                        CheckInfo checkInfo = new CheckInfo();
                        // 旧.項番
                        checkInfo.old_num = oldInfo.num;
                        // 旧.タイトル
                        checkInfo.old_title = oldInfo.title;
                        // 旧.ID
                        checkInfo.old_id = oldInfo.id;
                        // 旧.ID結合済
                        if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                        // 新.項番
                        checkInfo.new_num = newInfo.num;
                        // 新.タイトル
                        checkInfo.new_title = newInfo.title;
                        // 新.ID
                        checkInfo.new_id = newInfo.id;
                        // 新.ID結合済
                        if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                        // 新.ID（修正候補）
                        checkInfo.new_id_show = newInfo.id;
                        // 新.ID（修正候補）結合済
                        if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                        // check merge 
                        if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                        {
                            checkInfo.diff = "結合追加";
                            checkInfo.new_id_color = "red";
                            returnCode = 1;
                        }
                        else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                        {
                            checkInfo.diff = "結合解除";
                            checkInfo.new_id_color = "red";
                            returnCode = 1;
                        }

                        syoriList.Add(checkInfo);
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        oldTitleExist = true;
                    }

                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        oldIdExist = true;
                    }
                }

                // 書誌情報（旧）.タイトルと書誌情報（旧）.IDが書誌情報（新）に存在しない場合
                if (!oldTitleExist && !oldIdExist)
                {
                    // 比較結果（削除）を作成する
                    CheckInfo checkInfo = new CheckInfo();
                    // 旧.項番
                    checkInfo.old_num = oldInfo.num;
                    // 旧.タイトル
                    checkInfo.old_title = oldInfo.title;
                    // 旧.ID
                    checkInfo.old_id = oldInfo.id;
                    // 旧.ID結合済
                    if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                    // 差異内容
                    checkInfo.diff = "削除";

                    deleteList.Add(checkInfo);
                }
            }

            // 新規判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                bool newTitleExist = false;
                bool newIdExist = false;

                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        newIdExist = true;
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        newTitleExist = true;
                    }
                }

                // 書誌情報（新）.タイトルと書誌情報（新）.IDが書誌情報（旧）に存在しない場合
                if (!newTitleExist && !newIdExist)
                {
                    // 比較結果（新規）を作成する
                    CheckInfo checkInfo = new CheckInfo();
                    // 新.項番
                    checkInfo.new_num = newInfo.num;
                    // 新.項番（色）
                    checkInfo.new_num_color = "blue";
                    // 新.タイトル
                    checkInfo.new_title = newInfo.title;
                    // 新.タイトル（色）
                    checkInfo.new_title_color = "blue";
                    // 新.ID
                    checkInfo.new_id = newInfo.id;
                    // 新.ID結合済
                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                    // 新.ID（修正候補）
                    checkInfo.new_id_show = newInfo.id;
                    // 新.ID（修正候補）結合済
                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                    // 新.ID（色）
                    checkInfo.new_id_color = "blue";

                    // 差異内容
                    checkInfo.diff = "新規追加";

                    // ＋結合追加
                    if (!newInfo.mergeto.Equals(""))
                    {
                        checkInfo.diff = "新規追加・結合追加";

                    }

                    syoriList.Add(checkInfo);
                }
            }

            // ID不一致判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // リストに存在するか
                    CheckInfo hasOne = syoriList.Where(p => p.new_id.Equals(newInfo.id)).FirstOrDefault();
                    if (hasOne != null)
                    {
                        break;
                    }

                    // 書誌情報（新）.タイトル＝書誌情報（旧）.タイトル
                    if (oldInfo.title.Equals(newInfo.title))
                    {
                        // 書誌情報（新）.ID<>書誌情報（旧）.ID
                        if (!oldInfo.id.Equals(newInfo.id))
                        {
                            // 項番階層
                            string oldNum = oldInfo.num;
                            string newNum = newInfo.num;
                            int oldNumKaisou = oldNum.Split('.').Length;
                            int newNumKaisou = newNum.Split('.').Length;

                            // (旧.見出しレベルが3 階層かつ新.見出しレベルが４階層) 
                            // または　(旧.見出しレベルが4 階層かつ新.見出しレベルが3階層) )の場合
                            if ((oldNumKaisou == 3 && newNumKaisou == 4)
                                || (oldNumKaisou == 4 && newNumKaisou == 3))
                            {
                                // 比較結果（見出しレベル変更）を作成する
                                CheckInfo checkInfo = new CheckInfo();
                                // 旧.項番
                                checkInfo.old_num = oldInfo.num;
                                // 旧.タイトル
                                checkInfo.old_title = oldInfo.title;
                                // 旧.ID
                                checkInfo.old_id = oldInfo.id;
                                // 旧.ID結合済
                                if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                // 新.項番
                                checkInfo.new_num = newInfo.num;
                                // 新.項番（色）
                                checkInfo.new_num_color = "red";
                                // 新.タイトル
                                checkInfo.new_title = newInfo.title;
                                // 新.ID
                                checkInfo.new_id = newInfo.id;
                                // 新.ID結合済
                                if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                // 新.ID（修正候補）
                                checkInfo.new_id_show = newInfo.id;
                                // 新.ID（修正候補）結合済
                                if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                // 新.ID（色）
                                checkInfo.new_id_color = "red";
                                // 差異内容
                                checkInfo.diff = "見出しレベル変更";

                                syoriList.Add(checkInfo);
                            }
                            else
                            {
                                // 構成変更に伴うID変更
                                bool isHenko = false;
                                if (oldNumKaisou == 4 && newNumKaisou == 4)
                                {
                                    string[] oldids = oldInfo.id.Split('#');
                                    string[] newids = newInfo.id.Split('#');

                                    if (oldids.Length == 2 && newids.Length == 2
                                        && oldids[1].Equals(newids[1]))
                                    {

                                        // 比較結果（構成変更に伴うID変更）を作成する
                                        CheckInfo checkInfo2 = new CheckInfo();
                                        // 旧.項番
                                        checkInfo2.old_num = oldInfo.num;
                                        // 旧.タイトル
                                        checkInfo2.old_title = oldInfo.title;
                                        // 旧.ID
                                        checkInfo2.old_id = oldInfo.id;
                                        // 旧.ID結合済
                                        if (!oldInfo.mergeto.Equals("")) { checkInfo2.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                        // 新.項番
                                        checkInfo2.new_num = newInfo.num;
                                        // 新.項番（色）
                                        checkInfo2.new_num_color = "red";
                                        // 新.タイトル
                                        checkInfo2.new_title = newInfo.title;
                                        // 新.ID
                                        checkInfo2.new_id = newInfo.id;
                                        // 新.ID結合済
                                        if (!newInfo.mergeto.Equals("")) { checkInfo2.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                        // 新.ID（修正候補）
                                        checkInfo2.new_id_show = newInfo.id;
                                        // 新.ID（修正候補）結合済
                                        if (!newInfo.mergeto.Equals("")) { checkInfo2.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                                        // 新.ID（色）
                                        checkInfo2.new_id_color = "red";
                                        // 差異内容
                                        checkInfo2.diff = "構成変更に伴うID変更";

                                        syoriList.Add(checkInfo2);

                                        isHenko = true;
                                    }

                                }

                                if (!isHenko)
                                {
                                    // 比較結果（ID不一致）を作成する
                                    CheckInfo checkInfo = new CheckInfo();
                                    // 旧.項番
                                    checkInfo.old_num = oldInfo.num;
                                    // 旧.タイトル
                                    checkInfo.old_title = oldInfo.title;
                                    // 旧.ID
                                    checkInfo.old_id = oldInfo.id;
                                    // 旧.ID結合済
                                    if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                                    // 新.項番
                                    checkInfo.new_num = newInfo.num;
                                    // 新.項番（色）
                                    // 旧.項番<>新.項番の場合、赤
                                    if (!oldInfo.num.Equals(newInfo.num))
                                    {
                                        checkInfo.new_num_color = "red";
                                    }
                                    // 新.タイトル
                                    checkInfo.new_title = newInfo.title;
                                    // 新.ID
                                    checkInfo.new_id = newInfo.id;
                                    // 新.ID結合済
                                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                    // 新.ID（色）
                                    checkInfo.new_id_color = "red";
                                    // 新.ID（修正候補）
                                    checkInfo.new_id_show = oldInfo.id;
                                    // 新/ID（修正候補）結合済
                                    if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }
                                    // 差異内容
                                    checkInfo.diff = "ID不一致？";
                                    // 差異内容（色）
                                    checkInfo.diff_color = "red";

                                    // 修正処理（候補）
                                    checkInfo.editshow = "旧IDに戻す";

                                    // check merge 
                                    if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                                    {
                                        checkInfo.diff = "ID不一致？・結合追加";
                                    }
                                    else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                                    {
                                        checkInfo.diff = "ID不一致？・結合解除";
                                    }

                                    syoriList.Add(checkInfo);

                                    returnCode = 1;
                                }
                            }
                        }
                    }
                }
            }

            // タイトル変更判定
            foreach (HeadingInfo newInfo in newInfos)
            {
                // リストに存在するか
                CheckInfo hasOne = syoriList.Where(p => p.new_id.Equals(newInfo.id)).FirstOrDefault();
                if (hasOne != null)
                {
                    continue;
                }

                foreach (HeadingInfo oldInfo in oldInfos)
                {
                    // 書誌情報（新）.ID＝書誌情報（旧）.IDが存在する場合
                    if (oldInfo.id.Equals(newInfo.id))
                    {
                        // 書誌情報（新）.タイトル<>書誌情報（旧）.タイトル
                        if (!oldInfo.title.Equals(newInfo.title))
                        {
                            // 比較結果（タイトル変更）を作成する
                            CheckInfo checkInfo = new CheckInfo();
                            // 旧.項番
                            checkInfo.old_num = oldInfo.num;
                            // 旧.タイトル
                            checkInfo.old_title = oldInfo.title;
                            // 旧.ID
                            checkInfo.old_id = oldInfo.id;
                            // 旧・ID結合済
                            if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                            // 新.項番
                            checkInfo.new_num = newInfo.num;

                            // 新.項番（色）
                            // 旧.項番<>新.項番の場合、赤
                            if (!oldInfo.num.Equals(newInfo.num))
                            {
                                checkInfo.new_num_color = "red";
                            }

                            // 新.タイトル
                            checkInfo.new_title = newInfo.title;
                            // 新.タイトル（色）
                            checkInfo.new_title_color = "red";
                            // 新.ID
                            checkInfo.new_id = newInfo.id;
                            // 新.ID結合済
                            if (!newInfo.mergeto.Equals("")) { checkInfo.new_id = newInfo.id + " (" + newInfo.mergeto + ")"; }
                            // 新.ID（修正候補）
                            checkInfo.new_id_show = newInfo.id;
                            // 新.ID（修正候補）結合済
                            if (!newInfo.mergeto.Equals("")) { checkInfo.new_id_show = newInfo.id + " (" + newInfo.mergeto + ")"; }

                            // 差異内容
                            checkInfo.diff = "●タイトル変更";

                            // 新規追加
                            checkInfo.edit = "○新規追加";

                            // 新規追加（色）
                            checkInfo.edit_color = "blue";

                            // check merge 
                            if (oldInfo.mergeto.Equals("") && !newInfo.mergeto.Equals(""))
                            {
                                checkInfo.diff = "●タイトル変更・結合追加";
                                checkInfo.new_id_color = "red";
                            }
                            else if (!oldInfo.mergeto.Equals("") && newInfo.mergeto.Equals(""))
                            {
                                checkInfo.diff = "●タイトル変更・結合解除";
                                checkInfo.new_id_color = "red";
                            }

                            syoriList.Add(checkInfo);

                            returnCode = 1;
                        }
                    }
                }
            }

            // 削除再判定
            foreach (HeadingInfo oldInfo in oldInfos)
            {
                var issyori = syoriList.Where(p => p.old_num.Equals(oldInfo.num)).ToList();
                if (issyori != null && issyori.Count > 0)
                {
                    continue;
                }

                var isdelete = deleteList.Where(p => p.old_num.Equals(oldInfo.num)).ToList();
                if (isdelete != null && isdelete.Count > 0)
                {
                    continue;
                }

                // 比較結果（削除）を作成する
                CheckInfo checkInfo = new CheckInfo();
                // 旧.項番
                checkInfo.old_num = oldInfo.num;
                // 旧.タイトル
                checkInfo.old_title = oldInfo.title;
                // 旧.ID
                checkInfo.old_id = oldInfo.id;
                // 旧・ID結合済
                if (!oldInfo.mergeto.Equals("")) { checkInfo.old_id = oldInfo.id + " " + oldInfo.mergeto + ""; }
                // 差異内容
                checkInfo.diff = "削除";

                deleteList.Add(checkInfo);
            }

            // ソート
            deleteList = deleteList.OrderBy(rec => rec.old1).ThenBy(rec =>
            rec.old2).ThenBy(rec => rec.old3).ThenBy(rec => rec.old4).ToList();

            // ソート
            syoriList = syoriList.OrderBy(rec => rec.new1).ThenBy(rec =>
                rec.new2).ThenBy(rec => rec.new3).ThenBy(rec => rec.new4).ToList();

            if (deleteList.Count > 0)
            {
                int i = 0;
                bool stopFlag = false;

                for (int j = 0; j < syoriList.Count; j++)
                {
                    while (!stopFlag && checkSortInfo(deleteList[i], syoriList, j))
                    {
                        checkResult.Add(deleteList[i]);
                        i++;

                        if (deleteList.Count == i)
                        {
                            stopFlag = true;
                        }
                    }

                    checkResult.Add(syoriList[j]);
                }

                while (i < deleteList.Count)
                {
                    checkResult.Add(deleteList[i]);
                    i++;
                }
            }
            else
            {
                checkResult = syoriList;
            }
            if (newInfos.Count == oldInfos.Count)
            {
                foreach (HeadingInfo newInfo in newInfos)
                {
                    var checkHeadingInfo = oldInfos.Where(x => x.id == newInfo.id && x.num == newInfo.num && x.mergeto == newInfo.mergeto && x.title == newInfo.title);
                    if (checkHeadingInfo == null)
                    {
                        returnCode = 1;
                        break;
                    }
                }

            }
            else
            {
                returnCode = 1;
            }


            return returnCode;
        }


        private bool checkSortInfo(CheckInfo old, List<CheckInfo> newInfos, int j)
        {
            bool ret = false;

            CheckInfo newInfo = newInfos[j];

            if (old.old1 < newInfo.old1)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 < newInfo.old2)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 == newInfo.old2 && old.old3 < newInfo.old3)
            {
                ret = true;
            }
            else if (old.old1 == newInfo.old1 && old.old2 == newInfo.old2 && old.old3 == newInfo.old3 && old.old4 < newInfo.old4)
            {
                ret = true;
            }

            for (int k = j + 1; k < newInfos.Count; k++)
            {
                CheckInfo newInfoK = newInfos[k];

                if (string.IsNullOrEmpty(newInfoK.old_id))
                {
                    continue;
                }

                if (old.old1 > newInfoK.old1)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 > newInfoK.old2)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 == newInfoK.old2 && old.old3 > newInfoK.old3)
                {
                    ret = false;
                }
                else if (old.old1 == newInfoK.old1 && old.old2 == newInfoK.old2 && old.old3 == newInfoK.old3 && old.old4 > newInfoK.old4)
                {
                    ret = false;
                }
            }

            return ret;
        }

        private void button11_Click(object sender, RibbonControlEventArgs e)
        {

        }
        // SOURCELINK追加==========================================================================END
    }
}
