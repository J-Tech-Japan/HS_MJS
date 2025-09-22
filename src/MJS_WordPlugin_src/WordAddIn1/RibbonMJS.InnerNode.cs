// RibbonMJS.InnerNode.cs

using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void InnerNode(Dictionary<string, string> styleName, XmlNode objTargetNode, XmlNode seekNode)
        {

            if (seekNode.NodeType == XmlNodeType.Text)
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, true));
            }
            else if (seekNode.NodeType == XmlNodeType.Element)
            {
                string thisStyleName = getStyleName(styleName, seekNode);

                if (seekNode.Name == "a")
                {
                    string refname = ((XmlElement)seekNode).GetAttribute("name");
                    if (refname.Contains("_Ref"))
                    {
                        objTargetNode.AppendChild(objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("span")));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("name", refname);
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "ref");
                    }
                }

                if ((seekNode.Name == "table") || (seekNode.Name == "tr") || (seekNode.Name == "td"))
                {
                    AppendTableElement(styleName, objTargetNode, seekNode, thisStyleName);
                }

                else if (Regex.IsMatch(seekNode.Name, @"^(?:[bui]|ins|h\d+)$", RegexOptions.IgnoreCase) ||
                        ((objTargetNode.SelectSingleNode("ancestor-or-self::*[starts-with(@class, 'Heading')]") != null) && (seekNode.Name == "p")) ||
                       (seekNode.SelectNodes("@class[. = 'msoIns']").Count != 0))
                {
                    foreach (XmlNode Child in seekNode.ChildNodes)
                    {
                        InnerNode(styleName, objTargetNode, Child);
                    }
                    return;
                }
                else
                {
                    if (Regex.IsMatch(thisStyleName, "表見出し"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading_table");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'Heading_table']").Count == 0))
                        {
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading_table");
                        }
                    }

                    else if (Regex.IsMatch(thisStyleName, "MJS_画像（操作の流れ）"))
                    {
                        AppendMJSOperationFlowElement(objTargetNode, seekNode, "MJS_oflowPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順番号"))
                    {
                        AppendMJSOperationFlowElement(objTargetNode, seekNode, "MJS_oflow_stepNum");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順結果"))
                    {
                        AppendMJSOperationFlowElement(objTargetNode, seekNode, "MJS_oflow_stepResult");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_手順"))
                    {
                        AppendMJSOperationFlowElement(objTargetNode, seekNode, "MJS_oflow_step");
                    }
                    else if (Regex.IsMatch(thisStyleName, "MJS_操作の流れ_補足"))
                    {
                        AppendMJSOperationFlowElement(objTargetNode, seekNode, "MJS_oflow_sub");
                    }

                    else if (seekNode.SelectNodes("@class[. = 'MJS_qa_td_Qblock']").Count != 0)
                    {

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_td_Qblock");

                    }
                    else if (seekNode.SelectNodes("@class[. = 'MJS_qa_td_Ablock']").Count != 0)
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_td_Ablock");

                    }
                    else if (seekNode.SelectNodes("@class[. = 'Q＆A']").Count != 0)
                    {

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "Q＆A");

                    }

                    else if (Regex.IsMatch(thisStyleName, "Q＆A"))
                    {
                        AppendQandAElement(styleName, objTargetNode, seekNode, thisStyleName);
                    }

                    else if (Regex.IsMatch(thisStyleName, "リード文"))
                    {
                        AppendLeadSentenceElement(objTargetNode, seekNode, thisStyleName);
                    }


                    else if (Regex.IsMatch(thisStyleName, "下線"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_ul");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "処理フロー"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (seekNode.SelectNodes("ancestor::*[@class = 'MJS_flow']").Count == 0)
                        {
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_flow");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "参照先"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_ref");
                    }

                    else if (
                        Regex.IsMatch(thisStyleName, "選択肢等[2２]") ||
                        Regex.IsMatch(thisStyleName, "選択肢等") ||
                        Regex.IsMatch(thisStyleName, "選択肢-説明等[2２]") ||
                        Regex.IsMatch(thisStyleName, "選択肢.*説明等"))
                    {
                        AppendChoiceElement(objTargetNode, seekNode, thisStyleName);
                    }

                    else if (Regex.IsMatch(thisStyleName, "箇条書き[2２]"))
                    {
                        AppendListItemElement(objTargetNode, seekNode, "MJS_listItem2");
                    }
                    else if (Regex.IsMatch(thisStyleName, "箇条書き"))
                    {
                        AppendListItemElement(objTargetNode, seekNode, "MJS_listItem");
                    }

                    else if (Regex.IsMatch(thisStyleName, "表内-項目_センタリング"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem_center");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内-項目_右寄せ"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem_right");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*タイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*箇条"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                        {
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_tdListItem']").Count == 0))
                        {
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdListItem");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*項目"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdItem");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表内.*本文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_tdText']").Count == 0))
                        {
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdText");
                        }
                        else
                        {
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "(?:コラム.*アイコン|事項.*アイコン用?)"))
                    {
                        if (seekNode.SelectNodes(".//img").Count != 0)
                        {
                            foreach (XmlNode Icon in seekNode.SelectNodes(".//img"))
                            {
                                ((XmlElement)Icon).SetAttribute("width", "80");
                            }
                        }

                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_icons");
                    }
                    else if (Regex.IsMatch(thisStyleName, "コラム.*本文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                        if (seekNode.SelectNodes("ancestor::*[@class = 'MJS_columnText']").Count == 0)
                        {
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnText");
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "コラム.*タイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "見出し.*手順"))
                    {
                        if (seekNode.SelectNodes(".//img").Count != 0)
                        {
                            ((XmlElement)seekNode.SelectSingleNode(".//img")).SetAttribute("width", "35");
                            if (seekNode.SelectSingleNode(".//img").NextSibling.InnerText == "　")
                            {
                                seekNode.SelectSingleNode(".//img").NextSibling.InnerText = ((char)160).ToString();
                            }
                        }
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processTitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*本文内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_bodyPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*表内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tdPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*コラム内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_columnPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "画像.*手順内"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_flowPict");
                    }
                    else if (Regex.IsMatch(thisStyleName, "メニュー[2２]"))
                    {
                        if (seekNode.Name == "p")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_Menu2");
                        }
                        else
                        {
                            foreach (XmlNode Child in seekNode.ChildNodes)
                            {
                                InnerNode(styleName, objTargetNode, Child);
                            }
                            return;
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "メニュー"))
                    {
                        if (seekNode.Name == "p")
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_menu");
                        }
                        else
                        {
                            foreach (XmlNode Child in seekNode.ChildNodes)
                            {
                                InnerNode(styleName, objTargetNode, Child);
                            }
                            return;
                        }
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順結果"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processResult");
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順番号リセット用"))
                    {
                        return;
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順文"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_process");
                    }
                    else if (Regex.IsMatch(thisStyleName, "手順補足"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_processSuppl");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルタイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_title");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルサブタイトル"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_subtitle");
                    }
                    else if (Regex.IsMatch(thisStyleName, "マニュアルバージョン"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "manual_version");
                    }
                    else if ((objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_process') or starts-with(@class, 'MJS_process')]").Count != 0) &&
                             (seekNode.SelectNodes("@style[contains(., 'color:#1F497D')]").Count != 0))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        string bold = "";
                        if (seekNode.SelectNodes("ancestor-or-self::b").Count != 0) bold = "font-weight:bold;";
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#1F497D;" + bold);
                    }

                    else if (Regex.IsMatch(thisStyleName, "ui"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if ((seekNode.Name == "p") && (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "") && (seekNode.SelectNodes(".//img").Count != 0))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "margin-left:2em;");
                    }
                    else if ((seekNode.Name == "span") && (seekNode.ChildNodes.Count == 0))
                    {
                        return;
                    }
                    else if (seekNode.SelectNodes("@class[. = 'MsoHyperlink']").Count != 0)
                    {
                        foreach (XmlNode Child in seekNode.ChildNodes)
                        {
                            InnerNode(styleName, objTargetNode, Child);
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
                                Regex.IsMatch(((XmlElement)seekNode).GetAttribute("style"), @"(?<![A-z\d-])color:"))
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            string thisColor = Regex.Replace(((XmlElement)objTargetNode.LastChild).GetAttribute("style"), @"^.*(?<![A-z\d-])(color:[^;]+;).+$", "$1");
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "font-family:Wingdings;" + thisColor);
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "font-family:Wingdings");
                        }
                    }
                    else if (seekNode.SelectNodes("@style[contains(translate(., ' ', ''), 'color:#246A98;font-weight:normal')]").Count != 0)
                    {
                        if (objTargetNode.SelectNodes("ancestor-or-self::*[(@class = 'MJS_qa_Q')]").Count != 0)
                        {
                            return;
                        }
                        else
                        {
                            objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#246A98;font-weight:normal");
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
                            ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                            ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "color:#8EAADB");
                        }
                    }

                    else if (Regex.IsMatch(thisStyleName, "タブ見出し"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "Heading4_tab");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }
                    else if (Regex.IsMatch(thisStyleName, "表脚注"))
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_tableFootnote");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                    }

                    else
                    {
                        objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                        ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                    }
                }

                foreach (XmlNode Child in seekNode.ChildNodes)
                {
                    InnerNode(styleName, objTargetNode.LastChild, Child);
                }
            }
        }

    }
}