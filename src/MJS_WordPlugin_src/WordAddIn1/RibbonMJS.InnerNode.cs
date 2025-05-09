using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
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
    }
}
