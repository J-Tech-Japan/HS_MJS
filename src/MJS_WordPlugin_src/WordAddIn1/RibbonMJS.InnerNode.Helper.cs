using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void AppendTableElement(
            Dictionary<string, string> styleName,
            XmlNode objTargetNode,
            XmlNode seekNode,
            string thisStyleName)
        {
            if (Regex.IsMatch(thisStyleName, "参照先"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                string baseStyle = "";
                if (objTargetNode.LastChild.SelectNodes("@style").Count != 0)
                {
                    baseStyle = ((XmlElement)objTargetNode.LastChild).GetAttribute("style");
                }
                ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "text-align:right; font-size:90%;" + baseStyle);
            }
            else if (seekNode.Name == "table")
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                string thisStyle = ((XmlElement)objTargetNode.LastChild).GetAttribute("style");
                if (Regex.IsMatch(seekNode.InnerXml, @"border-top:double[\s　]+#[^\s|　]+?[\s　]+4.5pt") && Regex.IsMatch(seekNode.InnerXml, @"border-bottom:double[\s　]+#[^\s|　]+?[\s　]+4.5pt"))
                {
                    if (Regex.IsMatch(thisStyle, @"(?<![A-z\d-])width:"))
                    {
                        ((XmlElement)objTargetNode.LastChild).SetAttribute("style", Regex.Replace(thisStyle, @"(?<![A-z\d-])width:[^;]+;", ""));
                    }
                }
                else if (Regex.IsMatch(thisStyle, @"(?<![A-z\d-])width:"))
                {
                    ((XmlElement)objTargetNode.LastChild).SetAttribute("style", Regex.Replace(thisStyle, @"(?<![A-z\d-])(?<=width:)[\d\.]+",
                    Convert.ToString(Math.Round(decimal.Parse(Regex.Replace(thisStyle, @"^.*?width:([\d\.]+)\w+;.*?$", "$1")) * 1.15m, 1, MidpointRounding.AwayFromZero))));
                }
            }
            else if (seekNode.Name == "td")
            {
                foreach (XmlNode childs in seekNode.SelectNodes(".//*[boolean(@class)]"))
                {
                    if (Regex.IsMatch(getStyleName(styleName, childs), "コラム.*アイコン"))
                    {
                        ((XmlElement)seekNode).SetAttribute("width", "80");
                        break;
                    }
                    else if (Regex.IsMatch(getStyleName(styleName, childs), "コラム"))
                    {
                        ((XmlElement)seekNode).RemoveAttribute("width");
                        if (seekNode.SelectNodes("@style").Count != 0)
                        {
                            ((XmlElement)seekNode).SetAttribute("style", Regex.Replace(((XmlElement)seekNode).GetAttribute("style"), "(?:^| )width:[^;]+;", ""));
                        }
                        break;
                    }
                    else if (childs.Name == "table")
                    {
                        ReplaceTableNodeWithQADiv(childs);
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

        // 「MJS_操作の流れ」関連の要素（手順番号・手順・手順結果・補足・画像など）を表すXMLノードを、
        // 指定したclass属性を付与してターゲットノードに追加
        private void AppendMJSOperationFlowElement(XmlNode targetNode, XmlNode sourceNode, string className)
        {
            var imported = targetNode.OwnerDocument.ImportNode(sourceNode, false);
            targetNode.AppendChild(imported);
            var elem = imported as XmlElement;
            if (elem != null)
            {
                elem.RemoveAttribute("style");
                elem.SetAttribute("class", className);
            }
        }

        // Q＆A形式のtableノードをdivノードに変換し、親ノードで置換する
        private void ReplaceTableNodeWithQADiv(XmlNode tableNode)
        {
            XmlNode divNode = tableNode.OwnerDocument.CreateElement("div");
            divNode.Attributes.Append(tableNode.OwnerDocument.CreateAttribute("class")).Value = "Q＆A";

            foreach (XmlNode trNode in tableNode.SelectNodes(".//tr"))
            {
                if (trNode.SelectNodes(".//p[@class='MJS-QA']").Count != 0)
                {
                    XmlNode qBlockDivNode = divNode.OwnerDocument.CreateElement("div");
                    foreach (XmlNode childNode in trNode.SelectNodes(".//td"))
                    {
                        XmlNodeList pNodes = childNode.SelectNodes(".//p");
                        foreach (XmlNode pNode in pNodes)
                        {
                            if (pNode.SelectNodes("@class[. = 'MJS-QA']").Count == 0)
                            {
                                qBlockDivNode.AppendChild(pNode);
                            }
                        }
                        ((XmlElement)qBlockDivNode).SetAttribute("class", "MJS_qa_td_Qblock");
                        divNode.AppendChild(qBlockDivNode);
                    }
                }
                else if (trNode.SelectNodes(".//p[@class='MJSQAA']").Count != 0)
                {
                    XmlNode aBlockDivNode = divNode.OwnerDocument.CreateElement("div");
                    foreach (XmlNode childNode in trNode.SelectNodes(".//td"))
                    {
                        XmlNodeList pNodes = childNode.SelectNodes(".//p");
                        foreach (XmlNode pNode in pNodes)
                        {
                            if (pNode.SelectNodes("@class[. = 'MsoNormal']").Count == 0)
                            {
                                aBlockDivNode.AppendChild(pNode);
                            }
                        }
                        ((XmlElement)aBlockDivNode).SetAttribute("class", "MJS_qa_td_Ablock");
                        divNode.AppendChild(aBlockDivNode);
                    }
                }
            }
            tableNode.ParentNode.ReplaceChild(divNode, tableNode);
        }

        // Q＆A形式の要素をHTML/XMLノードとして適切に追加・変換する
        private void AppendQandAElement(
            Dictionary<string, string> styleName,
            XmlNode objTargetNode,
            XmlNode seekNode,
            string thisStyleName)
        {
            if (thisStyleName == "MJS_Q＆A_Q")
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_Q");
            }
            else if (Regex.IsMatch(thisStyleName, "アイコン"))
            {
                if (seekNode.SelectNodes(".//img").Count != 0)
                {
                    foreach (XmlNode Icon in seekNode.SelectNodes(".//img"))
                    {
                        ((XmlElement)Icon).SetAttribute("width", "80");
                    }
                }
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_icons");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
            else if (thisStyleName == "MJS_Q＆A_A")
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                {
                    ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                }

                if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_qa_A']").Count == 0))
                {
                    ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_A");
                }
            }
            else if (thisStyleName == "MJS_Q＆A_A継続")
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.CreateElement("p"));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                ((XmlElement)objTargetNode.LastChild).SetAttribute("style", "padding-left: 40px;");
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_qa_A_cont");
            }
            else
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
            }
        }


        // 「リード文」スタイルのノードを適切なclassで追加する
        private void AppendLeadSentenceElement(XmlNode objTargetNode, XmlNode seekNode, string thisStyleName)
        {
            if (Regex.IsMatch(thisStyleName, "リード文.*[1１2２]"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_leadSentence1_2");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
            else if (Regex.IsMatch(thisStyleName, "リード文.*[3３4４]"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_leadSentence3_4");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
            else
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_lead");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
        }

        // 「選択肢」関連のノードを適切なclassで追加する
        private void AppendChoiceElement(XmlNode objTargetNode, XmlNode seekNode, string thisStyleName)
        {
            if (Regex.IsMatch(thisStyleName, "選択肢等[2２]"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choice2");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
            else if (Regex.IsMatch(thisStyleName, "選択肢等"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choice");
            }
            else if (Regex.IsMatch(thisStyleName, "選択肢-説明等[2２]"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choiceDesc2");
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");
            }
            else if (Regex.IsMatch(thisStyleName, "選択肢.*説明等"))
            {
                objTargetNode.AppendChild(objTargetNode.OwnerDocument.ImportNode(seekNode, false));
                ((XmlElement)objTargetNode.LastChild).RemoveAttribute("style");

                if (Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") == "")
                {
                    ((XmlElement)objTargetNode.LastChild).RemoveAttribute("class");
                }

                if ((Regex.Replace(seekNode.InnerText, @"[\s　\u00A0]", "") != "") && (objTargetNode.SelectNodes("ancestor-or-self::*[@class = 'MJS_choiceDesc']").Count == 0))
                {
                    ((XmlElement)objTargetNode.LastChild).SetAttribute("class", "MJS_choiceDesc");
                }
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
    }
}