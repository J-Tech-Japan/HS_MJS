﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 検索用HTMLファイルと検索用JSファイルを生成するメソッド
        // objBody: 本文XML, objToc: 目次XML, mergeScript: マージ用スクリプト, searchJs: 検索JSテンプレート
        public void GenerateSearchFiles(XmlDocument objBody, string rootPath, string exportDir, string docid, string htmlTemplate1, string htmlTemplate2, string htmlCoverTemplate1, string htmlCoverTemplate2, XmlDocument objToc, Dictionary<string, string> mergeScript, string searchJs)
        {
            StreamWriter sw;

            // 検索ワード格納用XMLを初期化
            XmlDocument searchWords = new XmlDocument();
            searchWords.LoadXml("<div class='search'></div>");

            // 本文XML内の各div（ページ）ごとに処理
            foreach (XmlNode splithtml in objBody.SelectNodes("/result/div"))
            {
                string thisId = ((XmlElement)splithtml).GetAttribute("id");

                // id, style属性を削除
                ((XmlElement)splithtml).RemoveAttribute("id");
                ((XmlElement)splithtml).RemoveAttribute("style");

                // 表紙ページの場合
                if (thisId == docid + "00000")
                {
                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + thisId + ".html", false, Encoding.UTF8);
                    string coverBody = "";

                    // manual_クラスを持つ要素を抽出（表紙用）
                    foreach (XmlNode coverItem in splithtml.SelectNodes(".//*[starts-with(@class, 'manual_')]"))
                    {
                        coverBody += coverItem.OuterXml;
                    }

                    // 表紙テンプレートを書き込み（coverBodyは未使用）
                    //sw.Write(htmlCoverTemplate1 + coverBody + htmlCoverTemplate2);
                    sw.Write(htmlCoverTemplate1 + htmlCoverTemplate2);
                    sw.Close();
                }
                else
                {
                    string htmlTemplate1cpy = htmlTemplate1;
                    // 目次XMLに該当ページが存在する場合
                    if (objToc.SelectSingleNode(".//item[@href = '" + thisId + "']") != null)
                    {
                        // タイトルをテンプレートに埋め込む
                        htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, "<title></title>", "<title>" + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + thisId + "']")).GetAttribute("title") + "</title>");
                        string breadcrumb = "";
                        // パンくずリスト表示用div生成
                        XmlElement breadcrumbDisplay = objBody.CreateElement("div");
                        breadcrumbDisplay.SetAttribute("style", "text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;");

                        string tocId = "";

                        // 目次階層をたどりパンくずリスト生成
                        foreach (XmlNode tocItem in objToc.SelectNodes(".//item[@href = '" + thisId + "']/ancestor-or-self::item"))
                        {
                            if (breadcrumb != "")
                            {
                                breadcrumb += " > ";
                                breadcrumbDisplay.AppendChild(objBody.CreateTextNode(" > "));
                            }
                            breadcrumb += ((XmlElement)tocItem).GetAttribute("title");

                            // href属性がある場合はリンク化
                            if (tocItem.SelectSingleNode("@href") != null)
                            {
                                breadcrumbDisplay.AppendChild(objBody.CreateElement("a"));
                                ((XmlElement)breadcrumbDisplay.LastChild).SetAttribute("href", "./" + makeHrefWithMerge(mergeScript, ((XmlElement)tocItem).GetAttribute("href")) + "");
                                breadcrumbDisplay.LastChild.InnerText = ((XmlElement)tocItem).GetAttribute("title");
                            }
                            else
                            {
                                breadcrumbDisplay.AppendChild(objBody.CreateTextNode(((XmlElement)tocItem).GetAttribute("title")));
                            }

                            // tocId（目次階層ID）生成
                            if (tocId != "")
                            {
                                tocId += ".";
                            }

                            int precedingItemCount = tocItem.SelectNodes("preceding-sibling::item[boolean(item)]|self::item[boolean(item)]").Count;
                            tocId += precedingItemCount.ToString();

                            if (tocItem.SelectSingleNode("item") == null)
                            {
                                tocId += "_";
                                tocId += (tocItem.SelectNodes("preceding-sibling::item[not(boolean(item)) and (count(preceding-sibling::item[boolean(item)]) = " + precedingItemCount + ")]").Count + 1).ToString();
                            }
                        }

                        // テンプレート内の♪をtocIdで置換
                        htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, "♪", tocId);

                        // 検索用テキスト生成（エスケープ処理）
                        string searchText = splithtml.InnerText.Replace("&", "&amp;").Replace("<", "&lt;");
                        string displayText = searchText;

                        // 長すぎる場合は90文字で省略
                        if (searchText.Length >= 90)
                        {
                            displayText = displayText.Substring(0, 90) + " ...";
                        }

                        // 全角→半角変換テーブル
                        //string[] wide = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン" };
                        //string[] narrow = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" };

                        //// 全角→半角変換
                        //for (int i = 0; i < wide.Length; i++)
                        //{
                        //    searchText = Regex.Replace(searchText, wide[i], narrow[i]);
                        //}

                        // 全角→半角変換メソッドをUtils.TextProcessing.csから呼び出し
                        searchText = Utils.ConvertWideToNarrow(searchText);

                        searchText = searchText.ToLower();

                        // 検索ワード情報をsearchWordsに追加
                        searchWords.DocumentElement.AppendChild(searchWords.CreateElement("div"));
                        ((XmlElement)searchWords.DocumentElement.LastChild).SetAttribute("id", thisId);
                        searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>" + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>" + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + thisId + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='displayText'>" + displayText + "</div><div class='search_word'>" + searchText + "</div>";

                        // パンくず情報をmetaタグに埋め込む
                        htmlTemplate1cpy = Regex.Replace(htmlTemplate1cpy, @"<meta name=""topic-breadcrumbs"" content="""" />", @"<meta name=""topic-breadcrumbs"" content=""" + breadcrumb + @""" />");
                        
                        // パンくず表示divを本文先頭に挿入
                        splithtml.InsertBefore(breadcrumbDisplay, splithtml.FirstChild);
                    }

                    // ページ内リンクのhrefを調整
                    if (!String.IsNullOrEmpty(thisId))
                    {
                        foreach (XmlNode nd in splithtml.SelectNodes(".//a[contains(@href, '" + thisId + ".html')]"))
                        {
                            if (((XmlElement)nd).GetAttribute("href").Contains("#"))
                                ((XmlElement)nd).SetAttribute("href", Regex.Replace(((XmlElement)nd).GetAttribute("href"), @"^.*?(#.*?)$", "$1", RegexOptions.Singleline));
                            else
                                ((XmlElement)nd).SetAttribute("href", "#");
                        }
                    }

                    // HTML本文生成
                    sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + thisId + ".html", false, Encoding.UTF8);
                    string htmlBody = htmlTemplate1cpy + splithtml.OuterXml + htmlTemplate2;
                    
                    // 手順番号spanのクラス付与・変換などの正規表現処理
                    htmlBody = Regex.Replace(htmlBody, @"<p[^>]*?class=""MJS_oflow_step([^""]*?)""[^>]*?>(.*?)<span[^>]*?>(.*?)</span>(.*?)</p>", @"<p class=""MJS_oflow_step$1""><span class=""MJS_oflow_stepNum$2"">$3</span>$4</p>", RegexOptions.Singleline);
                    
                    // 特定文字（è）を手順結果spanに変換
                    htmlBody = Regex.Replace(htmlBody, @"<span class=""MJS_oflow_stepNum"">(è)</span>", @"<span class=""MJS_oflow_stepResult""></span>", RegexOptions.Singleline);
                    
                    // 手順結果pタグ内のspan削除
                    htmlBody = Regex.Replace(htmlBody, @"<p[^>]*?class=""MJS_oflow_stepResult([^""]*?)""[^>]*?>(.*?)<span[^>]*?>(.*?)</span>(.*?)</p>", @"<p class=""MJS_oflow_stepResult"">$4</p>", RegexOptions.Singleline);
                    
                    // 手順番号span内の入れ子spanを除去
                    htmlBody = Regex.Replace(htmlBody, @"<span class=""MJS_oflow_stepNum""><span[^>]*?>(.*?)</span>(.*?)</span>", @"<span class=""MJS_oflow_stepNum"">$1$2</span>", RegexOptions.Singleline);

                    sw.Write(htmlBody);
                    sw.Close();
                }
            }

            // 検索用JSファイル生成
            sw = new StreamWriter(rootPath + "\\" + exportDir + "\\search.js", false, Encoding.UTF8);
            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;").Replace(@"\u", @"\\u").Replace(@"\U", @"\\U")));
            sw.Close();

            // 表紙HTMLが存在しない場合は生成
            if (!File.Exists(rootPath + "\\" + exportDir + "\\" + docid + "00000.html"))
            {
                sw = new StreamWriter(rootPath + "\\" + exportDir + "\\" + docid + "00000.html", false, Encoding.UTF8);
                sw.Write(htmlCoverTemplate1 + htmlCoverTemplate2);
                sw.Close();
            }
        }
    }
}
