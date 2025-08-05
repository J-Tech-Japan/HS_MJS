// GenerateHTMLButton.XMLExportTocAsJsFiles.cs

using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Xml;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 目次ファイルの作成
        // Webヘルプシステムのナビゲーション機能用に、階層化された目次データをJavaScript形式で出力
        // 目次の階層構造（3レベルまで）を解析し、各レベルごとにJavaScriptファイルを生成する
        public void ExportTocAsJsFiles(XmlDocument objToc, string rootPath, string exportDir, Dictionary<string, string> mergeScript)
        {
            // 各階層の目次JSON文字列を格納する変数
            string htmlToc = "";    // ルートレベル（第1階層）
            string htmlToc1 = "";   // 第2階層
            string htmlToc2 = "";   // 第3階層
            string htmlToc3 = "";   // 第4階層

            // XMLドキュメントのルートアイテムを処理（最上位レベルの目次項目）
            foreach (XmlNode toc in objToc.SelectNodes("/result/item"))
            {
                // ルートレベルの目次JSON構造を作成（常にbook型で、toc1キーを持つ）
                htmlToc = @"{""type"":""book"",""name"":""" + ((XmlElement)toc).GetAttribute("title") + @""",""key"":""toc1""}";

                // 第2階層の子項目を処理
                foreach (XmlNode toc1 in toc.SelectNodes("item"))
                {
                    // 複数項目がある場合はカンマで区切る
                    if (htmlToc1 != "")
                    {
                        htmlToc1 = htmlToc1 + ",";
                    }

                    // 項目の種類を決定（子項目があるかどうかで「book」または「item」を設定）
                    htmlToc1 = htmlToc1 + @"{""type"":""" + (toc1.SelectNodes("item").Count != 0 ? "book" : "item") + @""",""name"":""" + ((XmlElement)toc1).GetAttribute("title") + @"""";

                    // 子項目がある場合（book型）は、対応するJavaScriptファイルのキーを設定
                    // XPath式で先行する同レベルのbook型項目数を計算し、ファイル名を決定
                    if (toc1.SelectNodes("item").Count != 0)
                    {
                        htmlToc1 += @",""key"":""toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + @"""";
                    }

                    // href属性が存在する場合はURL情報を追加（mergeScriptで変換処理）
                    if (((XmlElement)toc1).GetAttribute("href") != "")
                    {
                        htmlToc1 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc1).GetAttribute("href")) + @"""";
                    }

                    htmlToc1 += "}";

                    // 第3階層の子項目を処理
                    foreach (XmlNode toc2 in toc1.SelectNodes("item"))
                    {
                        // 複数項目がある場合はカンマで区切る
                        if (htmlToc2 != "")
                        {
                            htmlToc2 = htmlToc2 + ",";
                        }

                        // 項目の種類を決定（子項目があるかどうかで「book」または「item」を設定）
                        htmlToc2 += @"{""type"":""" + (toc2.SelectNodes("item").Count != 0 ? "book" : "item") + @""",""name"":""" + ((XmlElement)toc2).GetAttribute("title") + @"""";

                        // 子項目がある場合（book型）は、対応するJavaScriptファイルのキーを設定
                        if (toc2.SelectNodes("item").Count != 0)
                        {
                            htmlToc2 += @",""key"":""toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + @"""";
                        }
                        
                        // href属性が存在する場合はURL情報を追加
                        if (((XmlElement)toc2).GetAttribute("href") != "")
                        {
                            htmlToc2 += @",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc2).GetAttribute("href")) + @"""";
                        }

                        htmlToc2 += "}";

                        // 第4階層の子項目を処理（最下位レベル、常にitem型）
                        foreach (XmlNode toc3 in toc2.SelectNodes("item"))
                        {
                            // 複数項目がある場合はカンマで区切る
                            if (htmlToc3 != "")
                            {
                                htmlToc3 += ",";
                            }

                            // 最下位レベルは常にitem型でURL必須
                            htmlToc3 += @"{""type"":""item"",""name"":""" + ((XmlElement)toc3).GetAttribute("title") + @""",""url"":""" + makeHrefWithMerge(mergeScript, ((XmlElement)toc3).GetAttribute("href")) + @"""}";
                        }

                        // 第4階層の項目がある場合、対応するJavaScriptファイルを出力
                        if (htmlToc3 != "")
                        {
                            // XPath式で先行するbook型項目数を計算してファイル名を決定し、whxdataフォルダに保存
                            using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc2.SelectNodes("preceding::item[boolean(item)]").Count + 3) + ".new.js", false, Encoding.UTF8))
                            {
                                // Webヘルプシステム用のJavaScript形式で出力
                                sw.WriteLine("(function() {");
                                sw.WriteLine("var toc =  [" + htmlToc3 + "];");
                                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                                sw.WriteLine("})();");
                            }
                            // 次の項目処理のために文字列をクリア
                            htmlToc3 = "";
                        }
                    }

                    // 第3階層の項目がある場合、対応するJavaScriptファイルを出力
                    if (htmlToc2 != "")
                    {
                        // XPath式で先行するbook型項目数を計算してファイル名を決定
                        using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc" + (toc1.SelectNodes("preceding::item[boolean(item)]").Count + 2) + ".new.js", false, Encoding.UTF8))
                        {
                            // Webヘルプシステム用のJavaScript形式で出力
                            sw.WriteLine("(function() {");
                            sw.WriteLine("var toc =  [" + htmlToc2 + "];");
                            sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                            sw.WriteLine("})();");
                        }
                        // 次の項目処理のために文字列をクリア
                        htmlToc2 = "";
                    }
                }

                // 第2階層のJavaScriptファイルを出力（toc1.new.js）
                using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc1.new.js", false, Encoding.UTF8))
                {
                    // Webヘルプシステム用のJavaScript形式で出力
                    sw.WriteLine("(function() {");
                    sw.WriteLine("var toc =  [" + htmlToc1 + "];");
                    sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                    sw.WriteLine("})();");
                }
            }

            // ルートレベルのJavaScriptファイルを出力（toc.new.js）
            using (StreamWriter sw = new StreamWriter(rootPath + "\\" + exportDir + "\\whxdata\\toc.new.js", false, Encoding.UTF8))
            {
                // Webヘルプシステム用のJavaScript形式で出力
                sw.WriteLine("(function() {");
                sw.WriteLine("var toc =  [" + htmlToc + "];");
                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                sw.WriteLine("})();");
            }
        }
    }
}