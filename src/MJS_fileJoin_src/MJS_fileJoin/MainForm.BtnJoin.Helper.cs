// MainForm.BtnJoin.Helper.cs

using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // 入力バリデーション
        private bool ValidateInput()
        {
            if (tbOutputDir.Text == "")
            {
                MessageBox.Show("出力ディレクトリをご指定ください。");
                return false;
            }
            if (!Directory.Exists(tbOutputDir.Text))
            {
                MessageBox.Show("出力ディレクトリが存在しません。");
                return false;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("格納フォルダをご指定ください。");
                return false;
            }
            if (lbHtmlList.Items.Count == 0)
            {
                MessageBox.Show("変換したHTMLファイルが格納されているフォルダーが登録されていません。");
                return false;
            }
            int fileCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                fileCount += bookInfo[htmlDir].Select("Column1 = true").Count();
            }
            if (fileCount == 0)
            {
                MessageBox.Show("コンテンツが選択されていません。");
                return false;
            }
            foreach (string htmlDir in lbHtmlList.Items)
            {
                if (!Directory.Exists(htmlDir))
                {
                    MessageBox.Show("「" + htmlDir + "」は削除されたか、名前が変更されています。");
                    return false;
                }
            }
            return true;
        }

        // 出力ディレクトリを準備
        private void PrepareOutputDirectory()
        {
            exportDir = textBox2.Text;
            string outputPath = Path.Combine(tbOutputDir.Text, exportDir);

            if (Directory.Exists(outputPath))
            {
                Directory.Delete(outputPath, true);
            }

            Directory.CreateDirectory(outputPath);

            // 最初のHTMLフォルダの内容をコピー
            CopyDirectoryRecursive(lbHtmlList.Items[0].ToString(), outputPath);
        }

        // 出力ディレクトリを準備
        // 暫定的に実装しているメソッド
        // タイムスタンプ付きの新しいフォルダを作成してHTMLフォルダの内容をコピー

        //private void PrepareOutputDirectory()
        //{
        //    // 新しいフォルダ名をタイムスタンプで生成（例: export_20240605_153045）
        //    string newExportDir = textBox2.Text + DateTime.Now.ToString("yyyyMMdd_HHmmss");
        //    string outputPath = Path.Combine(tbOutputDir.Text, newExportDir);

        //    // 新しいディレクトリを作成
        //    Directory.CreateDirectory(outputPath);

        //    // 最初のHTMLフォルダの内容をコピー
        //    CopyDirectoryRecursive(lbHtmlList.Items[0].ToString(), outputPath);

        //    // exportDir変数を新しいフォルダ名に更新（他の処理で利用する場合）
        //    exportDir = newExportDir;
        //}

        // HTMLファイルリストの作成
        private List<string> CreateHtmlFileList()
        {
            var htmlFiles = new List<string>();
            foreach (string htmlDir in lbHtmlList.Items)
            {
                foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
                {
                    htmlFiles.Add(selRow["Column4"].ToString() + ".html");
                }
            }
            return htmlFiles;
        }

        // chbListOutputがチェックされている場合にjoinList.xmlを出力する
        private void OutputJoinListXml()
        {
            if (!chbListOutput.Checked)
                return;

            XmlDocument list = new XmlDocument();
            list.PreserveWhitespace = true;
            list.LoadXml("<joinList></joinList>");
            if (tbChangeTitle.Enabled)
            {
                list.DocumentElement.AppendChild(list.CreateWhitespace("\n\t"));
                list.DocumentElement.AppendChild(list.CreateElement("changeTitle"));
                list.DocumentElement.LastChild.InnerText = tbChangeTitle.Text;
            }
            if (tbAddTop.Enabled)
            {
                list.DocumentElement.AppendChild(list.CreateWhitespace("\n\t"));
                list.DocumentElement.AppendChild(list.CreateElement("addTopLevel"));
                list.DocumentElement.LastChild.InnerText = tbAddTop.Text;
            }

            list.DocumentElement.AppendChild(list.CreateWhitespace("\n\t"));
            XmlNode htmllist = list.DocumentElement.AppendChild(list.CreateElement("htmlList"));

            foreach (string htmlDir in lbHtmlList.Items)
            {
                htmllist.AppendChild(list.CreateWhitespace("\n\t\t"));
                XmlNode htmlitem = htmllist.AppendChild(list.CreateElement("item"));
                ((XmlElement)htmlitem).SetAttribute("src", htmlDir);

                foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
                {
                    htmlitem.AppendChild(list.CreateWhitespace("\n\t\t\t"));
                    XmlNode checkedNode = htmlitem.AppendChild(list.CreateElement("checked"));
                    ((XmlElement)checkedNode).SetAttribute("id", selRow["Column4"].ToString());
                }
                htmlitem.AppendChild(list.CreateWhitespace("\n\t\t"));
            }
            htmllist.AppendChild(list.CreateWhitespace("\n\t"));

            list.DocumentElement.AppendChild(list.CreateWhitespace("\n\t"));
            list.DocumentElement.AppendChild(list.CreateElement("outputDir"));
            ((XmlElement)list.DocumentElement.LastChild).SetAttribute("src", tbOutputDir.Text);
            list.DocumentElement.AppendChild(list.CreateWhitespace("\n"));

            list.Save(Path.Combine(tbOutputDir.Text, "joinList.xml"));
        }

        // インデックスページの準備処理
        private XmlNode PrepareIndexPage(
            string htmlDir,
            string outputDir,
            XmlNode objTocRoot,
            XmlDocument objToc,
            TextBox tbChangeTitle,
            TextBox tbAddTop)
        {
            // index.htmlが未作成かつ元フォルダに存在する場合
            if (!File.Exists(Path.Combine(outputDir, "index.html")) && File.Exists(Path.Combine(htmlDir, "index.html")))
            {
                string indexHtml;
                using (var sr = new StreamReader(Path.Combine(htmlDir, "index.html")))
                {
                    indexHtml = sr.ReadToEnd();
                }

                if (tbChangeTitle.Enabled)
                {
                    indexHtml = Regex.Replace(indexHtml, "<title>.+</title>", "<title>" + tbChangeTitle.Text + "</title>", RegexOptions.IgnoreCase);
                }
                else if (tbAddTop.Enabled)
                {
                    indexHtml = Regex.Replace(indexHtml, "<title>.+</title>", "<title>" + tbAddTop.Text + "</title>", RegexOptions.IgnoreCase);
                }

                using (var sw = new StreamWriter(Path.Combine(outputDir, "index.html"), false, Encoding.UTF8))
                {
                    sw.Write(indexHtml);
                }

                string coverPage = Regex.Match(indexHtml, @"gDefaultTopic = ""#(.+?)"";").Groups[1].Value;
                File.Copy(Path.Combine(htmlDir, coverPage), Path.Combine(outputDir, coverPage));

                if (coverPage.Contains("00000"))
                {
                    CopyDirectoryWithOverwriteOption(
                        Path.Combine(Path.Combine(htmlDir, "template"), "images"),
                        Path.Combine(Path.Combine(outputDir, "template"), "images"),
                        true);
                }

                if (tbAddTop.Enabled)
                {
                    objTocRoot.InnerXml = @"<item title=""" + tbAddTop.Text + @"""/>";
                    objTocRoot = objTocRoot.LastChild;
                }
            }
            return objTocRoot;
        }

        // 目次アイテムごとのHTMLファイルを処理し、gTopicIdを書き換えて保存
        private void UpdateHtmlFilesWithTocId(XmlDocument objToc, string outputDir, string exportDir)
        {
            foreach (XmlElement tocItem in objToc.SelectNodes(".//item[boolean(@href)]"))
            {
                StreamReader sr = null;
                StreamWriter sw = null;
                string htmlFileName;
                if (tocItem.GetAttribute("href").Contains("#"))
                {
                    string[] parts = tocItem.GetAttribute("href").Split('#');
                    if (parts.Length >= 2)
                    {
                        string result = parts[1];
                        htmlFileName = result + ".html";
                    }
                    else
                    {
                        continue;
                    }
                }
                else
                {
                    htmlFileName = tocItem.GetAttribute("href") + ".html";
                }

                string htmlFilePath = Path.Combine(tbOutputDir.Text, exportDir, htmlFileName);
                sr = new StreamReader(htmlFilePath);
                string selHtml = sr.ReadToEnd();
                sr.Close();

                string tocId = "";
                foreach (XmlElement objTocItem in tocItem.SelectNodes("ancestor-or-self::item"))
                {
                    if (tocId != "")
                    {
                        tocId += ".";
                    }
                    int precedingItemCount = objTocItem.SelectNodes("preceding-sibling::item[boolean(item)]|self::item[boolean(item)]").Count;
                    tocId += precedingItemCount.ToString();
                    if (objTocItem.SelectSingleNode("item") == null)
                    {
                        tocId += "_";
                        tocId += (objTocItem.SelectNodes("preceding-sibling::item[not(boolean(item)) and (count(preceding-sibling::item[boolean(item)]) = " + precedingItemCount + ")]").Count + 1).ToString();
                    }
                }

                selHtml = Regex.Replace(selHtml, @"(?<=gTopicId[\s]*=[\s]*"")[^""]*(?="")", tocId);

                // 書き込み
                sw = new StreamWriter(htmlFilePath, false, Encoding.UTF8);
                sw.Write(selHtml);
                sw.Close();
            }
        }

        // HTML出力後の処理をメソッド化
        private void AfterHtmlOutput(string outputDirPath)
        {
            DialogResult selectMess = MessageBox.Show(
                outputDirPath + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？",
                "HTML出力成功",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (selectMess == DialogResult.Yes)
            {
                try
                {
                    Process.Start(Path.Combine(outputDirPath, "index.html"));
                }
                catch
                {
                    MessageBox.Show("HTMLの出力に失敗しました。", "HTML出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
            }
            if (checkBox2.Checked)
            {
                tabControl1.SelectedIndex = 1;
                listBox2.Items.Clear();
                listBox2.Items.Add(outputDirPath);
                button12.PerformClick();
            }
        }

        private void CreateToc(XmlNode objToc)
        {
            string htmlToc = "";
            foreach (XmlNode toc in objToc.SelectNodes("item"))
            {
                if (htmlToc != "")
                {
                    htmlToc = htmlToc + ",";
                }

                htmlToc = htmlToc + @"{""type"":""";

                if (toc.SelectNodes("item").Count != 0)
                {
                    htmlToc = htmlToc + "book";
                }
                else
                {
                    htmlToc = htmlToc + "item";
                }

                htmlToc += @""",""name"":""" + ((XmlElement)toc).GetAttribute("title") + @"""";

                if (toc.SelectNodes("item").Count != 0)
                {
                    htmlToc += @",""key"":""toc" + (toc.SelectNodes("preceding::item[boolean(item)]").Count + toc.SelectNodes("ancestor-or-self::item").Count) + @"""";
                }

                if (((XmlElement)toc).GetAttribute("href") != "")
                {
                    htmlToc += @",""url"":""" + ((XmlElement)toc).GetAttribute("href") + @".html""";
                }

                htmlToc += "}";

                if (toc.SelectNodes("item").Count != 0)
                {
                    CreateToc(toc);
                }
            }

            if (htmlToc != "")
            {

                if (Regex.IsMatch(htmlToc, @"""url""\s*:\s*""([^""]*)#([^""]*)"""))
                {
                    htmlToc = Regex.Replace(htmlToc, @"""url""\s*:\s*""([^""]*)#([^""]*)""", match =>
                    {
                        string url = match.Groups[1].Value;
                        string hash = match.Groups[2].Value;

                        return $@"""url"": ""{url}.html#{hash}""";
                    });
                }

                int itemCount = objToc.SelectNodes("preceding::item[boolean(item)]").Count + objToc.SelectNodes("ancestor-or-self::item").Count;
                StreamWriter sw = new StreamWriter(tbOutputDir.Text + "\\" + exportDir + "\\whxdata\\toc" + ((itemCount != 0) ? itemCount.ToString() : "") + ".new.js", false, Encoding.UTF8);
                sw.WriteLine("(function() {");
                sw.WriteLine("var toc =  [" + htmlToc + "];");
                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                sw.WriteLine("})();");
                sw.Close();
            }
        }

        private void MergeHeaderFile()
        {
            string mergeText = "";
            string headerFilePath = "";
            foreach (string file in lbHtmlList.Items)
            {
                string[] files = Directory.GetFiles(file, "*.html");
                string pathName = "";

                foreach (string f in files)
                {
                    if (Regex.IsMatch(Path.GetFileName(f), @"^[A-Z]{3}\d+\.html$"))
                    {
                        pathName = Regex.Replace(Path.GetFileName(f), @"\d+\.html$", "");
                        break;
                    }
                }
                using (StreamReader sr = new StreamReader(Path.Combine(Path.Combine(Path.GetDirectoryName(file), "headerFile"), pathName + ".txt")))
                {
                    mergeText += sr.ReadToEnd();
                }
                if (!headerFilePath.Contains(pathName + "_")) headerFilePath += pathName + "_";
            }
            List<string> ls = new List<string>();

            if (!Directory.Exists(Path.Combine(tbOutputDir.Text, "headerFile"))) Directory.CreateDirectory(Path.Combine(tbOutputDir.Text, "headerFile"));
            using (StreamWriter sw = new StreamWriter(Path.Combine(tbOutputDir.Text, "headerFile\\" + Regex.Replace(headerFilePath, @"_$", "")) + ".txt"))
            using (StringReader sr = new StringReader(mergeText))
            {
                while (sr.Peek() > 0)
                {
                    string lineText = sr.ReadLine();
                    if (!ls.Contains(Regex.Replace(lineText, @"^.*?\t.*?\t(.*?)$", "$1")))
                    {
                        ls.Add(Regex.Replace(lineText, @"^.*?\t.*?\t(.*?)$", "$1"));
                        sw.WriteLine(lineText);
                    }
                }
            }
        }
    }
}
