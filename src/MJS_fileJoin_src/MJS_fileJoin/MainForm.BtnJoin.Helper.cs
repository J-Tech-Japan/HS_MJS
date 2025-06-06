using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
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
        //private void PrepareOutputDirectory()
        //{
        //    string outputPath = Path.Combine(tbOutputDir.Text, exportDir);

        //    if (Directory.Exists(outputPath))
        //    {
        //        Directory.Delete(outputPath, true);
        //    }

        //    Directory.CreateDirectory(outputPath);

        //    // 最初のHTMLフォルダの内容をコピー
        //    CopyDirectory(lbHtmlList.Items[0].ToString(), outputPath);
        //}

        // 出力ディレクトリを準備
        private void PrepareOutputDirectory()
        {
            // 新しいフォルダ名をタイムスタンプで生成（例: export_20240605_153045）
            string newExportDir = textBox2.Text + DateTime.Now.ToString("yyyyMMdd_HHmmss");
            string outputPath = Path.Combine(tbOutputDir.Text, newExportDir);

            // 新しいディレクトリを作成
            Directory.CreateDirectory(outputPath);

            // 最初のHTMLフォルダの内容をコピー
            CopyDirectory(lbHtmlList.Items[0].ToString(), outputPath);

            // exportDir変数を新しいフォルダ名に更新（他の処理で利用する場合）
            exportDir = newExportDir;
        }

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
                    CopyDirectory(
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
    }
}
