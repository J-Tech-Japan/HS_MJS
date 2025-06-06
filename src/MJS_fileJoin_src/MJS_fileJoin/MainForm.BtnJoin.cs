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
        // webHelpを結合し、指定した出力ディレクトリに統合HTMLコンテンツを生成する
        private void btnJoin_Click(object sender, EventArgs e)
        {
            Cursor prevCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            // バリデーション
            if (!ValidateInput())
            {
                Cursor.Current = prevCursor;
                return;
            }

            StreamReader sr = null;
            StreamWriter sw = null;

            List<string> errorList = new List<string>();

            //テンプレート展開
            //System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            //using (Stream stream = assembly.GetManifestResourceStream("MJS_fileJoin.htmlTemplates.zip"))
            //{
            //    FileStream fs = File.Create(tbOutputDir.Text + "\\htmlTemplates.zip");
            //    stream.Seek(0, SeekOrigin.Begin);
            //    stream.CopyTo(fs);
            //    fs.Close();
            //}


            //if (Directory.Exists(tbOutputDir.Text + "\\htmlTemplates"))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\htmlTemplates", true);
            //}

            //ZipFile.ExtractToDirectory(tbOutputDir.Text + "\\htmlTemplates.zip", tbOutputDir.Text);

            //if (Directory.Exists(tbOutputDir.Text + "\\" + exportDir))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\" + exportDir, true);
            //}
            //Directory.Move(tbOutputDir.Text + "\\htmlTemplates", tbOutputDir.Text + "\\" + exportDir);

            //File.Delete(tbOutputDir.Text + "\\htmlTemplates.zip");

            //' Ver - 2023.16.08 - VyNL - ↑ - 追加'

            // 出力ディレクトリの準備
            // ここでexportDir変数に新しいフォルダ名が格納される
            PrepareOutputDirectory();

            //if (Directory.Exists(tbOutputDir.Text + "\\" + exportDir))
            //{
            //    Directory.Delete(tbOutputDir.Text + "\\" + exportDir, true);
            //}

            //Directory.CreateDirectory(tbOutputDir.Text + "\\" + exportDir);

            //CopyDirectory(lbHtmlList.Items[0].ToString(), tbOutputDir.Text + "\\" + exportDir);

            XmlDocument objToc = new XmlDocument();
            XmlNode objTocRoot = null;

            XmlDocument searchWords = new System.Xml.XmlDocument();
            searchWords.LoadXml("<div class='search'></div>");

            objToc.LoadXml(@"<result></result>");
            objTocRoot = objToc.DocumentElement;

            //各webHelpフォルダ処理

            // HTMLファイルリストの作成
            List<string> lsfiles = CreateHtmlFileList();

            //List<string> lsfiles = new List<string>();
            //foreach (string htmlDir in lbHtmlList.Items)
            //    foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
            //        lsfiles.Add(selRow["Column4"].ToString() + ".html");

            int picCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                picCount++;
                List<string> pics = new List<string>();
                foreach (string file in Directory.GetFiles(htmlDir + "\\pict", "*.*", SearchOption.AllDirectories))
                    pics.Add(Path.GetFileName(file));

                string outputDir = Path.Combine(tbOutputDir.Text, exportDir);

                // インデックスページ準備
                objTocRoot = PrepareIndexPage(htmlDir, outputDir, objTocRoot, objToc, tbChangeTitle, tbAddTop);

                // HTMLファイルのコピーと加工処理
                ProcessHtmlFiles(htmlDir, outputDir, picCount, lsfiles, objTocRoot, objToc, searchWords, errorList);
            }

            //全文検索ファイル出力
            sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, "search.js"), false, Encoding.UTF8);
            //            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2"), @"(?<=>)([^<]*?)'([^<]*?)(?=<)", "$1&apos;$2")));
            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;")));
            sw.Close();
            //Dictionary<string, string> mergeScript = new Dictionary<string, string>();
            foreach (XmlElement tocItem in objToc.SelectNodes(".//item[boolean(@href)]"))
            {
                if (tocItem.GetAttribute("href").Contains("#"))
                {
                    string[] parts = tocItem.GetAttribute("href").Split('#');

                    if (parts.Length >= 2)
                    {
                        string result = parts[1];
                        sr = new StreamReader(Path.Combine(tbOutputDir.Text, exportDir, result + ".html"));

                    }

                }
                else
                {
                    sr = new StreamReader(Path.Combine(tbOutputDir.Text, exportDir, tocItem.GetAttribute("href") + ".html"));
                }
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
                if (tocItem.GetAttribute("href").Contains("#"))
                {
                    string[] parts = tocItem.GetAttribute("href").Split('#');

                    if (parts.Length >= 2)
                    {
                        string result = parts[1];
                        sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, result + ".html"), false, Encoding.UTF8);
                    }

                }
                else
                {
                    sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, tocItem.GetAttribute("href") + ".html"), false, Encoding.UTF8);
                }

                //string pattern = @"mergePage = {(.*?)};";
                //Match match = Regex.Match(selHtml, pattern, RegexOptions.Singleline);

                //if (match.Success)
                //{
                //    string mergePageData = match.Groups[1].Value;

                //    // Extract key-value pairs from mergePageData
                //    pattern = @"(\w+):'(\w+)'";
                //    MatchCollection matches = Regex.Matches(mergePageData, pattern);

                //    // Output the extracted key-value pairs
                //    foreach (Match m in matches)
                //    {
                //        string key = m.Groups[1].Value;
                //        string value = m.Groups[2].Value;
                //        if (!String.IsNullOrEmpty(key) && !String.IsNullOrEmpty(key)&& !mergeScript.Any(x => x.Key == key && x.Value == value))
                //            mergeScript.Add(key, value);
                //    }
                //}
                sw.Write(selHtml);
                sw.Close();
            }

            //目次出力
            createToc(objToc.DocumentElement);

            // chbListOutputがチェックされている場合にjoinList.xmlを出力する
            OutputJoinListXml();

            

            //書誌情報ファイルのマージ
            mergeHeaderFile();

            Cursor.Current = prevCursor;

            DialogResult selectMess = MessageBox.Show(tbOutputDir.Text + "\\" + exportDir + "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？", "HTML出力成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (selectMess == DialogResult.Yes)
            {
                try
                {
                    Process.Start(tbOutputDir.Text + "\\" + exportDir + @"\index.html");
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
                listBox2.Items.Add(tbOutputDir.Text + "\\" + exportDir);
                button12.PerformClick();
            }
        }
    }
}
