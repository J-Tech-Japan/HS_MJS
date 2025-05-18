using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;
using Microsoft.Office.Interop.Word;
using DocMergerComponent;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.Net;
using System.Xml.Linq;
using System.Runtime.InteropServices.ComTypes;
using System.Runtime.InteropServices;


namespace MJS_fileJoin
{
    public partial class MainForm : Form
    {
        public Dictionary<string, System.Data.DataTable> bookInfo = new Dictionary<string, System.Data.DataTable>();
        public string exportDir = "";
        public List<ListViewItem> logen = new List<ListViewItem>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnSelectJoinList_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(Path.GetDirectoryName(tbSelectJoinList.Text)))
                {
                    openFileDialog1.InitialDirectory = Path.GetDirectoryName(tbSelectJoinList.Text);
                }
            }
            catch (Exception ex)
            {
            }
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            tbSelectJoinList.Text = openFileDialog1.FileName;
        }

        private void tbSelectJoinList_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            for (int i = 0; i < s.Length; i++)
            {
                if (File.Exists(s[i]))
                {
                    tbSelectJoinList.Text = s[i];
                    break;
                }
            }
        }

        private void tbSelectJoinList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void chbChangeTitle_CheckedChanged(object sender, EventArgs e)
        {
            if (chbChangeTitle.Checked)
            {
                tbChangeTitle.Enabled = true;
            }
            else
            {
                tbChangeTitle.Enabled = false;
            }
        }

        private void chbAddTop_CheckedChanged(object sender, EventArgs e)
        {
            if (chbAddTop.Checked)
            {
                tbAddTop.Enabled = true;
            }
            else
            {
                tbAddTop.Enabled = false;
            }
        }

        private void btnHtmlListFile_Click(object sender, EventArgs e)
        {
            if ((folderBrowserDialog1.ShowDialog() == DialogResult.OK) && !lbHtmlList.Items.Contains(folderBrowserDialog1.SelectedPath))
            {
                addHtmlDir(folderBrowserDialog1.SelectedPath);
            }
        }
        
        private void lbHtmlList_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            List<string> webHelpFol = new List<string>();
            foreach (string folder in s)
            {
                if (File.Exists(folder)) continue;
                if (Path.GetFileName(folder) == "webHelp")
                    webHelpFol.Add(folder);
                else
                {
                    string[] fol = Directory.GetDirectories(folder, "webHelp", SearchOption.AllDirectories);
                    foreach (string webhelp in fol) webHelpFol.Add(webhelp);
                }
            }

            for (int i = 0; i < webHelpFol.Count; i++)
            {
                if (!addHtmlDir(webHelpFol[i])) continue;
            }
        }

        private void lbHtmlList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void lbHtmlList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbHtmlList.SelectedIndex == -1)
            {
                dataGridView1.DataSource = null;
            }
            else if (bookInfo.ContainsKey(lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString()))
            {
                dataGridView1.AutoGenerateColumns = true;
                dataGridView1.DataSource = bookInfo[lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString()];
                dataGridView1.Columns[0].Width = 32;
                dataGridView1.Columns[1].Width = 45;
                dataGridView1.Columns[2].Width = 180;
                dataGridView1.Columns[3].Width = 70;
            }
        }

        private void btnHtmlListUp_Click(object sender, EventArgs e)
        {
            if ((lbHtmlList.SelectedIndex != -1) && (lbHtmlList.SelectedIndex != 0))
            {
                object tmp = lbHtmlList.Items[lbHtmlList.SelectedIndex - 1];
                lbHtmlList.Items[lbHtmlList.SelectedIndex - 1] = lbHtmlList.Items[lbHtmlList.SelectedIndex];
                lbHtmlList.Items[lbHtmlList.SelectedIndex] = tmp;
                lbHtmlList.SelectedIndex = lbHtmlList.SelectedIndex - 1;
            }
        }

        private void btnHtmlListDown_Click(object sender, EventArgs e)
        {
            if ((lbHtmlList.SelectedIndex != -1) && (lbHtmlList.SelectedIndex != (lbHtmlList.Items.Count - 1)))
            {
                object tmp = lbHtmlList.Items[lbHtmlList.SelectedIndex + 1];
                lbHtmlList.Items[lbHtmlList.SelectedIndex + 1] = lbHtmlList.Items[lbHtmlList.SelectedIndex];
                lbHtmlList.Items[lbHtmlList.SelectedIndex] = tmp;
                lbHtmlList.SelectedIndex = lbHtmlList.SelectedIndex + 1;
            }
        }

        private void btnHtmlListDel_Click(object sender, EventArgs e)
        {
            if (lbHtmlList.SelectedIndex != -1)
            {
                bookInfo.Remove(lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString());
                lbHtmlList.Items.RemoveAt(lbHtmlList.SelectedIndex);
            }
        }

        private void btnOutputDir_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                tbOutputDir.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void tbOutputDir_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            for (int i = 0; i < s.Length; i++)
            {
                if (Directory.Exists(s[i]))
                {
                    tbOutputDir.Text = s[i];
                    break;
                }
            }
        }

        private void tbOutputDir_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btnJoin_Click(object sender, EventArgs e)
        {
            Cursor prevCursor = Cursor.Current;
            Cursor.Current = Cursors.WaitCursor;

            StreamReader sr = null;
            StreamWriter sw = null;

            if (tbOutputDir.Text == "")
            {
                MessageBox.Show("出力ディレクトリをご指定ください。");
                return;
            }
            if (!Directory.Exists(tbOutputDir.Text))
            {
                MessageBox.Show("出力ディレクトリが存在しません。");
                return;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("格納フォルダをご指定ください。");
                return;
            }
            else exportDir = textBox2.Text;
            if (lbHtmlList.Items.Count == 0)
            {
                MessageBox.Show("変換したHTMLファイルが格納されているフォルダーが登録されていません。");
                return;
            }
            int fileCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                fileCount += bookInfo[htmlDir].Select("Column1 = true").Count();
            }
            if (fileCount == 0)
            {
                MessageBox.Show("コンテンツが選択されていません。");
                return;
            }
            foreach (string htmlDir in lbHtmlList.Items)
            {
                if (!Directory.Exists(htmlDir))
                {
                    MessageBox.Show("「" + htmlDir + "」は削除されたか、名前が変更されています。");
                    return;
                }
            }

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
            if (Directory.Exists(tbOutputDir.Text + "\\" + exportDir))
            {
                Directory.Delete(tbOutputDir.Text + "\\" + exportDir, true);
            }

            Directory.CreateDirectory(tbOutputDir.Text + "\\" + exportDir);

            CopyDirectory(lbHtmlList.Items[0].ToString(), tbOutputDir.Text + "\\" + exportDir);

            XmlDocument objToc = new XmlDocument();
            XmlNode objTocRoot = null;

            XmlDocument searchWords = new System.Xml.XmlDocument();
            searchWords.LoadXml("<div class='search'></div>");

            objToc.LoadXml(@"<result></result>");
            objTocRoot = objToc.DocumentElement;

            //各webHelpフォルダ処理

            List<string> lsfiles = new List<string>();
            foreach (string htmlDir in lbHtmlList.Items)
                foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
                    lsfiles.Add(selRow["Column4"].ToString() + ".html");

            int picCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                picCount++;
                List<string> pics = new List<string>();
                foreach (string file in Directory.GetFiles(htmlDir + "\\pict", "*.*", SearchOption.AllDirectories))
                    pics.Add(Path.GetFileName(file));

                string outputDir = Path.Combine(tbOutputDir.Text, exportDir);

                //インデックスページ準備
                if (!File.Exists(Path.Combine(outputDir, "index.html")) && File.Exists(Path.Combine(htmlDir, "index.html")))
                {
                    sr = new StreamReader(Path.Combine(htmlDir, "index.html"));
                    string indexHtml = sr.ReadToEnd();
                    sr.Close();

                    if (tbChangeTitle.Enabled)
                    {
                        indexHtml = Regex.Replace(indexHtml, "<title>.+</title>", "<title>" + tbChangeTitle.Text + "</title>", RegexOptions.IgnoreCase);
                    }
                    else if (tbAddTop.Enabled)
                    {
                        indexHtml = Regex.Replace(indexHtml, "<title>.+</title>", "<title>" + tbAddTop.Text + "</title>", RegexOptions.IgnoreCase);
                    }

                    sw = new StreamWriter(Path.Combine(outputDir, "index.html"), false, Encoding.UTF8);
                    sw.Write(indexHtml);
                    sw.Close();

                    string coverPage = Regex.Match(indexHtml, @"gDefaultTopic = ""#(.+?)"";").Groups[1].Value;
                    File.Copy(Path.Combine(htmlDir, coverPage), Path.Combine(outputDir, coverPage));

                    if (coverPage.Contains("00000"))
                    {
                        CopyDirectory(Path.Combine(Path.Combine(htmlDir, "template"), "images"), Path.Combine(Path.Combine(outputDir, "template"), "images"), true);
                    }

                    if (tbAddTop.Enabled)
                    {
                        objTocRoot.InnerXml = @"<item title=""" + tbAddTop.Text + @"""/>";
                        objTocRoot = objTocRoot.LastChild;
                    }
                }
                foreach (DataRow selRow in bookInfo[htmlDir].Select("Column1 = true"))
                {
                    if (!File.Exists(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html")))
                    {
                        errorList.Add("「" + Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html") + "」は存在しません。");
                        continue;
                    }

                    if (File.Exists(Path.Combine(outputDir, selRow["Column4"].ToString() + ".html")) && selRow["Column4"].ToString().Contains("00000"))
                    {

                        continue;
                    }

                    File.Copy(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html"), Path.Combine(outputDir, selRow["Column4"].ToString() + ".html"), true);

                    sr = new StreamReader(Path.Combine(htmlDir, selRow["Column4"].ToString() + ".html"));
                    string selHtml = sr.ReadToEnd();
                    sr.Close();

                    string[] coverKINs = { "EdgeTracker_logo50mm.png", "hyousi.png", "MJS_LOGO_255.gif" };
                    foreach (string coverKIN in coverKINs)
                    {
                        if (File.Exists(Path.Combine(htmlDir, "pict", coverKIN)) && !File.Exists(Path.Combine(outputDir, "pict", coverKIN)))
                            File.Copy(Path.Combine(htmlDir, "pict", coverKIN), Path.Combine(outputDir, "pict", coverKIN));
                    }

                    if (Regex.IsMatch(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\].+?"""))
                    {
                        //string dirName = Path.Combine("pict", selRow["Column4"].ToString().Substring(0, 3));
                        string dirName = "pict";
                        if (!Directory.Exists(Path.Combine(outputDir, dirName)))
                        {
                            Directory.CreateDirectory(Path.Combine(outputDir, dirName));
                        }

                        foreach (Match m in Regex.Matches(selHtml, @"<img(?: [^ />]+)* src=""pict[/\\](.+?)"""))
                        {
                            if (!File.Exists(Path.Combine(outputDir, dirName, Path.GetFileNameWithoutExtension(m.Groups[1].Value) + "_" + picCount.ToString("00") + Path.GetExtension(m.Groups[1].Value))))
                            {
                                File.Copy(Path.Combine(htmlDir, "pict", m.Groups[1].Value), Path.Combine(outputDir, dirName, Path.GetFileNameWithoutExtension(m.Groups[1].Value) + "_" + picCount.ToString("00") + Path.GetExtension(m.Groups[1].Value)));
                            }
                        }

                        selHtml = Regex.Replace(selHtml, @"(<img(?: [^ />]+)* src="")pict[/\\](.+?)(\.\w+"")", "$1" + dirName + "/$2_" + picCount.ToString("00") + "$3");

                        //selHtml = Regex.Replace(selHtml, @"(<img(?: [^ />]+)* src="")pict/(.+?"")", "$1" + dirName + "/$2");
                    }

                    if (Regex.IsMatch(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>"))
                    {
                        string[] breadcrumbs = Regex.Replace(Regex.Match(selHtml, @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">((?:.(?!</div>))+.)</div>").Groups[1].Value, "<.+?>", "").Split(new string[] { " &gt; " }, StringSplitOptions.None);
                        // get href by regex from selHtml
                        //var urls = Regex.Match(selHtml, "<a href\\s*=\\s*\"(?<url>.*?)\"").Groups["url"].Value;

                        Regex r = new Regex(@"<a.*?href=("")(?<href>.*?)(""|').*?>(?<value>.*?)</a>");
                        MatchCollection urls2 = r.Matches(selHtml);


                        for (int i = 0; i < breadcrumbs.Length; i++)
                        {
                            // get href by from urls2
                            string urls = "";
                            string title = "";
                            foreach (Match match in urls2)
                            {
                                title = match.Groups["value"].Value.ToString();
                                if (title == breadcrumbs[i])
                                {
                                    urls = match.Groups["href"].Value.ToString();
                                } else
                                {
                                    urls = "";
                                }
                            }


                            string itemNodeLevel = "";
                            for (int j = 0; j <= i; j++)
                            {
                                itemNodeLevel += "/item[last()]";
                            }

                            if (objTocRoot.SelectSingleNode("." + itemNodeLevel + "[@title='" + breadcrumbs[i] + "']") == null)
                            {
                                itemNodeLevel = "";
                                for (int j = 0; j < i; j++)
                                {
                                    itemNodeLevel += "/item[last()]";
                                }
                                XmlElement addItem = (XmlElement)objTocRoot.SelectSingleNode("." + itemNodeLevel).AppendChild(objToc.CreateElement("item"));
                                addItem.SetAttribute("title", breadcrumbs[i]);

                                if (i == (breadcrumbs.Length - 1))
                                {
                                    /*if (selRow.Table.Columns["Column5"] != null && !String.IsNullOrEmpty(selRow["Column5"].ToString()))
                                    {
                                        //addItem.SetAttribute("href", selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + '#' + selRow["Column4"].ToString());
                                        addItem.SetAttribute("href", "./" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + ".html" + "#" + selRow["Column4"].ToString());
                                    }
                                    else
                                    {*/
                                        // get the href from current file
                                        if(urls != "" 
                                            && urls.Contains("http") == false
                                            && urls.Contains(".html") == true
                                            && urls.Contains("#") == true)
                                            
                                        {
                                            addItem.SetAttribute("href", urls.Replace(".html", "").Replace("./", ""));
                                        }
                                        else
                                        {
                                            addItem.SetAttribute("href", selRow["Column4"].ToString());
                                        }
                                   // }

                                    XmlElement breadcrumbDisplay = objToc.CreateElement("div");
                                    string breadcrumb = "";
                                    string tocId = "";

                                    foreach (XmlElement objTocItem in addItem.SelectNodes("ancestor-or-self::item"))
                                    {
                                        if (breadcrumb != "")
                                        {
                                            breadcrumb += " > ";
                                            breadcrumbDisplay.AppendChild(objToc.CreateTextNode(" > "));
                                        }
                                        breadcrumb += ((XmlElement)objTocItem).GetAttribute("title");

                                        if (objTocItem.SelectSingleNode("@href") != null)
                                        {
                                            breadcrumbDisplay.AppendChild(objToc.CreateElement("a"));
                                            string href = "./" + ((XmlElement)objTocItem).GetAttribute("href") + ".html";
                                            if(((XmlElement)objTocItem).GetAttribute("href").Contains(".html"))
                                            {
                                                href = ((XmlElement)objTocItem).GetAttribute("href");
                                            }
                                            ((XmlElement)breadcrumbDisplay.LastChild).SetAttribute("href", href);
                                            breadcrumbDisplay.LastChild.InnerText = ((XmlElement)objTocItem).GetAttribute("title");
                                        }
                                        else
                                        {
                                            breadcrumbDisplay.AppendChild(objToc.CreateTextNode(((XmlElement)objTocItem).GetAttribute("title")));
                                        }

                                    }
                                    selHtml = Regex.Replace(selHtml, @"(?<=<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">)(?:.(?!</div>))+.(?=</div>)", breadcrumbDisplay.InnerXml);
                                    selHtml = Regex.Replace(selHtml, @"(?<=<meta name=""topic-breadcrumbs"" content="")[^""]*(?="")", breadcrumb);


                                    searchWords.DocumentElement.AppendChild(searchWords.CreateElement("div"));
                                    ((System.Xml.XmlElement)searchWords.DocumentElement.LastChild).SetAttribute("id", selRow["Column4"].ToString());
                                    string bodyStr = Regex.Replace(Regex.Replace(Regex.Replace(Regex.Replace(selHtml, "\r?\n", ""), "^.+<body[^>]*>(.+?)</body>.*$", "$1", RegexOptions.Multiline), @"<div style=""text-align:right; font-size:10pt; line-height:15pt; punctuation-wrap:simple;"">.+?</div>", ""), "<.+?>", "");

                                    string searchText = bodyStr.Replace("&", "&amp;").Replace("<", "&lt;");
                                    string displayText = searchText;
                                    if (searchText.Length >= 90)
                                    {
                                        displayText = displayText.Substring(0, 90) + " ...";
                                    }

                                    string[] wide = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン" };
                                    string[] narrow = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" };

                                    for (int p = 0; p < wide.Length; p++)
                                    {
                                        searchText = Regex.Replace(searchText, wide[p], narrow[p]);
                                    }
                                    searchText = searchText.ToLower();
                                    /*if (selRow.Table.Columns["Column5"] != null && !String.IsNullOrEmpty(selRow["Column5"].ToString()))
                                    {
                                        string searchHref_5 = selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "#" + selRow["Column4"].ToString();
                                        if(searchHref_5.Contains(".html") == false)
                                        {
                                            searchHref_5 = selRow["Column4"].ToString() + ".html" + "#" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "");
                                        }
                                        *//*searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>" 
                                            + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>" 
                                            + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;") 
                                            + "</div><div class='displayText'>" + displayText 
                                            + "</div><div class='search_word'>" + searchText + "</div>";*//*
                                        string textSearchWords = "<div class='search_breadcrumbs'>";
                                        textSearchWords += breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>";

                                        // check SelectSingleNode is null or not
                                        if (objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']") != null){
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref_5 + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }
                                        else if (objToc.SelectSingleNode(".//item[@href = '" + selRow["Column4"].ToString() + "']") != null)
                                        {
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + selRow["Column4"].ToString() + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }
                                        else if (objToc.SelectSingleNode(".//item[@href = '" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "']") != null)
                                        {
                                            textSearchWords += ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + selRow["Column5"].ToString().Replace("(", "").Replace(")", "") + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;");
                                        }

                                        textSearchWords += "</div><div class='displayText'>" + displayText + "</div><div class='search_word'>" + searchText + "</div>";
                                        searchWords.DocumentElement.LastChild.InnerXml = textSearchWords;
                                    }
                                    else
                                    {*/
                                        string searchHref = selRow["Column4"].ToString();
                                        if (urls != "" && urls.Contains("http") == false && urls.Contains(".html") == true && urls.Contains("#") == true)
                                        {
                                            searchHref = urls.Replace(".html", "").Replace("./", "");
                                        }
                                        searchWords.DocumentElement.LastChild.InnerXml = "<div class='search_breadcrumbs'>" 
                                            + breadcrumb.Replace("&", "&amp;").Replace("<", "&lt;") + "</div><div class='search_title'>" 
                                            + ((XmlElement)objToc.SelectSingleNode(".//item[@href = '" + searchHref + "']")).GetAttribute("title").Replace("&", "&amp;").Replace("<", "&lt;") 
                                            + "</div><div class='displayText'>" + displayText 
                                            + "</div><div class='search_word'>" + searchText + "</div>";
                                    //}


                                }
                            }
                        }
                    }

                    MatchCollection mc = Regex.Matches(selHtml, @"(?<=<a href="")(?!\./|http)(?:[^""]*?/)+([^""]*?)(?="")", RegexOptions.Singleline);
                    foreach (Match m in mc)
                    {
                        string[] splitText = m.Groups[1].Value.Split('#');
                        // check if the file is in the list
                        if (lsfiles.Contains(splitText[0]))
                            if (m.Groups[1].Value.Contains("html") == true)
                            {
                                selHtml = selHtml.Replace(m.Value, "./" + m.Groups[1].Value);
                            }
                            else
                            {
                                selHtml = selHtml.Replace(m.Value, "./" + m.Groups[1].Value + "html");
                            }
                    }

                    //                    selHtml = Regex.Replace(selHtml, @"(?<=<a href="")(?!\./|http)(?:[^""]*?/)+([^""]*?)(?="")", "./$1");
                    sw = new StreamWriter(Path.Combine(outputDir, selRow["Column4"].ToString() + ".html"), false, Encoding.UTF8);
                    sw.Write(selHtml);
                    sw.Close();
                }
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

            if (chbListOutput.Checked)
            {
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

        

        private void tbSelectJoinList_TextChanged(object sender, EventArgs e)
        {
            if (File.Exists(tbSelectJoinList.Text))
            {
                XmlDocument list = new XmlDocument();
                try
                {
                    list.Load(tbSelectJoinList.Text);
                    if (list.SelectSingleNode(".//changeTitle") != null)
                    {
                        chbChangeTitle.Checked = true;
                        tbChangeTitle.Enabled = true;
                        tbChangeTitle.Text = list.SelectSingleNode(".//changeTitle").InnerText;
                    }
                    if (list.SelectSingleNode(".//addTopLevel") != null)
                    {
                        chbAddTop.Checked = true;
                        tbAddTop.Enabled = true;
                        tbAddTop.Text = list.SelectSingleNode(".//addTopLevel").InnerText;
                    }

                    lbHtmlList.Items.Clear();
                    bookInfo.Clear();

                    foreach (XmlNode htmlList in list.SelectNodes(".//htmlList/item"))
                    {
                        string path = ((XmlElement)htmlList).GetAttribute("src");
                        int lbListItemIndex = lbHtmlList.Items.Add(path);

                        if (Directory.Exists(Path.Combine(Path.GetDirectoryName(path), "headerFile")))
                        {
                            string[] listFile = Directory.GetFiles(Path.Combine(Path.GetDirectoryName(path), "headerFile"), "???.txt");
                            if (listFile.Length == 0)
                            {
                                MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(path), "headerFile") + "」に書誌情報ファイルが存在しません。");
                            }
                            else
                            {
                                bookInfo[path] = new System.Data.DataTable();

                                bookInfo[path].Columns.Add("Column1", typeof(bool));
                                bookInfo[path].Columns.Add("Column2", typeof(string));
                                bookInfo[path].Columns.Add("Column3", typeof(string));
                                bookInfo[path].Columns.Add("Column4", typeof(string));

                                using (StreamReader sr = new StreamReader(listFile[0]))
                                {
                                    while (!sr.EndOfStream)
                                    {
                                        string[] lineStr = (sr.ReadLine()).Split('\t');
                                        if (lineStr[2].Contains("#")) continue;
                                        bookInfo[path].Rows.Add(((htmlList.SelectSingleNode(".//checked[@id='" + lineStr[2] + "']") != null) ? true : false), lineStr[0], lineStr[1], lineStr[2]);
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(path), "headerFile") + "」フォルダが存在しません。");
                            lbHtmlList.Items.RemoveAt(lbListItemIndex);
                        }
                    }

                    if (list.SelectSingleNode(".//outputDir") != null)
                    {
                        tbOutputDir.Text = list.SelectSingleNode(".//outputDir/@src").InnerText;
                    }
                }
                catch (XmlException xmlex)
                {
                    if (Regex.IsMatch(tbSelectJoinList.Text, @"\.xml$"))
                    {
                        MessageBox.Show("結合リストが破損しています。");
                    }
                    else
                    {
                        MessageBox.Show("XMLファイルを選択してください。");
                    }
                    tbSelectJoinList.Text = "";
                }
                catch (XPathException xpathex)
                {
                    MessageBox.Show("結合リストが破損しています。");
                    tbSelectJoinList.Text = "";
                }
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            tbSelectJoinList.Text = "";
            chbChangeTitle.Checked = false;
            tbChangeTitle.Text = "";
            tbChangeTitle.Enabled = false;
            chbAddTop.Checked = false;
            tbAddTop.Text = "";
            tbAddTop.Enabled = false;
            while (lbHtmlList.Items.Count != 0)
            {
                lbHtmlList.Items.RemoveAt(0);
            }
            dataGridView1.DataSource = null;
            bookInfo.Clear();
            tbOutputDir.Text = "";
            chbListOutput.Checked = false;
        }

        private void button4_Click(object sender, EventArgs e)
        {
            openFileDialog2.FileName = "";
            openFileDialog2.Filter = "docファイル(*.doc)|*.doc|すべてのファイル(*.*)|*.*";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog2.FileName) != ".doc")
                {
                    MessageBox.Show("docファイルを選択してください。");
                    return;
                }
                listBox1.Items.Add(openFileDialog2.FileName);
            }
            checkItems();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
            checkItems();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if ((listBox1.SelectedIndex != -1) && (listBox1.SelectedIndex != 0))
            {
                object tmp = listBox1.Items[listBox1.SelectedIndex - 1];
                listBox1.Items[listBox1.SelectedIndex - 1] = listBox1.Items[listBox1.SelectedIndex];
                listBox1.Items[listBox1.SelectedIndex] = tmp;
                listBox1.SelectedIndex = listBox1.SelectedIndex - 1;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if ((listBox1.SelectedIndex != -1) && (listBox1.SelectedIndex != (listBox1.Items.Count - 1)))
            {
                object tmp = listBox1.Items[listBox1.SelectedIndex + 1];
                listBox1.Items[listBox1.SelectedIndex + 1] = listBox1.Items[listBox1.SelectedIndex];
                listBox1.Items[listBox1.SelectedIndex] = tmp;
                listBox1.SelectedIndex = listBox1.SelectedIndex + 1;
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string folder in s)
            {
                if (File.Exists(folder) && Path.GetExtension(folder) == ".doc")
                {
                    if (!folder.Contains("~$") && !listBox1.Items.Contains(folder))
                        listBox1.Items.Add(folder);
                }
                else if (Directory.Exists(folder))
                {
                    string[] fol = Directory.GetFiles(folder, "*.doc", SearchOption.AllDirectories);
                    foreach (string webhelp in fol)
                    {
                        if (!webhelp.Contains("~$") && !listBox1.Items.Contains(webhelp))
                            listBox1.Items.Add(webhelp);
                    }
                }
            }
            checkItems();
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog2.SelectedPath;
            }
        }
        private void checkItems()
        {
            if (listBox1.Items.Count > 1) checkBox4.Checked = true;
            else checkBox4.Checked = false;
            if (listBox1.Items.Count == 1) checkBox5.Checked = true;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            while (listBox1.Items.Count != 0)
            {
                listBox1.Items.RemoveAt(0);
            }
            textBox1.Text = "";
            checkBox4.Checked = false;
            checkBox5.Checked = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            if (!checkBox4.Checked && !checkBox5.Checked)
            {
                MessageBox.Show("出力するファイル種別を選択してください。");
                return;
            }
            //if(checkBox4.Checked && listBox1.Items.Count <= 1)
            //{
            //    MessageBox.Show("結合するWordファイルを複数指定してください。");
            //    return;
            //}
            //if(textBox1.Text == "")
            //{
            //    MessageBox.Show("出力先を指定してください。");
            //    return;
            //}
            if (checkBox4.Checked)
                saveFileDialog2.Title = "結合済みWordファイル保存";
            else
                saveFileDialog2.Title = "PDFファイル保存";

            saveFileDialog2.InitialDirectory = Path.GetDirectoryName(listBox1.Items[0].ToString());
            if (!checkBox4.Checked)
                saveFileDialog2.FileName = Path.GetFileName(listBox1.Items[0].ToString()).Replace(".doc", ".pdf");
            else
                saveFileDialog2.FileName = Path.GetFileName(listBox1.Items[0].ToString());

            if (checkBox4.Checked)
                saveFileDialog2.Filter = "Word ファイル|*.doc|すべてのファイル|*.*";
            else
                saveFileDialog2.Filter = "PDF ファイル|*.pdf|すべてのファイル|*.*";

            string strOutputDir = "";
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                strOutputDir = saveFileDialog2.FileName.Replace(".pdf", ".doc");
            else
                return;

            groupBox1.Visible = true;
            if (checkBox4.Checked)
                label10.Text = "Word結合中...";
            else
                label10.Text = "PDF出力中...";

            string strOrigFile = "";
            List<string> strCopiesDir = new List<string>();
            bool bl = false;
            foreach (string file in listBox1.Items)
            {
                if (!bl)
                {
                    strOrigFile = file;
                    bl = true;
                    continue;
                }
                strCopiesDir.Add(file);
            }
            //string strOutputDir = textBox1.Text + "\\" + Path.GetFileName(listBox1.Items[0].ToString());
            DocMerger objMerger = new DocMerger();
            bool cntItems = false;
            if (listBox1.Items.Count == 1) cntItems = true;
            objMerger.Merge(strOrigFile, strCopiesDir, strOutputDir, this, checkBox4.Checked, checkBox5.Checked, cntItems);
            groupBox1.Visible = false;
            if (checkBox4.Checked)
            {
                MessageBox.Show("Wordの結合が完了しました。");
                DialogResult selectMess = MessageBox.Show(strOutputDir + "\r\nが出力されました。\r\n出力したWordを表示しますか？", "Word結合成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectMess == DialogResult.Yes)
                {
                    try
                    {
                        Process.Start(strOutputDir);
                    }
                    catch
                    {
                        MessageBox.Show("Wordの出力に失敗しました。", "Word出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            if (checkBox5.Checked)
            {
                MessageBox.Show("PDFの出力が完了しました。");
                DialogResult selectMess = MessageBox.Show(strOutputDir.Replace(".doc", ".pdf") + "\r\nが出力されました。\r\n出力したPDFを表示しますか？", "PDF結合成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectMess == DialogResult.Yes)
                {
                    try
                    {
                        Process.Start(strOutputDir.Replace(".doc", ".pdf"));
                    }
                    catch
                    {
                        MessageBox.Show("PDFの出力に失敗しました。", "PDF出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void listBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void listBox2_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string folder in s)
            {
                if (File.Exists(folder)) continue;
                bool bl = false;
                foreach (string lbi in listBox2.Items)
                    if (lbi == folder) bl = true;
                if (bl == false) listBox2.Items.Add(folder);
            }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if ((listBox2.SelectedIndex != -1) && (listBox2.SelectedIndex != 0))
            {
                object tmp = listBox2.Items[listBox2.SelectedIndex - 1];
                listBox2.Items[listBox2.SelectedIndex - 1] = listBox2.Items[listBox2.SelectedIndex];
                listBox2.Items[listBox2.SelectedIndex] = tmp;
                listBox2.SelectedIndex = listBox2.SelectedIndex - 1;
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if ((listBox2.SelectedIndex != -1) && (listBox2.SelectedIndex != (listBox2.Items.Count - 1)))
            {
                object tmp = listBox2.Items[listBox2.SelectedIndex + 1];
                listBox2.Items[listBox2.SelectedIndex + 1] = listBox2.Items[listBox2.SelectedIndex];
                listBox2.Items[listBox2.SelectedIndex] = tmp;
                listBox2.SelectedIndex = listBox2.SelectedIndex + 1;
            }
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex != -1)
            {
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            button13.Enabled = false;
            foreach (string folder in listBox2.Items)
            {
                if (Directory.Exists(folder)) linkCheck(folder);
                label14.Visible = false;
                progressBar2.Visible = false;
            }
            foreach (ListViewItem lvi in listView1.Items)
                if (lvi.BackColor == Color.Red) button13.Enabled = true;
        }
        private void linkCheck(string folder)
        {
            string[] files = Directory.GetFiles(folder, "*.html", SearchOption.AllDirectories);
            progressBar2.Visible = true;
            progressBar2.Value = 0;
            progressBar2.Maximum = files.Count() - 1;
            label14.Visible = true;
            foreach (string file in files)
            {
                label14.Text = Path.GetFileName(file) + " / " + Path.GetFileName(folder);
                label14.Refresh();
                progressBar2.Increment(1);
                string allText = "";
                using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
                {
                    allText = sr.ReadToEnd();
                }
                Regex r = new Regex(@"(?<=<a href="")http[^""]*?(?="")");
                MatchCollection mc = r.Matches(allText);
                //foreach (Match m in mc)
                //{
                //    Console.WriteLine(m.Value);
                //    //Console.WriteLine(GetStatusCode(m.Value).ToString());
                //}
                Regex r2 = new Regex(@"(?<=<a href="")([^""]*?)"">([^<]*?)(?=</a>)");
                MatchCollection mc2 = r2.Matches(allText);
                //' Ver - 2023.17.08 - VyNL - ↓ - 追加'
                Regex r3 = new Regex(@"(?<=<p\sclass=""MJS_ref""><span\sname=""([^""]+)""\sclass=""ref""\s*/>([^<]*?)<\/p>)");
                MatchCollection mc3 = r3.Matches(allText);

                Regex r4 = new Regex(@"refPage\s*=\s*{([\s\S]*?)}");
                MatchCollection mcRefPage = r4.Matches(allText);
                Regex r5 = new Regex(@"mergePage\s*=\s*{([\s\S]*?)}");
                MatchCollection mcMergePage = r5.Matches(allText);

                if (mcRefPage != null && mcRefPage.Count > 0)
                {
                    var refPage = mcRefPage[0].Groups[1].Value;

                    foreach (Match m4 in mc3)
                    {
                        if (Regex.IsMatch(mcRefPage[0].Groups[1].Value, m4.Groups[1].Value.Trim().Replace("_ref", "")))
                        {

                            MatchCollection matches = Regex.Matches(refPage, $"{m4.Groups[1].Value.Trim().Replace("_ref", "")}:(.*?])");

                            foreach (Match match in matches)
                            {
                                string content = match.Groups[1].Value.Trim().Replace("[", "").Replace("]", "");

                                string titleName = "";
                                string replaceTitleName = "";
                                string linkPage = "";
                                string[] parts = content.Split(',');
                                int indexOfComma = content.IndexOf(',');

                                if (mcMergePage != null && mcMergePage.Count > 0)
                                {
                                    var mergePage = mcMergePage[0].Groups[1].Value.Trim().Split(',');
                                    foreach (var key in mergePage)
                                    {
                                        string[] link = key.Trim().Split(':');
                                        foreach (var itemLink in link)
                                        {
                                            if (itemLink.Trim().Replace("'", "") == parts[0].Trim().Replace("'", ""))
                                            {
                                                int colonIndex = key.IndexOf(':');
                                                linkPage = key.Substring(0, colonIndex).Trim();
                                            }
                                        }
                                    }
                                    if (linkPage == "")
                                    {
                                        linkPage = parts[0].Trim().Replace("'", "");
                                    }
                                }

                                if (File.Exists(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html"))
                                {
                                    using (StreamReader sr = new StreamReader(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html", Encoding.UTF8))
                                    {
                                        titleName = sr.ReadToEnd();

                                    }
                                    if (Regex.IsMatch(titleName, @"<span name=""" + m4.Groups[1].Value.Replace("_ref", "") + @""" class=""ref"" />"))
                                    {
                                        using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
                                        {
                                            string titleNameFile = srFile.ReadToEnd();
                                            replaceTitleName = Regex.Match(titleNameFile, @"<p class=""MJS_ref""><span name=""" + m4.Groups[1].Value.Trim() + @""" class=""ref"" />(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                        }
                                        titleName = replaceTitleName;
                                    }
                                    else titleName = "";
                                    if (titleName.Contains(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", "")))
                                    {
                                        ListViewItem lvi = listView1.Items.Add(file);
                                        lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
                                        lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
                                        lvi.SubItems.Add("true");
                                        lvi.SubItems.Add(titleName);
                                        lvi.SubItems.Add("true");
                                    }

                                }
                                else
                                {
                                    ListViewItem lvi = listView1.Items.Add(file);
                                    lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
                                    lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
                                    lvi.SubItems.Add("false");
                                    lvi.SubItems.Add("none");
                                    lvi.SubItems.Add("false");
                                    lvi.BackColor = Color.Red;
                                }

                            }
                            logen.Clear();
                            foreach (ListViewItem lvi in listView1.Items)
                                logen.Add(lvi);
                        }

                    }

                }
                //' Ver - 2023.16.08 - VyNL - ↑ - 追加'
                foreach (Match m in mc2)
                {
                    string targetURL = "";
                    string anchor = "";
                    string replaceTitleName = "";
                  
                    if (Regex.IsMatch(m.Groups[1].Value, @"^#"))
                    {
                        targetURL = file;
                        anchor = m.Groups[1].Value.Replace("#", "");
                    }
                    else if (m.Groups[1].Value.Contains("#"))
                    {
                        // check link
                        string[] parts = m.Groups[1].Value.Split('#');
                        if (parts.Length >= 2 && parts[0].Contains(".html") == false)
                        {
                            // targetURL need concat with extension .html => pass check with extension
                            targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + parts[0] + ".html";
                        }
                        else
                        {
                            targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value.Split('#')[0];
                            
                        }
                        anchor = m.Groups[1].Value.Split('#')[1];
                        // check anchor have extension .html
                        if (anchor.Contains(".html"))
                        {
                            anchor = anchor.Replace(".html", "");
                        }
                    }
                    else
                    {
                        targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value;
                    }
                    string titleName = "";
                    if (Regex.IsMatch(m.Groups[1].Value, "^http", RegexOptions.Singleline | RegexOptions.IgnoreCase))
                    {
                        ListViewItem lvi = listView1.Items.Add(file);
                        lvi.SubItems.Add(m.Groups[1].Value);
                        try
                        {
                            lvi.SubItems.Add(GetStatusCode(m.Groups[1].Value).ToString());
                        }
                        catch
                        {
                            lvi.SubItems.Add("取得に失敗しました。");
                        }
                        lvi.SubItems.Add("");
                        lvi.SubItems.Add("");
                        lvi.SubItems.Add("");
                        lvi.BackColor = Color.Red;
                        continue;
                    }
                    if (File.Exists(targetURL))
                    {
                        using (StreamReader sr = new StreamReader(targetURL, Encoding.UTF8))
                        {
                            titleName = sr.ReadToEnd();
                            if (String.IsNullOrEmpty(anchor))
                            {
                                var head = new Regex(@"(?<=<title>)(.*?)(?=</title>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                                var h = head.Match(titleName);
                                if (h.Success)
                                {
                                    titleName = h.Groups[1].Value;
                                }
                                else titleName = "";
                            }
                            else
                            {

                                if (titleName.Contains(@"<p class=""Heading3a"" id=""" + anchor + @""">"))
                                {
                                    titleName = Regex.Match(titleName, @"<p class=""Heading3a"" id=""" + anchor + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                }
                                else if (titleName.Contains(@"<p class=""Heading3"" id=""" + anchor + @""">"))
                                {

                                    titleName = Regex.Match(titleName, @"<p class=""Heading3"" id=""" + anchor + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                    if (Regex.IsMatch(titleName, @"<span\s+name=""([^""]*)""\s+class=""ref""\s*/>") && titleName.Contains(m.Groups[2].Value))
                                    {
                                        titleName = m.Groups[2].Value;
                                    }

                                }
                                else if (titleName.Contains(@"<p class=""Heading3"" id=""" + m.Groups[1].Value.Replace(".html#", "＃") + @""">"))
                                {

                                    titleName = Regex.Match(titleName, @"<p class=""Heading3"" id=""" + m.Groups[1].Value.Replace(".html#", "＃") + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();

                                }
                                //' Ver - 2023.16.08 - VyNL - ↓ - 追加'
                                else if (File.Exists(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{anchor}.html"))
                                {
                                    using (StreamReader srAnchor = new StreamReader(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{anchor}.html", Encoding.UTF8))
                                    {
                                        titleName = srAnchor.ReadToEnd();
                                        var head = new Regex(@"(?<=<title>)(.*?)(?=</title>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                                        var h = head.Match(titleName);
                                        if (h.Success)
                                        {
                                            titleName = h.Groups[1].Value;
                                        }
                                        else titleName = "";
                                    }
                                }

                                //' Ver - 2023.16.08 - VyNL - ↑ - 追加'
                                else titleName = "";
                            }
                        }
                        // using regex clear tag <span> </span> in titleName
                        titleName = Regex.Replace(titleName, @"<span[^>]*?>", "");
                        titleName = Regex.Replace(titleName, @"</span>", "");
                        // check title same with title in link
                        if (titleName != m.Groups[2].Value)
                        {
                            // if not same add to listview
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255,192,203);//#ffc0cb pink
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else if (File.Exists(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", "")))
                    {
                        using (StreamReader sr = new StreamReader(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", ""), Encoding.UTF8))
                        {
                            titleName = sr.ReadToEnd();
                            var head = new Regex(@"(?<=class=""Heading3a"" id=""" + Regex.Replace(m.Groups[1].Value, @".*?#", "") + @""">)(.*?)(?=</p>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                            var h = head.Match(titleName);
                            if (h.Success)
                            {
                                titleName = h.Groups[1].Value;
                            }
                            else titleName = "";
                        }
                        if (titleName != m.Groups[2].Value)
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255, 192, 203);//#ffc0cb pink
                            //allCheck += file + "," + m.Groups[1].Value + "," + m.Groups[2].Value + ",false," + titleName + ",true" + "\r\n";
                            //Console.WriteLine(file + "\r\n" + m.Groups[1].Value + "\r\n" + m.Groups[2].Value + "\r\nfalse\r\n" + titleName + "\r\ntrue" + "\r\n");
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else if (m.Groups[1].Value.StartsWith("_Ref"))
                    {
                        using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
                        {
                            string titleNameFile = srFile.ReadToEnd();

                            replaceTitleName = Regex.Match(titleNameFile, @"<a\s+href=""" + m.Groups[1].Value + @""">([^<]+)<\/a>").Groups[1].Value.Trim();
                        }
                        titleName = replaceTitleName;

                        if (titleName != m.Groups[2].Value)
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255, 192, 203);//#ffc0cb pink
                            //allCheck += file + "," + m.Groups[1].Value + "," + m.Groups[2].Value + ",false," + titleName + ",true" + "\r\n";
                            //Console.WriteLine(file + "\r\n" + m.Groups[1].Value + "\r\n" + m.Groups[2].Value + "\r\nfalse\r\n" + titleName + "\r\ntrue" + "\r\n");
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else
                    {
                        ListViewItem lvi = listView1.Items.Add(file);
                        lvi.SubItems.Add(m.Groups[1].Value);
                        lvi.SubItems.Add(m.Groups[2].Value);
                        lvi.SubItems.Add("false");
                        lvi.SubItems.Add("none");
                        lvi.SubItems.Add("false");
                        lvi.BackColor = Color.Red;
                    }
                }
                logen.Clear();
                foreach (ListViewItem lvi in listView1.Items)
                    logen.Add(lvi);
            }
            //using (StreamWriter sw = new StreamWriter("./kekka.csv", false, Encoding.UTF8))
            //{
            //    sw.Write(allCheck);
            //}
            //Assembly myAssembly = Assembly.GetEntryAssembly();
            //string path = Path.GetDirectoryName(myAssembly.Location) + "/";
            //Process.Start(path + "./kekka.csv");
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1) this.Width = 800;
            else this.Width = 439;
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            //if (Regex.IsMatch(Path.GetFileName(listView1.SelectedItems[0].Text), @"^\w{3}\d{5}.html$"))
            //{
            //    System.Diagnostics.Process.Start(Path.GetDirectoryName(listView1.SelectedItems[0].Text) + @"\index.html", "#t=" + Path.GetFileName(listView1.SelectedItems[0].Text));
            //}
            System.Diagnostics.Process.Start(listView1.SelectedItems[0].Text);
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if ((folderBrowserDialog3.ShowDialog() == DialogResult.OK) && !listBox2.Items.Contains(folderBrowserDialog3.SelectedPath))
            {
                listBox2.Items.Add(folderBrowserDialog3.SelectedPath);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                List<ListViewItem> lo = new List<ListViewItem>();
                foreach (ListViewItem lvi in listView1.Items)
                    if (lvi.BackColor == Color.Red || lvi.BackColor == Color.FromArgb(255, 192, 203))
                        lo.Add(lvi);
                listView1.Items.Clear();
                foreach (ListViewItem o in lo)
                    listView1.Items.Add(o);
            }
            else
            {
                if (logen.Count > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem lvi in logen)
                        listView1.Items.Add(lvi);
                }
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                ListViewToCSV(listView1, saveFileDialog1.FileName, true);
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked) textBox2.Text = "webHelp";
            if (checkBox3.Checked) textBox2.Enabled = true;
            else textBox2.Enabled = false;
        }
    }
}
