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

            // 目次アイテムごとのHTMLファイルを処理し、gTopicIdを書き換えて保存
            UpdateHtmlFilesWithTocId(objToc, tbOutputDir.Text, exportDir);

            //目次出力
            CreateToc(objToc.DocumentElement);

            // chbListOutputがチェックされている場合にjoinList.xmlを出力する
            OutputJoinListXml();

            //書誌情報ファイルのマージ
            mergeHeaderFile();

            Cursor.Current = prevCursor;

            AfterHtmlOutput(Path.Combine(tbOutputDir.Text, exportDir));
        }
    }
}
