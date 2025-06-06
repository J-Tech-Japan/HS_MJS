using System;
using System.Collections.Generic;
using System.Data;
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

            // 入力バリデーション
            if (!ValidateInput())
            {
                Cursor.Current = prevCursor;
                return;
            }

            StreamWriter sw = null;

            List<string> errorList = new List<string>();

            // 出力ディレクトリの準備
            // exportDir変数に新しいフォルダ名が格納される
            PrepareOutputDirectory();

            XmlDocument objToc = new XmlDocument();
            XmlNode objTocRoot = null;

            XmlDocument searchWords = new System.Xml.XmlDocument();
            searchWords.LoadXml("<div class='search'></div>");

            objToc.LoadXml(@"<result></result>");
            objTocRoot = objToc.DocumentElement;

            // HTMLファイルリストの作成
            List<string> lsfiles = CreateHtmlFileList();

            int picCount = 0;

            foreach (string htmlDir in lbHtmlList.Items)
            {
                picCount++;
                var pictDir = Path.Combine(htmlDir, "pict");
                List<string> pics = Directory.Exists(pictDir)
                    ? Directory.GetFiles(pictDir, "*.*", SearchOption.AllDirectories)
                        .Select(Path.GetFileName)
                        .ToList()
                    : new List<string>();

                string outputDir = Path.Combine(tbOutputDir.Text, exportDir);

                // インデックスページ準備
                objTocRoot = PrepareIndexPage(htmlDir, outputDir, objTocRoot, objToc, tbChangeTitle, tbAddTop);

                // HTMLファイルのコピーと加工処理
                ProcessHtmlFiles(htmlDir, outputDir, picCount, lsfiles, objTocRoot, objToc, searchWords, errorList);
            }

            //全文検索ファイル出力
            sw = new StreamWriter(Path.Combine(tbOutputDir.Text, exportDir, "search.js"), false, Encoding.UTF8);
            // sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2"), @"(?<=>)([^<]*?)'([^<]*?)(?=<)", "$1&apos;$2")));
            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;")));
            sw.Close();

            // 目次アイテムごとのHTMLファイルを処理し、gTopicIdを書き換えて保存
            UpdateHtmlFilesWithTocId(objToc, tbOutputDir.Text, exportDir);

            //目次出力
            CreateToc(objToc.DocumentElement);

            // chbListOutputがチェックされている場合にjoinList.xmlを出力する
            OutputJoinListXml();

            //書誌情報ファイルのマージ
            MergeHeaderFile();

            Cursor.Current = prevCursor;

            AfterHtmlOutput(Path.Combine(tbOutputDir.Text, exportDir));
        }
    }
}
