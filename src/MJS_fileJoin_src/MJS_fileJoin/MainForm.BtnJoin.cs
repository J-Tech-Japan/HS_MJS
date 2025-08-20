// MainForm.BtnJoin.cs

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
            string searchJsPath = Path.Combine(tbOutputDir.Text, exportDir, "search.js");
            sw = new StreamWriter(searchJsPath, false, Encoding.UTF8);
            // sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2"), @"(?<=>)([^<]*?)'([^<]*?)(?=<)", "$1&apos;$2")));
            sw.Write(Regex.Replace(searchJs, "♪", Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline).Replace("'", "&apos;")));
            sw.Close();

            // search.jsファイルの内容を変数に格納
            //var mergedSearchTitles = ExtractSearchTitlesFromFile(searchJsPath);
            //var sourceSearchTitles = ExtractSourceSearchTitles();
            
            // mergedSearchTitlesに含まれているが、sourceSearchTitlesには含まれていないタイトルのリストを作成
            //var newTitles = mergedSearchTitles.Except(sourceSearchTitles).ToList();
            
            // 必要に応じて抽出されたタイトルを使用
            // mergedSearchTitles: 生成されたsearch.jsファイルの内容
            // sourceSearchTitles: 結合元フォルダのsearch.jsファイルの内容を合わせたもの
            // newTitles: 統合時に新たに追加されたタイトル

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

            // 統合時に新たに追加されたタイトルのデバッグ表示（必要に応じて）
            //DisplayNewTitles(newTitles);
        }

        /// <summary>
        /// 統合時に新たに追加されたタイトルを表示する（デバッグ用）
        /// </summary>
        /// <param name="newTitles">新たに追加されたタイトルのリスト</param>
        private void DisplayNewTitles(List<string> newTitles)
        {
            if (newTitles.Count == 0)
            {
                MessageBox.Show("統合時に新たに追加されたタイトルはありません。", "新規タイトル", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("=== 統合時に新たに追加されたタイトル ===");
            sb.AppendLine();
            sb.AppendLine($"【新規追加タイトル数】: {newTitles.Count}件");
            sb.AppendLine();

            for (int i = 0; i < newTitles.Count; i++)
            {
                sb.AppendLine($"【{i + 1:D3}】 {newTitles[i]}");
            }

            MessageBox.Show(sb.ToString(), "新規タイトル", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        /// <summary>
        /// 結合元の各フォルダのsearch.jsファイルから<div class="search_title">の内容を抽出して合わせたリストを返す
        /// </summary>
        /// <returns>結合元フォルダのsearch_titleの内容を合わせたリスト</returns>
        private List<string> ExtractSourceSearchTitles()
        {
            var allSourceTitles = new List<string>();

            try
            {
                // 結合元の各フォルダのsearch.jsファイルの内容を抽出
                foreach (string htmlDir in lbHtmlList.Items)
                {
                    string sourceSearchJsPath = Path.Combine(htmlDir, "search.js");
                    var sourceTitles = ExtractSearchTitlesFromFile(sourceSearchJsPath);
                    allSourceTitles.AddRange(sourceTitles);
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合はログに記録（必要に応じて）
                // 現在は何もしない
            }

            return allSourceTitles;
        }

        /// <summary>
        /// 指定されたsearch.jsファイルから<div class="search_title">の内容を抽出してリストで返す
        /// </summary>
        /// <param name="searchJsFilePath">search.jsファイルのパス</param>
        /// <returns>抽出されたsearch_titleの内容のリスト</returns>
        private List<string> ExtractSearchTitlesFromFile(string searchJsFilePath)
        {
            var titles = new List<string>();

            try
            {
                if (!File.Exists(searchJsFilePath))
                {
                    return titles;
                }

                string fileContent = File.ReadAllText(searchJsFilePath, Encoding.UTF8);
                
                // search.jsファイルからsearchWordsの変数内容を抽出
                var searchWordsMatch = Regex.Match(fileContent, @"var searchWords = \$\('(.*?)'\);", RegexOptions.Singleline);
                if (!searchWordsMatch.Success)
                {
                    return titles;
                }

                string searchWordsContent = searchWordsMatch.Groups[1].Value;
                
                // <div class="search_title">の内容をすべて抽出
                var titleMatches = Regex.Matches(searchWordsContent, @"<div\s+class=['""]search_title['""]>(.*?)</div>", RegexOptions.Singleline);
                
                foreach (Match match in titleMatches)
                {
                    string titleContent = match.Groups[1].Value;
                    // HTMLエンティティをデコード
                    titleContent = titleContent.Replace("&amp;", "&")
                                             .Replace("&lt;", "<")
                                             .Replace("&gt;", ">")
                                             .Replace("&quot;", "\"")
                                             .Replace("&apos;", "'");
                    
                    titles.Add(titleContent);
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合は空のリストを返す
            }

            return titles;
        }
    }
}
