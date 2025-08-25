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
using Microsoft.Office.Core;

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

            // 全結合元フォルダのheadingsWithCommentを統合するリスト
            List<string> allHeadingsWithComment = new List<string>();

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

                // headingsWithComment.txtの存在確認と読み込み
                string headingsWithCommentPath = Path.Combine(htmlDir, "headingsWithComment.txt");
                
                if (File.Exists(headingsWithCommentPath))
                {
                    try
                    {
                        var loadedHeadings = Utils.ReadLinesFromFile(headingsWithCommentPath);
                        allHeadingsWithComment.AddRange(loadedHeadings);
                    }
                    catch (Exception ex)
                    {
                        // エラーハンドリング（必要に応じてログ出力やメッセージ表示）
                        errorList.Add($"headingsWithComment.txtの読み込みエラー ({htmlDir}): {ex.Message}");
                    }
                }

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

            // 結合したheadingsWithCommentの各タイトルに対して
            // RemoveSearchBlockByTitleを実行して検索対象から除外する
            foreach (string heading in allHeadingsWithComment)
            {
                RemoveSearchBlockByTitle(heading, tbOutputDir.Text, exportDir);
            }

            // 目次アイテムごとのHTMLファイルを処理し、gTopicIdを書き換えて保存
            UpdateHtmlFilesWithTocId(objToc, tbOutputDir.Text, exportDir);

            //目次出力
            CreateToc(objToc.DocumentElement);

            // chbListOutputがチェックされている場合にjoinList.xmlを出力する
            OutputJoinListXml();

            //書誌情報ファイルのマージ
            MergeHeaderFile();

            // 統合されたheadingsWithCommentをファイルに出力（必要に応じて）
            if (allHeadingsWithComment.Count > 0)
            {
                try
                {
                    string outputHeadingsPath = Path.Combine(tbOutputDir.Text, exportDir, "headingsWithComment.txt");
                    Utils.WriteLinesToFile(Path.Combine(tbOutputDir.Text, exportDir), "headingsWithComment.txt", allHeadingsWithComment);
                }
                catch (Exception ex)
                {
                    errorList.Add($"統合されたheadingsWithComment.txtの出力エラー: {ex.Message}");
                }
            }

            Cursor.Current = prevCursor;

            AfterHtmlOutput(Path.Combine(tbOutputDir.Text, exportDir));

            // 統合時に新たに追加されたタイトルのデバッグ表示（必要に応じて）
            //DisplayNewTitles(newTitles);
        }

        /// <summary>
        /// mergedSearchTitlesとsourceSearchTitlesの内容をそれぞれテキストファイルに書き込む
        /// </summary>
        /// <param name="mergedSearchTitles">統合後のsearch.jsから抽出されたタイトルリスト</param>
        /// <param name="sourceSearchTitles">結合元フォルダのsearch.jsから抽出されたタイトルリスト</param>
        /// <param name="outputBaseDir">出力ベースディレクトリ</param>
        /// <param name="exportDir">エクスポートディレクトリ名</param>
        private void WriteSearchTitlesToFiles(List<string> mergedSearchTitles, List<string> sourceSearchTitles, string outputBaseDir, string exportDir)
        {
            try
            {
                string outputDir = Path.Combine(outputBaseDir, exportDir);
                
                // mergedSearchTitlesの内容をファイルに出力
                string mergedTitlesFilePath = Path.Combine(outputDir, "merged_search_titles.txt");
                using (var writer = new StreamWriter(mergedTitlesFilePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("=== 統合後のsearch.jsから抽出されたタイトル ===");
                    writer.WriteLine($"総件数: {mergedSearchTitles.Count}件");
                    writer.WriteLine();
                    
                    for (int i = 0; i < mergedSearchTitles.Count; i++)
                    {
                        writer.WriteLine($"【{i + 1:D4}】 {mergedSearchTitles[i]}");
                    }
                }

                // sourceSearchTitlesの内容をファイルに出力
                string sourceTitlesFilePath = Path.Combine(outputDir, "source_search_titles.txt");
                using (var writer = new StreamWriter(sourceTitlesFilePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("=== 結合元フォルダのsearch.jsから抽出されたタイトル ===");
                    writer.WriteLine($"総件数: {sourceSearchTitles.Count}件");
                    writer.WriteLine();
                    
                    for (int i = 0; i < sourceSearchTitles.Count; i++)
                    {
                        writer.WriteLine($"【{i + 1:D4}】 {sourceSearchTitles[i]}");
                    }
                }

                // 新規追加されたタイトルも同時に出力
                var newTitles = mergedSearchTitles.Except(sourceSearchTitles).ToList();
                string newTitlesFilePath = Path.Combine(outputDir, "new_search_titles.txt");
                using (var writer = new StreamWriter(newTitlesFilePath, false, Encoding.UTF8))
                {
                    writer.WriteLine("=== 統合時に新たに追加されたタイトル ===");
                    writer.WriteLine($"新規追加件数: {newTitles.Count}件");
                    writer.WriteLine();
                    
                    for (int i = 0; i < newTitles.Count; i++)
                    {
                        writer.WriteLine($"【{i + 1:D3}】 {newTitles[i]}");
                    }
                }
            }
            catch (Exception ex)
            {
                // エラーが発生した場合はメッセージボックスで通知
                MessageBox.Show($"検索タイトルファイルの出力中にエラーが発生しました: {ex.Message}", 
                               "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
