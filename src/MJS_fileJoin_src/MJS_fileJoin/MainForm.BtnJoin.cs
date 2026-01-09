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

            // 全結合元フォルダのheadingsWithCommentを統合するリスト
            List<string> allHeadingsWithComment = new List<string>();

            // 全結合元フォルダのheadingsを統合するリスト
            List<string> allHeadings = new List<string>();

            // 結合元フォルダのリストを取得（search.js読み込み用）
            List<string> htmlDirs = lbHtmlList.Items.Cast<string>().ToList();

            int picCount = 0;

            foreach (string htmlDir in htmlDirs)
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

                // headings.txtの存在確認と読み込み
                string headingsPath = Path.Combine(htmlDir, "headings.txt");
                
                if (File.Exists(headingsPath))
                {
                    try
                    {
                        var loadedHeadings = Utils.ReadLinesFromFile(headingsPath);
                        allHeadings.AddRange(loadedHeadings);
                    }
                    catch (Exception ex)
                    {
                        // エラーハンドリング（必要に応じてログ出力やメッセージ表示）
                        errorList.Add($"headings.txtの読み込みエラー ({htmlDir}): {ex.Message}");
                    }
                }

                // インデックスページ準備
                objTocRoot = PrepareIndexPage(htmlDir, outputDir, objTocRoot, objToc, tbChangeTitle, tbAddTop);

                // HTMLファイルのコピーと加工処理
                ProcessHtmlFiles(htmlDir, outputDir, picCount, lsfiles, objTocRoot, objToc, searchWords, errorList);
            }

            //全文検索ファイル出力
            // searchWords XMLを処理（エスケープ処理）
            string processedSearchWordsXml = Regex.Replace(searchWords.OuterXml, @"(?<=>)([^<]*?)""([^<]*?)(?=<)", "$1&quot;$2", RegexOptions.Singleline)
                .Replace("'", "&apos;");
            
            // 結合元フォルダのsearch.jsを読み込み、searchWordsを書き換え
            string searchJsContent = GetSearchJsWithReplacedWords(htmlDirs, processedSearchWordsXml);
            
            string searchJsPath = Path.Combine(tbOutputDir.Text, exportDir, "search.js");
            sw = new StreamWriter(searchJsPath, false, Encoding.UTF8);
            sw.Write(searchJsContent);
            sw.Close();

            // 結合したheadingsWithCommentの各タイトルに対して
            // RemoveSearchBlockByTitleを実行して検索対象から除外する
            foreach (string heading in allHeadingsWithComment)
            {
                RemoveSearchBlockByTitle(heading, tbOutputDir.Text, exportDir);
            }

            // allHeadingsWithCommentでheadingsWithComment.txtを上書きする
            string exportDirPath = Path.Combine(tbOutputDir.Text, exportDir);
            if (allHeadingsWithComment.Count > 0)
            {
                try
                {
                    Utils.WriteLinesToFile(exportDirPath, "headingsWithComment.txt", allHeadingsWithComment);
                }
                catch (Exception ex)
                {
                    errorList.Add($"headingsWithComment.txtの書き込みエラー: {ex.Message}");
                }
            }

            // 結合したheadingsの各タイトルに対して
            // RemoveSearchBlockByTitleを実行して検索対象から除外する
            foreach (string heading in allHeadings)
            {
                RemoveSearchBlockByTitle(heading, tbOutputDir.Text, exportDir);
            }

            // allHeadingsでheadingsを上書きする
            if (allHeadings.Count > 0)
            {
                try
                {
                    Utils.WriteLinesToFile(exportDirPath, "headings.txt", allHeadings);
                }
                catch (Exception ex)
                {
                    errorList.Add($"headings.txtの書き込みエラー: {ex.Message}");
                }
            }

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
