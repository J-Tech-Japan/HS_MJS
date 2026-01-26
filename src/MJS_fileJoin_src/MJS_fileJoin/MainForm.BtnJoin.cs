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

            // 結合元のsearch.jsに"キャッシュ管理関数"が含まれているかチェック
            bool hasCacheManagement = HasCacheManagementFunction(htmlDirs);

            // "キャッシュ管理関数"が含まれていない場合、全角→半角変換を実行
            if (!hasCacheManagement)
            {
                System.Diagnostics.Trace.WriteLine("search.jsに\"キャッシュ管理関数\"が見つかりませんでした。全角→半角変換を実行します。");
                ApplyWideToNarrowConversion(searchWords);
            }
            else
            {
                System.Diagnostics.Trace.WriteLine("search.jsに\"キャッシュ管理関数\"が見つかりました。全角→半角変換をスキップします。");
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

        // 全角→半角変換を実行するメソッド
        private void ApplyWideToNarrowConversion(XmlDocument searchWords)
        {
            string[] wide = { "０", "１", "２", "３", "４", "５", "６", "７", "８", "９", "Ａ", "Ｂ", "Ｃ", "Ｄ", "Ｅ", "Ｆ", "Ｇ", "Ｈ", "Ｉ", "Ｊ", "Ｋ", "Ｌ", "Ｍ", "Ｎ", "Ｏ", "Ｐ", "Ｑ", "Ｒ", "Ｓ", "Ｔ", "Ｕ", "Ｖ", "Ｗ", "Ｘ", "Ｙ", "Ｚ", "ａ", "ｂ", "ｃ", "ｄ", "ｅ", "ｆ", "ｇ", "ｈ", "ｉ", "ｊ", "ｋ", "ｌ", "ｍ", "ｎ", "ｏ", "ｐ", "ｑ", "ｒ", "ｓ", "ｔ", "ｕ", "ｖ", "ｗ", "ｘ", "ｙ", "ｚ", "ガ", "ギ", "グ", "ゲ", "ゴ", "ザ", "ジ", "ズ", "ゼ", "ゾ", "ダ", "ヂ", "ヅ", "デ", "ド", "バ", "ビ", "ブ", "ベ", "ボ", "パ", "ピ", "プ", "ペ", "ポ", "。", "「", "」", "、", "ヲ", "ァ", "ィ", "ゥ", "ェ", "ォ", "ャ", "ュ", "ョ", "ッ", "ー", "ア", "イ", "ウ", "エ", "オ", "カ", "キ", "ク", "ケ", "コ", "サ", "シ", "ス", "セ", "ソ", "タ", "チ", "ツ", "テ", "ト", "ナ", "ニ", "ヌ", "ネ", "ノ", "ハ", "ヒ", "フ", "ヘ", "ホ", "マ", "ミ", "ム", "メ", "モ", "ヤ", "ユ", "ヨ", "ラ", "リ", "ル", "レ", "ロ", "ワ", "ン" };
            string[] narrow = { "0", "1", "2", "3", "4", "5", "6", "7", "8", "9", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "a", "b", "c", "d", "e", "f", "g", "h", "i", "j", "k", "l", "m", "n", "o", "p", "q", "r", "s", "t", "u", "v", "w", "x", "y", "z", "ｶﾞ", "ｷﾞ", "ｸﾞ", "ｹﾞ", "ｺﾞ", "ｻﾞ", "ｼﾞ", "ｽﾞ", "ｾﾞ", "ｿﾞ", "ﾀﾞ", "ﾁﾞ", "ﾂﾞ", "ﾃﾞ", "ﾄﾞ", "ﾊﾞ", "ﾋﾞ", "ﾌﾞ", "ﾍﾞ", "ﾎﾞ", "ﾊﾟ", "ﾋﾟ", "ﾌﾟ", "ﾍﾟ", "ﾎﾟ", "｡", "｢", "｣", "､", "ｦ", "ｧ", "ｨ", "ｩ", "ｪ", "ｫ", "ｬ", "ｭ", "ｮ", "ｯ", "ｰ", "ｱ", "ｲ", "ｳ", "ｴ", "ｵ", "ｶ", "ｷ", "ｸ", "ｹ", "ｺ", "ｻ", "ｼ", "ｽ", "ｾ", "ｿ", "ﾀ", "ﾁ", "ﾂ", "ﾃ", "ﾄ", "ﾅ", "ﾆ", "ﾇ", "ﾈ", "ﾉ", "ﾊ", "ﾋ", "ﾌ", "ﾍ", "ﾎ", "ﾏ", "ﾐ", "ﾑ", "ﾒ", "ﾓ", "ﾔ", "ﾕ", "ﾖ", "ﾗ", "ﾘ", "ﾙ", "ﾚ", "ﾛ", "ﾜ", "ﾝ" };

            // searchWords内のすべてのsearch_wordノードを処理
            XmlNodeList searchWordNodes = searchWords.SelectNodes("//div[@class='search_word']");
            foreach (XmlNode node in searchWordNodes)
            {
                string searchText = node.InnerText;
                
                // 全角→半角変換
                for (int p = 0; p < wide.Length; p++)
                {
                    searchText = searchText.Replace(wide[p], narrow[p]);
                }
                
                // 小文字化
                searchText = searchText.ToLower();
                
                // 変換後のテキストを設定
                node.InnerText = searchText;
            }
        }
    }
}
