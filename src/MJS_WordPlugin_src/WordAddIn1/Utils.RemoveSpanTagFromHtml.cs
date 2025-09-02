// Utils.RemoveSpanTagFromHtml.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 指定フォルダ内のすべてのHTMLファイルからシンプルな&lt;span&gt;タグを削除する
        /// ただし、name、class、id、style等の属性があるspanタグは対象外
        /// </summary>
        /// <param name="folderPath">HTMLファイルを含むフォルダのパス</param>
        /// <param name="excludeFileNames">除外するファイル名のリスト（拡張子含む）。nullの場合は全ファイルを処理</param>
        /// <returns>処理したファイル数。エラーの場合は-1</returns>
        public static int RemoveSimpleSpanTagsFromHtmlFolder(string folderPath, string[] excludeFileNames = null)
        {
            if (string.IsNullOrEmpty(folderPath))
            {
                System.Diagnostics.Debug.WriteLine("フォルダパスが指定されていません。");
                return -1;
            }

            if (!Directory.Exists(folderPath))
            {
                System.Diagnostics.Debug.WriteLine($"フォルダが存在しません: {folderPath}");
                return -1;
            }

            try
            {
                // HTMLファイルを取得
                string[] htmlFiles = Directory.GetFiles(folderPath, "*.html", SearchOption.AllDirectories);

                if (htmlFiles.Length == 0)
                {
                    System.Diagnostics.Debug.WriteLine($"HTMLファイルが見つかりません: {folderPath}");
                    return 0;
                }

                // 除外ファイル名のセットを作成（大文字小文字を区別しない）
                var excludeFileNamesSet = excludeFileNames != null 
                    ? new HashSet<string>(excludeFileNames, StringComparer.OrdinalIgnoreCase) 
                    : new HashSet<string>();

                int processedCount = 0;
                int totalRemovedTags = 0;
                int excludedCount = 0;

                foreach (string htmlFile in htmlFiles)
                {
                    string fileName = Path.GetFileName(htmlFile);
                    
                    // 除外対象ファイルかチェック
                    if (excludeFileNamesSet.Contains(fileName))
                    {
                        excludedCount++;
                        System.Diagnostics.Debug.WriteLine($"除外対象ファイルをスキップしました: {fileName}");
                        continue;
                    }

                    int removedTags = RemoveSimpleSpanTagsFromHtmlFile(htmlFile);
                    if (removedTags >= 0)
                    {
                        processedCount++;
                        totalRemovedTags += removedTags;
                    }
                }

                System.Diagnostics.Debug.WriteLine($"処理完了: {processedCount}個のHTMLファイル処理, {excludedCount}個のファイル除外, {totalRemovedTags}個のspanタグを削除しました。");
                return processedCount;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"フォルダ処理中にエラーが発生しました: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 単一のHTMLファイルからシンプルな&lt;span&gt;タグを削除する
        /// </summary>
        /// <param name="htmlFilePath">HTMLファイルのパス</param>
        /// <returns>削除されたspanタグの数。エラーの場合は-1</returns>
        public static int RemoveSimpleSpanTagsFromHtmlFile(string htmlFilePath)
        {
            if (string.IsNullOrEmpty(htmlFilePath))
            {
                System.Diagnostics.Debug.WriteLine("HTMLファイルパスが指定されていません。");
                return -1;
            }

            if (!File.Exists(htmlFilePath))
            {
                System.Diagnostics.Debug.WriteLine($"HTMLファイルが存在しません: {htmlFilePath}");
                return -1;
            }

            try
            {
                // ファイル内容を読み込み
                string content;
                using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
                {
                    content = reader.ReadToEnd();
                }

                // 削除前のspanタグ数をカウント
                int originalSpanCount = CountSimpleSpanTags(content);

                // シンプルなspanタグを削除
                string cleanedContent = RemoveSimpleSpanTags(content);

                // 削除後のspanタグ数をカウント
                int remainingSpanCount = CountSimpleSpanTags(cleanedContent);
                int removedCount = originalSpanCount - remainingSpanCount;

                // ファイルに書き戻し
                if (removedCount > 0)
                {
                    using (var writer = new StreamWriter(htmlFilePath, false, Encoding.UTF8))
                    {
                        writer.Write(cleanedContent);
                    }
                    System.Diagnostics.Debug.WriteLine($"{Path.GetFileName(htmlFilePath)}: {removedCount}個のシンプルなspanタグを削除しました。");
                }

                return removedCount;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"HTMLファイル処理中にエラーが発生しました: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 文字列からシンプルな&lt;span&gt;タグを削除する
        /// 属性のないspanタグのみを対象とし、中身は保持する
        /// </summary>
        /// <param name="input">処理対象の文字列</param>
        /// <returns>クリーニング済みの文字列</returns>
        private static string RemoveSimpleSpanTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // シンプルな<span>タグ（属性なし）をマッチする正規表現
            // 開始タグ: <span>（空白や属性なし）
            // 終了タグ: </span>
            // 中身は保持する
            string pattern = @"<span\s*>(.*?)</span>";

            // タグを削除し、中身のみを残す
            string result = Regex.Replace(input, pattern, "$1", RegexOptions.IgnoreCase | RegexOptions.Singleline);

            return result;
        }

        /// <summary>
        /// 文字列内のシンプルな&lt;span&gt;タグの数をカウントする
        /// </summary>
        /// <param name="input">カウント対象の文字列</param>
        /// <returns>見つかったシンプルなspanタグの数</returns>
        private static int CountSimpleSpanTags(string input)
        {
            if (string.IsNullOrEmpty(input))
                return 0;

            string pattern = @"<span\s*>";
            var matches = Regex.Matches(input, pattern, RegexOptions.IgnoreCase);
            return matches.Count;
        }

        /// <summary>
        /// HTMLファイル内のシンプルなspanタグの統計情報を取得する
        /// </summary>
        /// <param name="htmlFilePath">HTMLファイルのパス</param>
        /// <returns>統計情報の文字列</returns>
        public static string GetHtmlSpanTagStatistics(string htmlFilePath)
        {
            if (string.IsNullOrEmpty(htmlFilePath) || !File.Exists(htmlFilePath))
                return "HTMLファイルが見つかりません。";

            try
            {
                string content;
                using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
                {
                    content = reader.ReadToEnd();
                }

                int simpleSpanCount = CountSimpleSpanTags(content);
                
                // すべてのspanタグをカウント（属性ありも含む）
                int allSpanCount = Regex.Matches(content, @"<span[^>]*>", RegexOptions.IgnoreCase).Count;

                var result = new StringBuilder();
                result.AppendLine("HTML spanタグ統計:");
                result.AppendLine("==================");
                result.AppendLine($"ファイルパス: {htmlFilePath}");
                result.AppendLine($"ファイルサイズ: {new FileInfo(htmlFilePath).Length:N0} bytes");
                result.AppendLine($"全spanタグ数: {allSpanCount}");
                result.AppendLine($"シンプルなspanタグ数: {simpleSpanCount}");
                result.AppendLine($"属性付きspanタグ数: {allSpanCount - simpleSpanCount}");

                if (simpleSpanCount > 0)
                {
                    result.AppendLine();
                    result.AppendLine("シンプルなspanタグの例:");
                    var matches = Regex.Matches(content, @"<span\s*>(.*?)</span>", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    int displayCount = 0;
                    foreach (Match match in matches)
                    {
                        if (displayCount >= 5) // 最初の5個のみ表示
                        {
                            result.AppendLine($"... (他{simpleSpanCount - 5}個)");
                            break;
                        }
                        string content_preview = match.Groups[1].Value;
                        if (content_preview.Length > 50)
                        {
                            content_preview = content_preview.Substring(0, 50) + "...";
                        }
                        result.AppendLine($"  <span>{content_preview}</span>");
                        displayCount++;
                    }
                }

                return result.ToString();
            }
            catch (Exception ex)
            {
                return $"統計情報の取得中にエラーが発生しました: {ex.Message}";
            }
        }
    }
}
