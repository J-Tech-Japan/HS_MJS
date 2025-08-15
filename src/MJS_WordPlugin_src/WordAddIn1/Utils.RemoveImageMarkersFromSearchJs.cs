// Utils.RemoveImageMarkersFromSearchJs.cs

using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// search.jsファイルから[imagemarker:xxx]パターンを削除する
        /// xxxは任意の文字列。大文字・小文字は区別しない。
        /// </summary>
        /// <param name="searchJsFilePath">search.jsファイルのパス</param>
        /// <returns>削除されたマーカーの数。エラーの場合は-1</returns>
        public static int RemoveImageMarkersFromSearchJs(string searchJsFilePath)
        {
            if (string.IsNullOrEmpty(searchJsFilePath))
            {
                System.Diagnostics.Debug.WriteLine("search.jsファイルパスが指定されていません。");
                return -1;
            }

            if (!File.Exists(searchJsFilePath))
            {
                System.Diagnostics.Debug.WriteLine($"search.jsファイルが存在しません: {searchJsFilePath}");
                return -1;
            }

            try
            {
                // ファイル内容を読み込み
                string content;
                using (var reader = new StreamReader(searchJsFilePath, Encoding.UTF8))
                {
                    content = reader.ReadToEnd();
                }

                // 削除前のマーカー数をカウント
                int originalMarkerCount = CountImageMarkers(content);

                // [imagemarker:xxx]パターンを削除（大文字小文字区別なし）
                string cleanedContent = RemoveImageMarkerPatterns(content);

                // 削除後のマーカー数をカウント
                int remainingMarkerCount = CountImageMarkers(cleanedContent);
                int removedCount = originalMarkerCount - remainingMarkerCount;

                // ファイルに書き戻し
                using (var writer = new StreamWriter(searchJsFilePath, false, Encoding.UTF8))
                {
                    writer.Write(cleanedContent);
                }

                System.Diagnostics.Debug.WriteLine($"search.jsから{removedCount}個のイメージマーカーを削除しました。");
                return removedCount;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"search.jsの処理中にエラーが発生しました: {ex.Message}");
                return -1;
            }
        }

        /// <summary>
        /// 指定されたディレクトリ内のsearch.jsファイルから[imagemarker:xxx]パターンを削除する
        /// </summary>
        /// <param name="directoryPath">search.jsが含まれるディレクトリのパス</param>
        /// <returns>削除されたマーカーの数。エラーの場合は-1</returns>
        public static int RemoveImageMarkersFromSearchJsInDirectory(string directoryPath)
        {
            if (string.IsNullOrEmpty(directoryPath))
            {
                System.Diagnostics.Debug.WriteLine("ディレクトリパスが指定されていません。");
                return -1;
            }

            if (!Directory.Exists(directoryPath))
            {
                System.Diagnostics.Debug.WriteLine($"ディレクトリが存在しません: {directoryPath}");
                return -1;
            }

            string searchJsPath = Path.Combine(directoryPath, "search.js");
            return RemoveImageMarkersFromSearchJs(searchJsPath);
        }

        /// <summary>
        /// 文字列から[imagemarker:xxx]パターンを削除する
        /// </summary>
        /// <param name="input">処理対象の文字列</param>
        /// <returns>クリーニング済みの文字列</returns>
        private static string RemoveImageMarkerPatterns(string input)
        {
            if (string.IsNullOrEmpty(input))
                return input;

            // [imagemarker:xxx]パターンにマッチする正規表現
            // ケースインセンシティブ（大文字小文字区別なし）
            string pattern = @"\[imagemarker:[^\]]*\]";

            string result = Regex.Replace(input, pattern, "", RegexOptions.IgnoreCase);

            // 連続した空白や改行を整理
            result = Regex.Replace(result, @"\s+", " ");
            result = Regex.Replace(result, @"\s*\n\s*", "\n");

            return result;
        }

        /// <summary>
        /// 文字列内の[imagemarker:xxx]パターンの数をカウントする
        /// </summary>
        /// <param name="input">カウント対象の文字列</param>
        /// <returns>見つかったマーカーの数</returns>
        private static int CountImageMarkers(string input)
        {
            if (string.IsNullOrEmpty(input))
                return 0;

            string pattern = @"\[imagemarker:[^\]]*\]";
            var matches = Regex.Matches(input, pattern, RegexOptions.IgnoreCase);
            return matches.Count;
        }

        /// <summary>
        /// search.jsファイル内のイメージマーカーの統計情報を取得する
        /// </summary>
        /// <param name="searchJsFilePath">search.jsファイルのパス</param>
        /// <returns>統計情報の文字列</returns>
        public static string GetSearchJsImageMarkerStatistics(string searchJsFilePath)
        {
            if (string.IsNullOrEmpty(searchJsFilePath) || !File.Exists(searchJsFilePath))
                return "search.jsファイルが見つかりません。";

            try
            {
                string content;
                using (var reader = new StreamReader(searchJsFilePath, Encoding.UTF8))
                {
                    content = reader.ReadToEnd();
                }

                int markerCount = CountImageMarkers(content);

                var result = new StringBuilder();
                result.AppendLine("search.js イメージマーカー統計:");
                result.AppendLine("==========================");
                result.AppendLine($"ファイルパス: {searchJsFilePath}");
                result.AppendLine($"ファイルサイズ: {new FileInfo(searchJsFilePath).Length:N0} bytes");
                result.AppendLine($"イメージマーカー数: {markerCount}");

                if (markerCount > 0)
                {
                    result.AppendLine();
                    result.AppendLine("見つかったマーカー:");
                    var matches = Regex.Matches(content, @"\[imagemarker:[^\]]*\]", RegexOptions.IgnoreCase);
                    int displayCount = 0;
                    foreach (Match match in matches)
                    {
                        if (displayCount >= 10) // 最初の10個のみ表示
                        {
                            result.AppendLine($"... (他{markerCount - 10}個)");
                            break;
                        }
                        result.AppendLine($"  {match.Value}");
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