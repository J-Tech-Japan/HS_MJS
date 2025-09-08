// Utils.ProcessImageMarkers.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// webhelpディレクトリ内のすべてのHTMLファイルを処理し、IMAGEMARKERに基づいて画像のsrcを変更
        /// </summary>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス</param>
        /// <param name="extractedImagesDirectory">extracted_imagesディレクトリの相対パス（例: "extracted_images"）</param>
        /// <returns>処理されたファイル数</returns>
        public static int ProcessImageMarkersInWebhelp(
            string webhelpDirectory,
            string extractedImagesDirectory = "extracted_images")
        {
            if (string.IsNullOrEmpty(webhelpDirectory))
                throw new ArgumentException("webhelpディレクトリが指定されていません。", nameof(webhelpDirectory));

            if (!Directory.Exists(webhelpDirectory))
                throw new DirectoryNotFoundException($"指定されたディレクトリが存在しません: {webhelpDirectory}");

            int processedFileCount = 0;

            try
            {
                // webhelpディレクトリ内のすべてのHTMLファイルを取得
                var htmlFiles = Directory.GetFiles(webhelpDirectory, "*.html", SearchOption.AllDirectories);

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        if (ProcessSingleHtmlFile(htmlFilePath, extractedImagesDirectory))
                        {
                            processedFileCount++;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"ファイル処理エラー ({htmlFilePath}): {ex.Message}");
                    }
                }

                return processedFileCount;
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"HTML画像マーカー処理中にエラーが発生しました: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// 単一のHTMLファイルを処理し、IMAGEMARKERに基づいて画像のsrcを変更
        /// </summary>
        /// <param name="htmlFilePath">処理対象のHTMLファイルパス</param>
        /// <param name="extractedImagesDirectory">extracted_imagesディレクトリの相対パス</param>
        /// <returns>変更が行われた場合はtrue</returns>
        private static bool ProcessSingleHtmlFile(string htmlFilePath, string extractedImagesDirectory)
        {
            // HTMLファイルを読み込み
            string htmlContent;
            using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
            {
                htmlContent = reader.ReadToEnd();
            }

            string originalContent = htmlContent;
            
            // webhelpディレクトリのパスを取得（HTMLファイルの親ディレクトリまたはその上位）
            string webhelpDirectory = Path.GetDirectoryName(htmlFilePath);
            while (!string.IsNullOrEmpty(webhelpDirectory))
            {
                if (Directory.Exists(Path.Combine(webhelpDirectory, extractedImagesDirectory)))
                {
                    break;
                }
                string parentDir = Directory.GetParent(webhelpDirectory)?.FullName;
                if (parentDir == webhelpDirectory) break; // ルートディレクトリに到達
                webhelpDirectory = parentDir;
            }

            // 画像とマーカーのペアを処理
            htmlContent = ProcessImageAndMarkerPairs(htmlContent, extractedImagesDirectory, webhelpDirectory);

            // 残りのIMAGEMARKERを削除
            htmlContent = RemoveRemainingImageMarkers(htmlContent);

            // 変更があった場合のみファイルを保存
            if (htmlContent != originalContent)
            {
                using (var writer = new StreamWriter(htmlFilePath, false, Encoding.UTF8))
                {
                    writer.Write(htmlContent);
                }
                return true;
            }

            return false;
        }

        /// <summary>
        /// <img>タグの直後にある[IMAGEMARKER:xxx]を処理し、imgのsrcを変更
        /// </summary>
        /// <param name="htmlContent">処理対象のHTML内容</param>
        /// <param name="extractedImagesDirectory">extracted_imagesディレクトリの相対パス</param>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス（古い画像削除用）</param>
        /// <returns>処理されたHTML内容</returns>
        private static string ProcessImageAndMarkerPairs(string htmlContent, string extractedImagesDirectory, string webhelpDirectory = null)
        {
            // 修正版：<img>タグと</p>の間にテキストがある場合も対応
            string pattern = @"(<img[^>]*>)([^<]*)</p>\s*<p[^>]*>\s*\[IMAGEMARKER:([^\]]+)\]\s*</p>";

            var regex = new Regex(pattern, RegexOptions.IgnoreCase | RegexOptions.Multiline);

            htmlContent = regex.Replace(htmlContent, match =>
            {
                string imgTag = match.Groups[1].Value;
                string additionalText = match.Groups[2].Value; // 追加テキスト
                string markerValue = match.Groups[3].Value;

                // imgタグのsrc属性を新しいパスに変更（古い画像ファイル削除含む）
                string newSrc = $"{extractedImagesDirectory}/{markerValue}.png";
                string updatedImgTag = UpdateImageSrc(imgTag, newSrc, webhelpDirectory);

                // 追加テキストも含めて返す
                return updatedImgTag + additionalText + "</p>";
            });

            return htmlContent;
        }

        /// <summary>
        /// <img>タグのsrc属性を新しい値に変更し、古い画像ファイルを削除
        /// </summary>
        /// <param name="imgTag">元の<img>タグ</param>
        /// <param name="newSrc">新しいsrc値</param>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス（古い画像削除用）</param>
        /// <returns>src属性が更新された<img>タグ</returns>
        private static string UpdateImageSrc(string imgTag, string newSrc, string webhelpDirectory = null)
        {
            // src属性のパターンをマッチ
            var srcPattern = @"src\s*=\s*[""']([^""']*)[""']";
            var srcRegex = new Regex(srcPattern, RegexOptions.IgnoreCase);

            if (srcRegex.IsMatch(imgTag))
            {
                var match = srcRegex.Match(imgTag);
                string oldSrc = match.Groups[1].Value;
                
                // 古い画像ファイルを削除
                if (!string.IsNullOrEmpty(webhelpDirectory) && !string.IsNullOrEmpty(oldSrc))
                {
                    DeleteOldImageFile(oldSrc, webhelpDirectory);
                }
                
                // 既存のsrc属性を新しい値に置換
                return srcRegex.Replace(imgTag, $"src=\"{newSrc}\"");
            }
            else
            {
                // src属性が存在しない場合は追加
                // <img の直後に挿入
                var insertPattern = @"(<img)(\s|>)";
                var insertRegex = new Regex(insertPattern, RegexOptions.IgnoreCase);
                
                if (insertRegex.IsMatch(imgTag))
                {
                    return insertRegex.Replace(imgTag, $"$1 src=\"{newSrc}\"$2");
                }
            }

            return imgTag;
        }

        /// <summary>
        /// 古い画像ファイルを削除
        /// </summary>
        /// <param name="oldSrc">古いsrc属性値</param>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス</param>
        private static void DeleteOldImageFile(string oldSrc, string webhelpDirectory)
        {
            try
            {
                // 相対パスを絶対パスに変換
                string oldImagePath;
                
                if (Path.IsPathRooted(oldSrc))
                {
                    // 絶対パスの場合
                    oldImagePath = oldSrc;
                }
                else
                {
                    // 相対パスの場合、webhelpディレクトリからの相対パス
                    oldImagePath = Path.Combine(webhelpDirectory, oldSrc);
                }
                
                // ファイルが存在する場合は削除
                if (File.Exists(oldImagePath))
                {
                    File.Delete(oldImagePath);
                    System.Diagnostics.Debug.WriteLine($"古い画像ファイルを削除しました: {oldImagePath}");
                }
            }
            catch (Exception ex)
            {
                // ファイル削除エラーはログ出力のみで処理を継続
                System.Diagnostics.Debug.WriteLine($"画像ファイル削除エラー ({oldSrc}): {ex.Message}");
            }
        }

        /// <summary>
        /// 残りの[IMAGEMARKER:xxx]パターンをすべて削除
        /// </summary>
        /// <param name="htmlContent">処理対象のHTML内容</param>
        /// <returns>マーカーが削除されたHTML内容</returns>
        private static string RemoveRemainingImageMarkers(string htmlContent)
        {
            // <p>タグで囲まれた[IMAGEMARKER:xxx]パターンを削除
            string paragraphMarkerPattern = @"<p[^>]*>\s*\[IMAGEMARKER:[^\]]+\]\s*</p>";
            htmlContent = Regex.Replace(htmlContent, paragraphMarkerPattern, "", RegexOptions.IgnoreCase | RegexOptions.Multiline);

            // その他の場所にある[IMAGEMARKER:xxx]パターンも削除
            string markerPattern = @"\[IMAGEMARKER:[^\]]+\]";
            htmlContent = Regex.Replace(htmlContent, markerPattern, "", RegexOptions.IgnoreCase);

            // 連続する空行や余分な空白を整理
            htmlContent = Regex.Replace(htmlContent, @"\n\s*\n\s*\n", "\n\n");

            return htmlContent;
        }

        /// <summary>
        /// 処理統計情報を取得
        /// </summary>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス</param>
        /// <returns>統計情報の文字列</returns>
        public static string GetImageMarkerProcessingStatistics(string webhelpDirectory)
        {
            if (string.IsNullOrEmpty(webhelpDirectory) || !Directory.Exists(webhelpDirectory))
                return "指定されたディレクトリが存在しません。";

            try
            {
                var htmlFiles = Directory.GetFiles(webhelpDirectory, "*.html", SearchOption.AllDirectories);
                int totalFiles = htmlFiles.Length;
                int filesWithMarkers = 0;
                int totalMarkers = 0;

                foreach (string htmlFilePath in htmlFiles)
                {
                    try
                    {
                        string content;
                        using (var reader = new StreamReader(htmlFilePath, Encoding.UTF8))
                        {
                            content = reader.ReadToEnd();
                        }

                        var markerMatches = Regex.Matches(content, @"\[IMAGEMARKER:[^\]]+\]", RegexOptions.IgnoreCase);
                        if (markerMatches.Count > 0)
                        {
                            filesWithMarkers++;
                            totalMarkers += markerMatches.Count;
                        }
                    }
                    catch (Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"統計取得エラー ({htmlFilePath}): {ex.Message}");
                    }
                }

                var statistics = new StringBuilder();
                statistics.AppendLine("画像マーカー処理統計:");
                statistics.AppendLine("====================");
                statistics.AppendLine($"総HTMLファイル数: {totalFiles}");
                statistics.AppendLine($"マーカーを含むファイル数: {filesWithMarkers}");
                statistics.AppendLine($"総マーカー数: {totalMarkers}");

                return statistics.ToString();
            }
            catch (Exception ex)
            {
                return $"統計取得中にエラーが発生しました: {ex.Message}";
            }
        }

        /// <summary>
        /// extracted_imagesディレクトリ内のファイル一覧を取得
        /// </summary>
        /// <param name="webhelpDirectory">webhelpディレクトリのパス</param>
        /// <param name="extractedImagesDirectory">extracted_imagesディレクトリの相対パス</param>
        /// <returns>抽出された画像ファイルのパス一覧</returns>
        public static List<string> GetExtractedImageFiles(
            string webhelpDirectory,
            string extractedImagesDirectory = "extracted_images")
        {
            string extractedImagesPath = Path.Combine(webhelpDirectory, extractedImagesDirectory);
            
            if (!Directory.Exists(extractedImagesPath))
                return new List<string>();

            try
            {
                return Directory.GetFiles(extractedImagesPath, "*.png", SearchOption.TopDirectoryOnly)
                    .Select(Path.GetFileName)
                    .ToList();
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"画像ファイル一覧取得エラー: {ex.Message}");
                return new List<string>();
            }
        }
    }
}