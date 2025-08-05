// GenerateHTMLButton.ExtractImageInfoFromHtml.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // 画像ファイル名と直前の<p>タグから10文字を抽出するメソッド
        private List<(string imageName, string precedingText)> ExtractImageInfoFromHtml(string htmlFilePath)
        {
            var imageInfoList = new List<(string imageName, string precedingText)>();

            if (!File.Exists(htmlFilePath))
            {
                return imageInfoList;
            }

            try
            {
                string htmlContent = File.ReadAllText(htmlFilePath, Encoding.UTF8);

                // HTMLから<img>タグを検索する正規表現
                var imgPattern = @"<img[^>]*\ssrc=[""']([^""']*)[""'][^>]*>";
                var imgMatches = Regex.Matches(htmlContent, imgPattern, RegexOptions.IgnoreCase);

                foreach (Match imgMatch in imgMatches)
                {
                    string imgSrc = imgMatch.Groups[1].Value;
                    string imageName = Path.GetFileName(imgSrc);

                    // 画像タグの位置を取得
                    int imgPosition = imgMatch.Index;

                    // 画像タグより前の部分から有効なテキストを持つ<p>タグを検索
                    string precedingText = FindValidPrecedingText(htmlContent, imgPosition);

                    imageInfoList.Add((imageName, precedingText));
                }
            }
            catch (Exception ex)
            {
                // エラーログを出力する場合はここに追加
                System.Diagnostics.Debug.WriteLine($"画像情報抽出エラー: {ex.Message}");
            }

            return imageInfoList;
        }

        // 画像タグより前にある有効なテキストを持つ<p>タグを検索するヘルパーメソッド
        private string FindValidPrecedingText(string htmlContent, int imgPosition)
        {
            // 画像タグより前の部分を取得
            string beforeImg = htmlContent.Substring(0, imgPosition);

            // <p>タグのパターン（より包括的な検索）
            var pPattern = @"<p[^>]*>(.*?)</p>";
            var pMatches = Regex.Matches(beforeImg, pPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);

            // 後ろから順番に<p>タグをチェック
            for (int i = pMatches.Count - 1; i >= 0; i--)
            {
                var pMatch = pMatches[i];
                string pContent = pMatch.Groups[1].Value;

                // HTMLタグを除去
                string cleanText = Regex.Replace(pContent, @"<[^>]+>", "");

                // HTMLエンティティをデコード
                cleanText = System.Net.WebUtility.HtmlDecode(cleanText);

                // 前後の空白を除去
                cleanText = cleanText.Trim();

                // 空文字列や空白のみでない場合は有効なテキストとして採用
                if (!string.IsNullOrEmpty(cleanText) && !string.IsNullOrWhiteSpace(cleanText))
                {
                    // 10文字を取得（文字列が10文字未満の場合はそのまま）
                    return cleanText.Length > 10 ? cleanText.Substring(0, 10) : cleanText;
                }
            }

            // 有効な<p>タグが見つからない場合は空文字列を返す
            return string.Empty;
        }

        // 出力ディレクトリ内のすべてのHTMLファイルから画像情報を抽出し、結果を出力する
        private void OutputImageInfoFromAllHtmlFiles(string outputDirPath, StreamWriter logWriter = null)
        {
            var allImageInfo = new List<(string fileName, string imageName, string precedingText)>();

            try
            {
                // 出力ディレクトリ内のすべてのHTMLファイルを取得
                var htmlFiles = Directory.GetFiles(outputDirPath, "*.html", SearchOption.TopDirectoryOnly);

                foreach (string htmlFile in htmlFiles)
                {
                    string fileName = Path.GetFileName(htmlFile);
                    var imageInfoList = ExtractImageInfoFromHtml(htmlFile);

                    foreach (var (imageName, precedingText) in imageInfoList)
                    {
                        allImageInfo.Add((fileName, imageName, precedingText));
                    }
                }

                // 結果をログファイルに出力
                if (logWriter != null)
                {
                    logWriter.WriteLine("=== 画像ファイル情報一覧 ===");
                    logWriter.WriteLine("HTMLファイル名,画像ファイル名,直前のテキスト(10文字)");

                    foreach (var (fileName, imageName, precedingText) in allImageInfo)
                    {
                        logWriter.WriteLine($"{fileName},{imageName},\"{precedingText}\"");
                    }

                    logWriter.WriteLine($"総画像数: {allImageInfo.Count}");
                    logWriter.WriteLine("");
                }

                // デバッグコンソールにも出力
                System.Diagnostics.Debug.WriteLine("=== 画像ファイル情報一覧 ===");
                foreach (var (fileName, imageName, precedingText) in allImageInfo)
                {
                    System.Diagnostics.Debug.WriteLine($"{fileName}: {imageName} - \"{precedingText}\"");
                }
            }
            catch (Exception ex)
            {
                if (logWriter != null)
                {
                    logWriter.WriteLine($"画像情報出力エラー: {ex.Message}");
                }
                System.Diagnostics.Debug.WriteLine($"画像情報出力エラー: {ex.Message}");
            }
        }
    }
}

