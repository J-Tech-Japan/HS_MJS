// Utils.MergeTwoFolders.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Drawing;
using System.Drawing.Imaging;
using System.Security.Cryptography;
using System.Windows.Forms;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 2つのフォルダ内のPNG画像を比較し、高画質な画像で置き換える
        /// </summary>
        /// <param name="sourceFolder1">ベースとなるフォルダ（このフォルダ内のファイルが置き換え対象）</param>
        /// <param name="sourceFolder2">比較元フォルダ（このフォルダ内により高画質な画像があれば置き換える）</param>
        public static void MergeTwoFolders(string sourceFolder1, string sourceFolder2)
        {
            try
            {
                if (!Directory.Exists(sourceFolder1))
                {
                    throw new DirectoryNotFoundException($"ソースフォルダ1が見つかりません: {sourceFolder1}");
                }

                if (!Directory.Exists(sourceFolder2))
                {
                    throw new DirectoryNotFoundException($"ソースフォルダ2が見つかりません: {sourceFolder2}");
                }

                // sourceFolder1のPNGファイルを取得
                var pngFiles1 = Directory.GetFiles(sourceFolder1, "*.png", SearchOption.TopDirectoryOnly);
                var pngFiles2 = Directory.GetFiles(sourceFolder2, "*.png", SearchOption.TopDirectoryOnly);

                if (pngFiles1.Length == 0)
                {
                    return; // 置き換える対象がない
                }

                if (pngFiles2.Length == 0)
                {
                    return; // 比較元がない
                }

                int replacedCount = 0;

                foreach (string file1 in pngFiles1)
                {
                    string fileName1 = Path.GetFileNameWithoutExtension(file1);
                    
                    // sourceFolder2で同じベース名または類似画像を探す
                    string bestMatch = FindBestMatchingImage(file1, pngFiles2);
                    
                    if (!string.IsNullOrEmpty(bestMatch))
                    {
                        if (IsHigherQuality(bestMatch, file1))
                        {
                            // バックアップを作成（念のため）
                            string backupPath = file1 + ".backup";
                            if (File.Exists(backupPath))
                            {
                                File.Delete(backupPath);
                            }
                            File.Copy(file1, backupPath);

                            // より高画質な画像で置き換え
                            File.Copy(bestMatch, file1, true);
                            replacedCount++;

                            Console.WriteLine($"置き換え完了: {Path.GetFileName(file1)} <- {Path.GetFileName(bestMatch)}");
                        }
                    }
                }

                if (replacedCount > 0)
                {
                    Console.WriteLine($"合計 {replacedCount} 個のファイルを高画質な画像で置き換えました。");
                }
                else
                {
                    Console.WriteLine("置き換えが必要な画像は見つかりませんでした。");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"フォルダ結合中にエラーが発生しました: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// 対象画像に最も類似した画像を候補リストから見つける
        /// </summary>
        private static string FindBestMatchingImage(string targetImage, string[] candidateImages)
        {
            string targetFileName = Path.GetFileNameWithoutExtension(targetImage);
            
            // 1. 完全一致するファイル名を探す
            var exactMatch = candidateImages.FirstOrDefault(img => 
                Path.GetFileNameWithoutExtension(img).Equals(targetFileName, StringComparison.OrdinalIgnoreCase));
            
            if (!string.IsNullOrEmpty(exactMatch))
            {
                return exactMatch;
            }

            // 2. 画像内容の類似性で判定
            try
            {
                using (var targetBitmap = new Bitmap(targetImage))
                {
                    string bestMatch = null;
                    double bestSimilarity = 0.0;

                    foreach (string candidate in candidateImages)
                    {
                        try
                        {
                            using (var candidateBitmap = new Bitmap(candidate))
                            {
                                double similarity = CalculateImageSimilarity(targetBitmap, candidateBitmap);
                                if (similarity > bestSimilarity && similarity > 0.85) // 85%以上の類似性
                                {
                                    bestSimilarity = similarity;
                                    bestMatch = candidate;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"画像読み込みエラー ({candidate}): {ex.Message}");
                            continue;
                        }
                    }

                    return bestSimilarity > 0.85 ? bestMatch : null;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"画像比較エラー ({targetImage}): {ex.Message}");
                return null;
            }
        }

        /// <summary>
        /// 2つの画像の類似性を計算（0.0-1.0の範囲）
        /// </summary>
        private static double CalculateImageSimilarity(Bitmap img1, Bitmap img2)
        {
            // サイズが大きく異なる場合は類似性が低いと判定
            double sizeRatio = Math.Min((double)img1.Width / img2.Width, (double)img2.Width / img1.Width) *
                              Math.Min((double)img1.Height / img2.Height, (double)img2.Height / img1.Height);
            
            if (sizeRatio < 0.5) // サイズが50%以上異なる場合
            {
                return 0.0;
            }

            // 小さいサイズにリサイズして比較を高速化
            const int compareSize = 32;
            using (var resized1 = ResizeImage(img1, compareSize, compareSize))
            using (var resized2 = ResizeImage(img2, compareSize, compareSize))
            {
                int totalPixels = compareSize * compareSize;
                int matchingPixels = 0;
                int tolerance = 30; // RGB値の許容差

                for (int x = 0; x < compareSize; x++)
                {
                    for (int y = 0; y < compareSize; y++)
                    {
                        Color pixel1 = resized1.GetPixel(x, y);
                        Color pixel2 = resized2.GetPixel(x, y);

                        int rDiff = Math.Abs(pixel1.R - pixel2.R);
                        int gDiff = Math.Abs(pixel1.G - pixel2.G);
                        int bDiff = Math.Abs(pixel1.B - pixel2.B);

                        if (rDiff <= tolerance && gDiff <= tolerance && bDiff <= tolerance)
                        {
                            matchingPixels++;
                        }
                    }
                }

                return (double)matchingPixels / totalPixels;
            }
        }

        /// <summary>
        /// 画像をリサイズする
        /// </summary>
        private static Bitmap ResizeImage(Bitmap original, int width, int height)
        {
            var resized = new Bitmap(width, height);
            using (var graphics = Graphics.FromImage(resized))
            {
                graphics.InterpolationMode = System.Drawing.Drawing2D.InterpolationMode.HighQualityBicubic;
                graphics.DrawImage(original, 0, 0, width, height);
            }
            return resized;
        }

        /// <summary>
        /// 画像1が画像2より高画質かどうかを判定
        /// </summary>
        private static bool IsHigherQuality(string imagePath1, string imagePath2)
        {
            try
            {
                var fileInfo1 = new FileInfo(imagePath1);
                var fileInfo2 = new FileInfo(imagePath2);

                // ファイルサイズによる簡易判定
                if (fileInfo1.Length > fileInfo2.Length * 1.2) // 20%以上大きい場合
                {
                    using (var img1 = new Bitmap(imagePath1))
                    using (var img2 = new Bitmap(imagePath2))
                    {
                        // 解像度による判定
                        long pixels1 = (long)img1.Width * img1.Height;
                        long pixels2 = (long)img2.Width * img2.Height;
                        
                        if (pixels1 > pixels2)
                        {
                            return true;
                        }
                        
                        // 同じ解像度の場合、ファイルサイズで判定（圧縮率が低い = 高画質）
                        if (pixels1 == pixels2 && fileInfo1.Length > fileInfo2.Length)
                        {
                            return true;
                        }
                    }
                }

                return false;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"画質比較エラー: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// 相対パスを取得する (.NET Framework 4.8互換)
        /// </summary>
        private static string GetRelativePath(string fromPath, string toPath)
        {
            if (string.IsNullOrEmpty(fromPath)) throw new ArgumentNullException(nameof(fromPath));
            if (string.IsNullOrEmpty(toPath)) throw new ArgumentNullException(nameof(toPath));

            try
            {
                var fromUri = new Uri(fromPath.EndsWith("\\") ? fromPath : fromPath + "\\");
                var toUri = new Uri(toPath);

                if (fromUri.Scheme != toUri.Scheme) 
                { 
                    return toPath; // path can't be made relative.
                }

                var relativeUri = fromUri.MakeRelativeUri(toUri);
                var relativePath = Uri.UnescapeDataString(relativeUri.ToString());

                if (toUri.Scheme.Equals("file", StringComparison.InvariantCultureIgnoreCase))
                {
                    relativePath = relativePath.Replace('/', Path.DirectorySeparatorChar);
                }

                return relativePath;
            }
            catch
            {
                // URI変換に失敗した場合は、単純なパス置換を試行
                if (toPath.StartsWith(fromPath))
                {
                    return toPath.Substring(fromPath.Length).TrimStart('\\', '/');
                }
                return toPath;
            }
        }

        /// <summary>
        /// フォルダ結合処理のオーバーロード（既存コードとの互換性のため）
        /// </summary>
        public static bool MergeTwoFolders(string sourceFolder1, string sourceFolder2, string destinationFolder, bool deleteSourceFolders = false)
        {
            try
            {
                // 宛先フォルダを作成
                if (!Directory.Exists(destinationFolder))
                {
                    Directory.CreateDirectory(destinationFolder);
                }

                // sourceFolder1のファイルをコピー
                if (Directory.Exists(sourceFolder1))
                {
                    foreach (string file in Directory.GetFiles(sourceFolder1, "*", SearchOption.AllDirectories))
                    {
                        string relativePath = GetRelativePath(sourceFolder1, file);
                        string destPath = Path.Combine(destinationFolder, relativePath);
                        
                        Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                        File.Copy(file, destPath, true);
                    }
                }

                // sourceFolder2のファイルをコピー（重複があれば高画質判定で置換）
                if (Directory.Exists(sourceFolder2))
                {
                    foreach (string file in Directory.GetFiles(sourceFolder2, "*", SearchOption.AllDirectories))
                    {
                        string relativePath = GetRelativePath(sourceFolder2, file);
                        string destPath = Path.Combine(destinationFolder, relativePath);
                        
                        Directory.CreateDirectory(Path.GetDirectoryName(destPath));
                        
                        if (File.Exists(destPath) && Path.GetExtension(file).ToLower() == ".png")
                        {
                            // PNG画像の場合、高画質判定で置換
                            if (IsHigherQuality(file, destPath))
                            {
                                File.Copy(file, destPath, true);
                            }
                        }
                        else
                        {
                            File.Copy(file, destPath, true);
                        }
                    }
                }

                // 元フォルダを削除（オプション）
                if (deleteSourceFolders)
                {
                    if (Directory.Exists(sourceFolder1))
                    {
                        Directory.Delete(sourceFolder1, true);
                    }
                    if (Directory.Exists(sourceFolder2))
                    {
                        Directory.Delete(sourceFolder2, true);
                    }
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"フォルダ結合エラー: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// フォルダ結合の統計情報を取得
        /// </summary>
        public static string GetMergeFolderStats(string sourceFolder1, string sourceFolder2, string destinationFolder)
        {
            try
            {
                int files1 = Directory.Exists(sourceFolder1) ? Directory.GetFiles(sourceFolder1, "*", SearchOption.AllDirectories).Length : 0;
                int files2 = Directory.Exists(sourceFolder2) ? Directory.GetFiles(sourceFolder2, "*", SearchOption.AllDirectories).Length : 0;
                int destFiles = Directory.Exists(destinationFolder) ? Directory.GetFiles(destinationFolder, "*", SearchOption.AllDirectories).Length : 0;

                return $"ソースフォルダ1: {files1}ファイル\nソースフォルダ2: {files2}ファイル\n結合後: {destFiles}ファイル";
            }
            catch (Exception ex)
            {
                return $"統計情報の取得に失敗しました: {ex.Message}";
            }
        }
    }
}
