using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace WordAddIn1
{
    internal partial class Utils
    {
        /// <summary>
        /// 指定されたフォルダにテキストファイルを作成し、リスト型変数の内容を書き込みます
        /// </summary>
        /// <typeparam name="T">リストの要素型</typeparam>
        /// <param name="folderPath">作成先フォルダのパス</param>
        /// <param name="fileName">作成するファイル名（拡張子含む）</param>
        /// <param name="list">書き込むリスト</param>
        /// <param name="encoding">文字エンコーディング（省略時はUTF-8）</param>
        /// <param name="separator">要素間の区切り文字（省略時は改行）</param>
        /// <returns>作成されたファイルのフルパス</returns>
        public static string WriteListToFile<T>(string folderPath, string fileName, IList<T> list, 
            Encoding encoding = null, string separator = null)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("フォルダパスが指定されていません。", nameof(folderPath));
            
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("ファイル名が指定されていません。", nameof(fileName));
            
            if (list == null)
                throw new ArgumentNullException(nameof(list));

            // デフォルト値の設定
            encoding = encoding ?? Encoding.UTF8;
            separator = separator ?? Environment.NewLine;

            try
            {
                // フォルダが存在しない場合は作成
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                var filePath = Path.Combine(folderPath, fileName);
                
                // リストの内容を文字列に変換
                var content = string.Join(separator, list.Select(item => item?.ToString() ?? string.Empty));
                
                // ファイルに書き込み
                File.WriteAllText(filePath, content, encoding);
                
                return filePath;
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException($"フォルダ '{folderPath}' への書き込み権限がありません。", ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                throw new InvalidOperationException($"指定されたパス '{folderPath}' が見つかりません。", ex);
            }
            catch (IOException ex)
            {
                throw new InvalidOperationException($"ファイル '{fileName}' の書き込み中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// 指定されたフォルダにテキストファイルを作成し、文字列リストの内容を書き込みます（簡易版）
        /// </summary>
        /// <param name="folderPath">作成先フォルダのパス</param>
        /// <param name="fileName">作成するファイル名（拡張子含む）</param>
        /// <param name="lines">書き込む文字列のリスト</param>
        /// <returns>作成されたファイルのフルパス</returns>
        public static string WriteLinesToFile(string folderPath, string fileName, IList<string> lines)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("フォルダパスが指定されていません。", nameof(folderPath));
            
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("ファイル名が指定されていません。", nameof(fileName));
            
            if (lines == null)
                throw new ArgumentNullException(nameof(lines));

            try
            {
                // フォルダが存在しない場合は作成
                if (!Directory.Exists(folderPath))
                {
                    Directory.CreateDirectory(folderPath);
                }

                var filePath = Path.Combine(folderPath, fileName);
                
                // ファイルに書き込み（UTF-8エンコーディング）
                File.WriteAllLines(filePath, lines, Encoding.UTF8);
                
                return filePath;
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException($"フォルダ '{folderPath}' への書き込み権限がありません。", ex);
            }
            catch (DirectoryNotFoundException ex)
            {
                throw new InvalidOperationException($"指定されたパス '{folderPath}' が見つかりません。", ex);
            }
            catch (IOException ex)
            {
                throw new InvalidOperationException($"ファイル '{fileName}' の書き込み中にエラーが発生しました。", ex);
            }
        }
    }
}
