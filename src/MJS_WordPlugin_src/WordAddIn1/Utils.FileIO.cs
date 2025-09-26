// Utils.FileIO.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

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

        /// <summary>
        /// 指定されたテキストファイルの内容を読み込み、リスト型変数に格納します
        /// </summary>
        /// <typeparam name="T">リストの要素型</typeparam>
        /// <param name="filePath">読み込むファイルのフルパス</param>
        /// <param name="encoding">文字エンコーディング（省略時はUTF-8）</param>
        /// <param name="separator">要素間の区切り文字（省略時は改行）</param>
        /// <param name="converter">文字列からT型への変換関数</param>
        /// <returns>読み込まれたリスト</returns>
        public static IList<T> ReadListFromFile<T>(string filePath, Encoding encoding = null, 
            string separator = null, Func<string, T> converter = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("ファイルパスが指定されていません。", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"指定されたファイル '{filePath}' が見つかりません。");

            // デフォルト値の設定
            encoding = encoding ?? Encoding.UTF8;
            separator = separator ?? Environment.NewLine;
            converter = converter ?? (str => (T)Convert.ChangeType(str, typeof(T)));

            try
            {
                // ファイルの内容を読み込み
                var content = File.ReadAllText(filePath, encoding);
                
                if (string.IsNullOrEmpty(content))
                    return new List<T>();

                // 区切り文字で分割
                var items = content.Split(new[] { separator }, StringSplitOptions.None);
                
                // T型に変換してリストを作成
                var result = new List<T>();
                foreach (var item in items)
                {
                    try
                    {
                        var convertedItem = converter(item);
                        result.Add(convertedItem);
                    }
                    catch (Exception ex)
                    {
                        throw new InvalidOperationException($"文字列 '{item}' を {typeof(T).Name} 型に変換できませんでした。", ex);
                    }
                }
                
                return result;
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException($"ファイル '{filePath}' への読み取り権限がありません。", ex);
            }
            catch (IOException ex)
            {
                throw new InvalidOperationException($"ファイル '{Path.GetFileName(filePath)}' の読み込み中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// 指定されたテキストファイルの内容を読み込み、文字列リストに格納します（簡易版）
        /// </summary>
        /// <param name="filePath">読み込むファイルのフルパス</param>
        /// <param name="encoding">文字エンコーディング（省略時はUTF-8）</param>
        /// <returns>読み込まれた文字列のリスト</returns>
        public static IList<string> ReadLinesFromFile(string filePath, Encoding encoding = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("ファイルパスが指定されていません。", nameof(filePath));

            if (!File.Exists(filePath))
                throw new FileNotFoundException($"指定されたファイル '{filePath}' が見つかりません。");

            // デフォルト値の設定
            encoding = encoding ?? Encoding.UTF8;

            try
            {
                // ファイルの内容を行ごとに読み込み
                var lines = File.ReadAllLines(filePath, encoding);
                return new List<string>(lines);
            }
            catch (UnauthorizedAccessException ex)
            {
                throw new InvalidOperationException($"ファイル '{filePath}' への読み取り権限がありません。", ex);
            }
            catch (IOException ex)
            {
                throw new InvalidOperationException($"ファイル '{Path.GetFileName(filePath)}' の読み込み中にエラーが発生しました。", ex);
            }
        }

        /// <summary>
        /// フォルダパスとファイル名から完全パスを組み立てて、テキストファイルの内容を読み込みます
        /// </summary>
        /// <typeparam name="T">リストの要素型</typeparam>
        /// <param name="folderPath">ファイルが保存されているフォルダのパス</param>
        /// <param name="fileName">読み込むファイル名（拡張子含む）</param>
        /// <param name="encoding">文字エンコーディング（省略時はUTF-8）</param>
        /// <param name="separator">要素間の区切り文字（省略時は改行）</param>
        /// <param name="converter">文字列からT型への変換関数</param>
        /// <returns>読み込まれたリスト</returns>
        public static IList<T> ReadListFromFile<T>(string folderPath, string fileName, 
            Encoding encoding = null, string separator = null, Func<string, T> converter = null)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("フォルダパスが指定されていません。", nameof(folderPath));
            
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("ファイル名が指定されていません。", nameof(fileName));

            var filePath = Path.Combine(folderPath, fileName);
            return ReadListFromFile<T>(filePath, encoding, separator, converter);
        }

        /// <summary>
        /// フォルダパスとファイル名から完全パスを組み立てて、文字列リストとして読み込みます（簡易版）
        /// </summary>
        /// <param name="folderPath">ファイルが保存されているフォルダのパス</param>
        /// <param name="fileName">読み込むファイル名（拡張子含む）</param>
        /// <param name="encoding">文字エンコーディング（省略時はUTF-8）</param>
        /// <returns>読み込まれた文字列のリスト</returns>
        public static IList<string> ReadLinesFromFile(string folderPath, string fileName, Encoding encoding = null)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("フォルダパスが指定されていません。", nameof(folderPath));
            
            if (string.IsNullOrWhiteSpace(fileName))
                throw new ArgumentException("ファイル名が指定されていません。", nameof(fileName));

            var filePath = Path.Combine(folderPath, fileName);
            return ReadLinesFromFile(filePath, encoding);
        }

        /// <summary>
        /// 現在のWord文書のフォルダにログファイルを設定
        /// </summary>
        /// <param name="logFileName">ログファイル名（省略時は既定値）</param>
        public static void ConfigureLogToDocumentFolder(Microsoft.Office.Interop.Word.Application application,
        string logFileName = "MJS_WordAddIn_logMarker.txt")
        {
            try
            {
                if (Globals.ThisAddIn.Application?.ActiveDocument == null)
                {
                    System.Diagnostics.Debug.WriteLine("アクティブな文書がありません");
                    return;
                }

                string documentPath = Globals.ThisAddIn.Application.ActiveDocument.Path;

                if (string.IsNullOrEmpty(documentPath))
                {
                    System.Diagnostics.Debug.WriteLine("文書が保存されていません。既定のログパスを使用します");
                    return;
                }

                // 既存のTraceListenerを検索して削除
                for (int i = System.Diagnostics.Trace.Listeners.Count - 1; i >= 0; i--)
                {
                    var listener = System.Diagnostics.Trace.Listeners[i];
                    if (listener is System.Diagnostics.TextWriterTraceListener textListener &&
                        listener.Name == "textFileListener")
                    {
                        System.Diagnostics.Trace.Listeners.RemoveAt(i);
                        listener.Dispose();
                        break;
                    }
                }

                // 新しいログファイルパスを作成
                string logFilePath = Path.Combine(documentPath, logFileName);

                // 新しいTraceListenerを追加
                var newListener = new System.Diagnostics.TextWriterTraceListener(logFilePath)
                {
                    Name = "textFileListener",
                    TraceOutputOptions = System.Diagnostics.TraceOptions.DateTime |
                                       System.Diagnostics.TraceOptions.ProcessId |
                                       System.Diagnostics.TraceOptions.ThreadId
                };

                System.Diagnostics.Trace.Listeners.Add(newListener);
                System.Diagnostics.Trace.AutoFlush = true;

                System.Diagnostics.Trace.WriteLine($"ログファイルパスを更新しました: {logFilePath}");
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"ログファイル設定エラー: {ex.Message}");
            }
        }
    }
}
