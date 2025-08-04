// GenerateHTMLButton.CollectMergeScript.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private Dictionary<string, string> CollectMergeScriptDict(Word.Document activeDocument)
        {
            var mergeScript = new Dictionary<string, string>();
            CollectMergeScript(activeDocument.Path, activeDocument.Name, mergeScript);
            return mergeScript;
        }

        // 見出し結合情報を書誌情報ファイル (headerFile) から収集
        public void CollectMergeScript(string documentPath, string documentName, Dictionary<string, string> mergeScript)
        {
            try
            {
                // ファイルパスを安全に結合
                // ドキュメント名の最初の3文字を抽出して、対応するヘッダーファイルのパスを生成
                string headerFilePath = Path.Combine(documentPath, "headerFile", Regex.Replace(documentName, "^(.{3}).+$", "$1") + ".txt");

                // ヘッダーファイルをUTF-8エンコーディングで読み込む
                using (StreamReader sr = new StreamReader(headerFilePath, Encoding.UTF8))
                {
                    // ファイルの終端まで1行ずつ読み込む
                    while (sr.Peek() >= 0)
                    {
                        string strBuffer = sr.ReadLine();

                        // 空行や空白行をスキップ
                        if (string.IsNullOrWhiteSpace(strBuffer)) continue;

                        // タブ区切りで行を分割
                        string[] info = strBuffer.Split('\t');

                        // 必要な情報が揃っている場合のみ処理を続行
                        if (info.Length == 4 && !string.IsNullOrEmpty(info[3]))
                        {
                            string key = info[2]; // 辞書のキー
                            string value = info[3].Replace("(", "").Replace(")", ""); // 値から括弧を削除

                            // 辞書に同じキーと値のペアが存在しない場合のみ追加
                            if (!mergeScript.ContainsKey(key) || mergeScript[key] != value)
                            {
                                mergeScript[key] = value;
                            }
                        }
                    }
                }
            }
            catch (FileNotFoundException ex)
            {
                // ファイルが見つからない場合のエラーメッセージを表示
                MessageBox.Show($"ファイルが見つかりません: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                // その他の例外が発生した場合のエラーメッセージを表示
                MessageBox.Show($"エラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }
}
