// ThisAddIn.Backup.cs

using System;
using System.IO;
using System.IO.Compression;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        /// 書誌情報IDが設定されている場合は true を返す
        private bool IsBookInfoIdSet()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc != null)
                {
                    foreach (Word.Bookmark bm in doc.Bookmarks)
                    {
                        // 書誌情報IDのパターン（例: 3文字+2桁+3桁）にマッチするか
                        if (System.Text.RegularExpressions.Regex.IsMatch(bm.Name, @"^[A-Z0-9]{3}\d{2}\d{3}$"))
                        {
                            return true; // 設定されている
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"書誌情報ID判定時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                // エラー時は「設定されていない」とみなす
            }
            return false; // 設定されていない
        }

        /// 書誌情報IDが設定されていない場合のみダイアログで通知する
        private void ShowBookInfoIdStatusDialog()
        {
            if (!IsBookInfoIdSet())
            {
                try
                {
                    var doc = this.Application.ActiveDocument;
                    if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                    {
                        string dir = Path.GetDirectoryName(doc.FullName);
                        string headerDir = Path.Combine(dir, "headerFile");
                        if (Directory.Exists(headerDir))
                        {
                            Directory.Delete(headerDir, true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"headerFileフォルダ削除時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }

                MessageBox.Show("書誌情報IDが設定されていません。", "書誌情報チェック", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
