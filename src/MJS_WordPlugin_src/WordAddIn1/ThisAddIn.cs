using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

/*
COM（Component Object Model）は、Microsoft が開発した技術で、異なるプログラミング言語やプロセス間でオブジェクトを共有するための仕組みです。
COM オブジェクトは、主に Windows アプリケーションや Office アドインの開発で使用されます。

特徴：
•プログラミング言語に依存しない。
•バイナリレベルでの互換性を提供。
•インターフェースを通じてオブジェクトにアクセス。

NET と COM の関係：
.NET はマネージコードを使用しますが、COM はアンマネージコードを使用します。
.NET から COM オブジェクトを使用する際には、相互運用性（Interop） が必要です。
        
相互運用性の仕組み：
•.NET では、Microsoft.Office.Interop 名前空間を使用して COM オブジェクトにアクセスします。
•例: Microsoft.Office.Interop.Word.Application は Word アプリケーションの COM オブジェクトを表します。
*/

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //SaveBackupOfActiveDocument();
            //this.Application.DocumentBeforeClose += Application_DocumentBeforeClose;

            //SaveDocxBackupIfDoc("backup");
            //ShowBookInfoIdStatusDialog();
            //OverwriteDocument();
            //ConvertBackupxToZip("backup");
            //UnzipBackupxZip("backup");
            //MoveUnzippedMediaFolderToDocumentDirectory("backup");
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                //DeleteBackupFilesOfActiveDocument();

                // 必要に応じてリソースを解放
                if (Globals.ThisAddIn.Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.Application);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"アドイン終了時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

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

        // ドキュメントを閉じる直前に呼ばれる
        private void Application_DocumentBeforeClose(Word.Document doc, ref bool Cancel)
        {
            try
            {
                // バックアップファイル削除処理を追加
                DeleteBackupFilesOfActiveDocument();

                //if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                //{
                //    string name = Path.GetFileNameWithoutExtension(doc.FullName);
                //    string dir = Path.GetDirectoryName(doc.FullName);
                //    string ext = Path.GetExtension(doc.FullName);

                //    if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
                //    {
                //        // 元ファイルのパスを生成
                //        string originalName = name.Substring(0, name.Length - "_backup".Length);
                //        string originalPath = Path.Combine(dir, originalName + ext);

                //        // 元ファイルが存在すれば開く
                //        if (File.Exists(originalPath))
                //        {
                //            this.Application.Documents.Open(originalPath);
                //        }
                //    }
                //}
            }
            catch { /* 例外は無視 */ }
        }

        private void DeleteBackupFilesOfActiveDocument()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                {
                    string dir = Path.GetDirectoryName(doc.FullName);
                    string ext = Path.GetExtension(doc.FullName);

                    // 指定フォルダ内の同じ拡張子のファイルを取得
                    var files = Directory.GetFiles(dir, "*" + ext);
                    foreach (var file in files)
                    {
                        string name = Path.GetFileNameWithoutExtension(file);
                        if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
                        {
                            try
                            {
                                File.Delete(file);
                            }
                            catch (Exception ex)
                            {
                                MessageBox.Show($"バックアップファイル削除時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"バックアップファイル削除処理でエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }


        #region VSTO で生成されたコード

        /// デザイナーのサポートに必要なメソッドです。
        /// メソッドの内容をコードエディターで変更しないでください。
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}