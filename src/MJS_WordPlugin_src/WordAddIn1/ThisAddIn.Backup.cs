using System;
using System.IO;
using System.IO.Compression;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        // アクティブドキュメントのバックアップを保存（フォルダ名を引数で指定）
        //private void SaveBackupOfActiveDocument()
        //{
        //    try
        //    {
        //        var doc = this.Application.ActiveDocument;
        //        if (doc != null && !string.IsNullOrEmpty(doc.FullName))
        //        {
        //            string originalPath = doc.FullName;
        //            string dir = Path.GetDirectoryName(originalPath);
        //            string name = Path.GetFileNameWithoutExtension(originalPath);
        //            string ext = Path.GetExtension(originalPath);

        //            // 末尾が "_backup" の場合は何もしない
        //            if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
        //                return;

        //            string backupName = name + "_backup" + ext;
        //            string backupPath = Path.Combine(dir, backupName);

        //            // すでにバックアップファイルが存在する場合は何もしない
        //            if (File.Exists(backupPath))
        //                return;

        //            // ファイルをコピーしてバックアップ作成
        //            File.Copy(originalPath, backupPath);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show($"バックアップ保存時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        //    }
        //}

        private void SaveBackupOfActiveDocument()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc != null && !string.IsNullOrEmpty(doc.FullName))
                {
                    string originalPath = doc.FullName;
                    string dir = Path.GetDirectoryName(originalPath);
                    string name = Path.GetFileNameWithoutExtension(originalPath);
                    string ext = Path.GetExtension(originalPath);

                    // 末尾が "_backup" の場合は何もしない
                    if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
                        return;

                    string backupName = name + "_backup" + ext;
                    string backupPath = Path.Combine(dir, backupName);

                    // すでにバックアップファイルが存在する場合は何もしない
                    if (File.Exists(backupPath))
                        return;

                    // ファイルをコピーしてバックアップ作成
                    File.Copy(originalPath, backupPath);

                    // 現在のドキュメントを閉じてバックアップファイルを開く
                    object saveChanges = false;
                    object originalDoc = doc;
                    this.Application.ActiveDocument.Close(ref saveChanges);
                    this.Application.Documents.Open(backupPath);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"バックアップ保存時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void SaveDocxBackupIfDoc(string backupFolderName)
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                    return;

                string originalPath = doc.FullName;
                string dir = Path.GetDirectoryName(originalPath);
                string name = Path.GetFileNameWithoutExtension(originalPath);

                // 末尾が "_backup" または "_backupx" の場合は何もしない
                if (name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase) ||
                    name.EndsWith("_backupx", StringComparison.OrdinalIgnoreCase))
                    return;

                // 互換モードか拡張子が .doc の場合にバックアップ
                bool isDocExtension = Path.GetExtension(originalPath).Equals(".doc", StringComparison.OrdinalIgnoreCase);
                bool isCompatibilityMode = false;
                try
                {
                    // CompatibilityMode: 15=docx, 14/12/11=doc互換
                    isCompatibilityMode = (doc.CompatibilityMode != 15);
                }
                catch
                {
                    // CompatibilityMode プロパティが使えない場合は無視
                }

                if (!isDocExtension && !isCompatibilityMode)
                    return;

                // backupフォルダのパスを作成（引数で指定されたフォルダ名を使用）
                string backupDir = Path.Combine(dir, backupFolderName);

                // backupフォルダがなければ作成
                if (!Directory.Exists(backupDir))
                {
                    Directory.CreateDirectory(backupDir);
                }

                string backupName = name + "_backupx.docx";
                string backupPath = Path.Combine(backupDir, backupName);

                if (File.Exists(backupPath))
                    return;

                // .docx 形式で保存
                doc.SaveAs2(backupPath, Word.WdSaveFormat.wdFormatXMLDocument);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"docxバックアップ作成時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void OverwriteDocument()
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                    return;

                string originalPath = doc.FullName;
                string dir = Path.GetDirectoryName(originalPath);
                string name = Path.GetFileNameWithoutExtension(originalPath);
                string ext = Path.GetExtension(originalPath);

                // ファイル名の末尾が "_backup" でなければ何もしない
                if (!name.EndsWith("_backup", StringComparison.OrdinalIgnoreCase))
                    return;

                // 現在のフォルダを対象にする
                string targetName = name.Substring(0, name.Length - "_backup".Length) + ext;
                string targetPath = Path.Combine(dir, targetName);

                // sampleファイルが存在すれば削除
                if (File.Exists(targetPath))
                {
                    File.Delete(targetPath);
                }

                // 自分自身をターゲットにコピー
                File.Copy(originalPath, targetPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"ファイルの上書き時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // _backupx.docx ファイルが存在すれば zip に変換する
        private void ConvertBackupxToZip(string backupFolderName)
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                    return;

                string dir = Path.GetDirectoryName(doc.FullName);
                string name = Path.GetFileNameWithoutExtension(doc.FullName);

                // backupFolderNameフォルダ配下の _backupx.docx を参照
                string backupDir = Path.Combine(dir, backupFolderName);
                string backupxName = name + "_backupx.docx";
                string backupxPath = Path.Combine(backupDir, backupxName);

                if (!File.Exists(backupxPath))
                    return;

                string zipName = name + "_backupx.zip";
                string zipPath = Path.Combine(backupDir, zipName);

                // 既存のzipは上書き
                if (File.Exists(zipPath))
                    File.Delete(zipPath);

                // .docxはzip形式なので、そのままコピーして拡張子をzipにするだけでOK
                File.Copy(backupxPath, zipPath);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"zip変換時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // _backupx.zip ファイルを解凍する
        private void UnzipBackupxZip(string backupFolderName)
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                    return;

                string dir = Path.GetDirectoryName(doc.FullName);
                string name = Path.GetFileNameWithoutExtension(doc.FullName);

                // zipファイルも backupFolderName フォルダ内にあるとする
                string backupDir = Path.Combine(dir, backupFolderName);
                string zipName = name + "_backupx.zip";
                string zipPath = Path.Combine(backupDir, zipName);

                if (!File.Exists(zipPath))
                    return;

                // 展開先ディレクトリも backupFolderName
                string extractDir = backupDir;

                // 既存の展開先フォルダがあれば削除
                if (Directory.Exists(extractDir))
                    Directory.Delete(extractDir, true);

                ZipFile.ExtractToDirectory(zipPath, extractDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"zip解凍時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        // backupFolderName/word/media フォルダをアクティブドキュメントのあるフォルダに移動
        private void MoveUnzippedMediaFolderToDocumentDirectory(string backupFolderName)
        {
            try
            {
                var doc = this.Application.ActiveDocument;
                if (doc == null || string.IsNullOrEmpty(doc.FullName))
                    return;

                string docDir = Path.GetDirectoryName(doc.FullName);

                // backupFolderName/word/media を参照
                string backupDir = Path.Combine(docDir, backupFolderName);
                string srcMediaDir = Path.Combine(backupDir, "word", "media");
                string destMediaDir = Path.Combine(docDir, "media");

                if (!Directory.Exists(srcMediaDir))
                    return;

                // 既存の media フォルダがあれば削除
                if (Directory.Exists(destMediaDir))
                    Directory.Delete(destMediaDir, true);

                // media フォルダを移動
                Directory.Move(srcMediaDir, destMediaDir);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"mediaフォルダ移動時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }
    }
}
