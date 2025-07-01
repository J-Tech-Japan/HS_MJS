# SaveBackupOfActiveDocument

## 概要
ThisAddIn.cs の `SaveBackupOfActiveDocument` メソッドは、アクティブなWord文書のバックアップファイル（_backup付き）を自動的に作成する処理を行います。

## 詳細な処理手順
1. アクティブな文書（ActiveDocument）が存在し、かつファイルパスが空でない場合のみ処理を続行します。
2. ファイル名の末尾が「_backup」であれば何もせず終了します（バックアップファイル自身はバックアップしない）。
3. バックアップファイル名（元のファイル名＋_backup）を生成します。
4. バックアップファイルが既に存在する場合は何もせず終了します。
5. バックアップファイルが存在しない場合は、元のファイルをコピーしてバックアップファイルを作成します。

## 目的
- アクティブな文書のバックアップを自動的に作成し、元ファイルの保護や復元を容易にします。
- 編集前の状態を保持したい場合や、誤操作によるデータ損失を防ぐために有効です。

## 注意点
- 既にバックアップファイルが存在する場合は新たに作成しません。
- バックアップ対象は「_backup」で終わらないファイルのみです。
- ファイル操作時に例外が発生した場合はエラーメッセージが表示されます。

```CSharp
// ThisAddIn.cs
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
        }
    }
    catch (Exception ex)
    {
        MessageBox.Show($"バックアップ保存時にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Warning);
    }
}
```
