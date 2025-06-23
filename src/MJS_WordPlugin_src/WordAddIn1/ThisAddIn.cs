using System;
using System.IO;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            //ConvertHyperlinkRefFieldsToRefFields();
            //ConvertAbsoluteHyperlinksToFilenameOnly();
            //try
            //{
            //    string version = Globals.ThisAddIn.Application.Version; // メジャーバージョン
            //    string buildNumber = Globals.ThisAddIn.Application.Build.ToString(); // ビルド番号
            //    string fullVersion = $"{version}.{buildNumber}";

            //    //Globals.ThisAddIn.Application.Caption = $"Word Add-In - Version: {fullVersion}";
            //    Globals.ThisAddIn.Application.Caption = $"Word Add-In - Version: {buildNumber}";
            //}
            //catch (Exception ex)
            //{
            //    MessageBox.Show($"アドインの初期化中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            //}
        }

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

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
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

        // HYPERLINKフィールドをREFフィールドに変換
        private void ConvertHyperlinkRefFieldsToRefFields()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;
            for (int i = doc.Fields.Count; i >= 1; i--)
            {
                var field = doc.Fields[i];
                try
                {
                    if (field.Type == Word.WdFieldType.wdFieldHyperlink)
                    {
                        string fieldCode = field.Code.Text;
                        var refIndex = fieldCode.IndexOf("_Ref", StringComparison.OrdinalIgnoreCase);
                        if (refIndex >= 0)
                        {
                            var refStart = refIndex;
                            var refEnd = refStart;
                            while (refEnd < fieldCode.Length && (char.IsLetterOrDigit(fieldCode[refEnd]) || fieldCode[refEnd] == '_'))
                                refEnd++;
                            string refId = fieldCode.Substring(refStart, refEnd - refStart);

                            // フィールド全体を選択して置換
                            field.Select();
                            var sel = doc.Application.Selection;
                            // 既存フィールドの範囲を取得
                            var rng = sel.Range;
                            // 既存フィールドを削除
                            field.Delete();
                            // REFフィールドを挿入
                            doc.Fields.Add(rng, Word.WdFieldType.wdFieldRef, $"{refId} \\h");
                            // 選択を解除
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"フィールド変換エラー: {ex.Message}");
                }
            }
            // 変換後にフィールドを更新
            doc.Fields.Update();
        }

        private void ConvertAbsoluteHyperlinksToFilenameOnly()
        {
            var doc = Globals.ThisAddIn.Application.ActiveDocument;

            for (int i = doc.Fields.Count; i >= 1; i--)
            {
                var field = doc.Fields[i];
                try
                {
                    if (field.Type == Word.WdFieldType.wdFieldHyperlink)
                    {
                        string fieldCode = field.Code.Text.Trim();
                        // HYPERLINK "file:///C:\..." の形式を抽出
                        var match = System.Text.RegularExpressions.Regex.Match(
                            fieldCode,
                            @"HYPERLINK\s+""file:///([A-Za-z]:\\[^""]+)""",
                            System.Text.RegularExpressions.RegexOptions.IgnoreCase);

                        if (match.Success)
                        {
                            string absPath = match.Groups[1].Value;
                            string fileName = Path.GetFileName(absPath);

                            // フィールド全体を選択して置換
                            field.Select();
                            var sel = doc.Application.Selection;
                            var rng = sel.Range;
                            field.Delete();
                            // ファイル名のみでHYPERLINKフィールドを再挿入
                            doc.Fields.Add(rng, Word.WdFieldType.wdFieldHyperlink, $"\"{fileName}\"");
                            sel.Collapse(Word.WdCollapseDirection.wdCollapseEnd);
                        }
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"HYPERLINK変換エラー: {ex.Message}");
                }
            }
            doc.Fields.Update();
        }





        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}