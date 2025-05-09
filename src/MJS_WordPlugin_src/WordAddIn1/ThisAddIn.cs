using System;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            try
            {
                string version = Globals.ThisAddIn.Application.Version; // メジャーバージョン
                string buildNumber = Globals.ThisAddIn.Application.Build.ToString(); // ビルド番号
                string fullVersion = $"{version}.{buildNumber}";

                //Globals.ThisAddIn.Application.Caption = $"Word Add-In - Version: {fullVersion}";
                Globals.ThisAddIn.Application.Caption = $"Word Add-In - Version: {buildNumber}";
            }
            catch (Exception ex)
            {
                MessageBox.Show($"アドインの初期化中にエラーが発生しました: {ex.Message}", "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
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