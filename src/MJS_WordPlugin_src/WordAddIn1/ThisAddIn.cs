// ThisAddIn.cs

using System;

namespace WordAddIn1
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // EnhMetaFileBitsを使用した画像・キャンバス抽出（アクティブドキュメントから直接、テキスト情報付き） ***
            //ExtractImagesAndCanvasFromActiveDocumentWithText();
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            try
            {
                // アドイン終了時に全ての画像マーカーを削除
                //if (Globals.ThisAddIn.Application != null && Globals.ThisAddIn.Application.Documents.Count > 0)
                //{
                //    RemoveMarkersFromActiveDocument();
                //}
                
                // 必要に応じてリソースを解放
                if (Globals.ThisAddIn.Application != null)
                {
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(Globals.ThisAddIn.Application);
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"アドイン終了時にエラーが発生しました: {ex.Message}");
                // シャットダウン時はメッセージボックスを表示しない（Wordの終了を妨げる可能性があるため）
            }
        }

        #region VSTO で生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// メソッドの内容をコードエディターで変更しないでください。
        /// </summary>
        private void InternalStartup()
        {
            Startup += new System.EventHandler(ThisAddIn_Startup);
            Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}