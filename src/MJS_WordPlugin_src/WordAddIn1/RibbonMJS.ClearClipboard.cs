using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void ClearClipboardSafely()
        {
            Thread staThread = new Thread(() =>
            {
                const int maxRetryCount = 5; // 最大リトライ回数
                const int retryDelay = 100; // リトライ間隔（ミリ秒）

                for (int attempt = 0; attempt < maxRetryCount; attempt++)
                {
                    try
                    {
                        // クリップボードをクリア
                        Clipboard.Clear();

                        // 空のデータを設定して確実にクリア
                        Clipboard.SetDataObject(new DataObject());

                        // 成功した場合はループを抜ける
                        return;
                    }
                    catch (COMException ex)
                    {
                        // ログに記録
                        Debug.WriteLine($"クリップボードのクリアに失敗しました (試行 {attempt + 1}/{maxRetryCount}): {ex.Message}");

                        // 最大リトライ回数に達した場合は例外を再スロー
                        if (attempt == maxRetryCount - 1)
                        {
                            throw;
                        }
                    }
                    catch (Exception ex)
                    {
                        // その他の例外もログに記録
                        Debug.WriteLine($"予期しないエラーが発生しました: {ex.Message}");
                        throw;
                    }

                    // リトライ間隔を待機
                    Thread.Sleep(retryDelay);
                }
            });

            staThread.SetApartmentState(ApartmentState.STA);
            staThread.Start();
            staThread.Join(); // スレッド終了を待機
        }
    }
}
