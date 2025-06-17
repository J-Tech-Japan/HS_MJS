using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // ヘルパーメソッド: 処理停止時の処理
        private void HandleProcessHalt()
        {
            var application = Globals.ThisAddIn.Application;
            application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            button3.Enabled = false;
            MessageBox.Show("スタイルチェックが停止しました。\r\nチェック済み項目は全て破棄されます。", "スタイルチェック停止", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        // ヘルパーメソッド: 処理成功時の処理
        private void HandleProcessSuccess(Word.Document document)
        {
            var application = Globals.ThisAddIn.Application;
            document.ShowRevisions = false;
            MessageBox.Show("スタイルチェックOKです。\r\n「HTML出力」ボタンをクリックするとHTMLが出力されます。", "スタイルチェックOK", MessageBoxButtons.OK, MessageBoxIcon.Asterisk);

            button3.Enabled = true;
            checkOK = true;
            application.WindowSelectionChange += new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            Application.DoEvents();
        }

        // ヘルパーメソッド: 処理失敗時の処理
        private void HandleProcessFailure()
        {
            var application = Globals.ThisAddIn.Application;
            application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);
            button3.Enabled = false;
            MessageBox.Show("スタイルチェックNGです。\r\n「校閲」タブ-「コメント」-「次へ」ボタンで\r\n使用できない書式を確認できます。", "スタイルチェックNG", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
