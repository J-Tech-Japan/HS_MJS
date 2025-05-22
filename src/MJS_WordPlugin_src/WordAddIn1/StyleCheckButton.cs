using Microsoft.Office.Tools.Ribbon;
using System.IO;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;


// リファクタリング済
namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void StyleCheckButton(object sender, RibbonControlEventArgs e)
        {
            var application = Globals.ThisAddIn.Application;
            var activeDocument = application.ActiveDocument;
            var selection = application.Selection;

            // ドキュメント変更イベントとウィンドウ選択変更イベントのハンドラーを解除
            application.DocumentChange -= new Word.ApplicationEvents4_DocumentChangeEventHandler(Application_DocumentChange);
            application.WindowSelectionChange -= new Word.ApplicationEvents4_WindowSelectionChangeEventHandler(Application_WindowSelectionChange);

            // スタイル名を格納するリストを初期化
            List<string> styleList = new List<string>();

            // 現在の選択範囲の開始位置と終了位置を取得
            int selectionStart = selection.Start;
            int selectionEnd = selection.End;

            // ドキュメント全体を選択するためにカーソルを末尾と先頭に移動
            selection.EndKey(Word.WdUnits.wdStory);
            Application.DoEvents();
            selection.HomeKey(Word.WdUnits.wdStory);
            Application.DoEvents();

            // アタッチされているテンプレートファイルのパスを取得
            string attachedTemplateFile = Path.Combine(
                activeDocument.get_AttachedTemplate().Path,
                activeDocument.get_AttachedTemplate().Name
            );

            // アタッチされているテンプレートファイルを開く
            Word.Document templateDocument = application.Documents.Open(attachedTemplateFile);

            // ログファイルにテンプレートファイルのパスを記録
            using (StreamWriter log = new StreamWriter(activeDocument.Path + "\\log.txt", true, Encoding.UTF8))
            {
                log.WriteLine("Attached template file: " + attachedTemplateFile);
            }

            // スタイルのリストを初期化
            button3.Enabled = false;

            // 画面更新を停止
            application.ScreenUpdating = false;

            // テンプレート内のスタイルを確認し、"MJS"を含むスタイル名をリストに追加
            foreach (Word.Style stl in templateDocument.Styles)
            {
                if (stl.NameLocal.Contains("MJS"))
                    styleList.Add(stl.NameLocal);
            }

            // テンプレートファイルを閉じる
            templateDocument.Close();

            // ドキュメントの最初のページに移動し、すべての変更履歴を承認
            selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToFirst);
            activeDocument.Revisions.AcceptAll();

            // ドキュメントの表示設定を変更（コメントや変更履歴の表示を制御）
            ConfigDocumentDisplay();

            // 処理フラグを初期化
            bool hasInvalidStyles = false;

            // ドキュメントのコメントを削除
            DeleteInvalidComments(activeDocument.Comments);

            // ドキュメントの先頭位置を選択
            activeDocument.Range(0, 0).Select();

            // 検索条件を設定
            ConfigSearchParameters();

            Application.DoEvents();

            // 検索を実行し、該当箇所にコメントを追加
            while (selection.Find.Execute())
            {
                selection.Range.Comments.Add(selection.Range,
                    "【改行なしスペース】\r\n使用できない文字列です。");
                hasInvalidStyles = true;
            }

            Application.DoEvents();

            // 図形が行内配置になっていない場合は校閲コメントを追加
            AddCommentForNonInlineShape(activeDocument, ref hasInvalidStyles);

            // 画面更新を再開
            application.ScreenUpdating = true;

            Application.DoEvents();

            // プログレスバーを表示
            ProgressBar.Show();

            // プログレスバーの最大値を設定（段落数）
            ProgressBar.SetProgressBar(activeDocument.Paragraphs.Count);

            // プログレスバーの進捗を初期化
            int progress = 0;

            // 処理時間を計測するためのストップウォッチを開始
            Stopwatch sw = Stopwatch.StartNew();

            // ドキュメントの最初のページに移動
            selection.GoTo(Word.WdGoToItem.wdGoToPage, Word.WdGoToDirection.wdGoToFirst);

            // 処理フラグを初期化
            bool isProcessing = false;
            bool isProcessHalted = false;

            // 段落を処理するメソッドを呼び出す
            ProcessParagraphs(activeDocument, styleList, sw, ref progress, ref isProcessing, ref isProcessHalted, ref hasInvalidStyles);

            // ストップウォッチを停止し、リソースを解放
            sw.Stop();
            sw = null;

            // 選択範囲を元に戻す
            RestoreSelection(selectionStart, selectionEnd);

            // 処理結果に応じた後処理
            if (isProcessHalted)
            {
                HandleProcessHalt();
            }
            else if (!hasInvalidStyles)
            {
                HandleProcessSuccess(activeDocument);
            }
            else
            {
                HandleProcessFailure();
            }

            // プログレスバーを閉じる
            ProgressBar.Close();
            ProgressBar.mInstance = null;
        }
    }
}
