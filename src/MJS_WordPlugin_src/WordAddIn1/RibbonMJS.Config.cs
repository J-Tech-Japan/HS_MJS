using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // ファイル名形式の規定
        private const string FileNamePattern = @"^[A-Z]{3}(_[^_]*?){2}\.docx*$";

        // 一般的なエラーメッセージ
        private const string ErrMsg = "エラーが発生しました。";

        // 開いているファイル名が正しくない場合に表示するメッセージ
        private const string ErrMsgInvalidFileName = "開いているWordのファイル名が正しくありません。\r\n下記の例を参考にファイル名を変更してください。\r\n\r\n(英半角大文字3文字)_(製品名)_(バージョンなど自由付加).doc\r\n\r\n例):「AAA_製品A_r1.doc」";
        private const string ErrMsgFileNameRule = "ファイル命名規則エラー";

        private const string ErrMsgTmpDocOpen = "同階層のtmp.docが開かれています。\r\ntmp.docを閉じてから実行してください。";
        private const string ErrMsgFile = "ファイルエラー";

        // HTML出力が成功した場合に表示するメッセージ
        private const string MsgHtmlOutputSuccess1 = "\r\nにHTMLが出力されました。\r\n出力したHTMLをブラウザで表示しますか？";
        private const string MsgHtmlOutputSuccess2 = "HTML出力成功";

        // HTML出力が失敗した場合に表示するメッセージ
        private const string ErrMsgHtmlOutputFailure1 = "HTMLの出力に失敗しました。";
        private const string ErrMsgHtmlOutputFailure2 = "HTML出力失敗。";


        // ヘルパーメソッド: ドキュメントの表示設定
        private void ConfigDocumentDisplay()
        {
            var view = Globals.ThisAddIn.Application.ActiveDocument.ActiveWindow.View;
            view.ShowRevisionsAndComments = true;
            view.ShowInkAnnotations = false;
            view.ShowComments = true;
            view.ShowInsertionsAndDeletions = false;
            view.ShowFormatChanges = false;
        }

        // ヘルパーメソッド: 検索条件の設定
        private void ConfigSearchParameters()
        {
            var find = Globals.ThisAddIn.Application.Selection.Find;
            // 検索条件を設定
            find.ClearFormatting();
            find.Replacement.ClearFormatting();
            find.Text = "^s";
            find.Forward = true;
            find.Wrap = Word.WdFindWrap.wdFindStop;
            find.Format = false;
            find.MatchCase = false;
            find.MatchWholeWord = false;
            find.MatchByte = false;
            find.MatchAllWordForms = false;
            find.MatchSoundsLike = false;
            find.MatchWildcards = false;
            find.MatchFuzzy = false;
        }
    }
}
