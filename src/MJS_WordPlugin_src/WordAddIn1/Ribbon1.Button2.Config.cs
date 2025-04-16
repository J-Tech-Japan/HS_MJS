using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
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
