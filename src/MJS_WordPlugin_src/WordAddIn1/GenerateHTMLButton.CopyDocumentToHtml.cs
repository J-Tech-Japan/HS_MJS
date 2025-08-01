using System.IO;
using Word = Microsoft.Office.Interop.Word;
using Application = System.Windows.Forms.Application;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        //ドキュメントを一時 HTML 用にコピー（旧版コード）
        private Word.Document CopyDocumentToHtml(Word.Application application, StreamWriter log)
        {
            ClearClipboardSafely();
            Application.DoEvents();
            application.Selection.WholeStory();
            application.Selection.Copy();
            Application.DoEvents();
            application.Selection.Collapse(Word.WdCollapseDirection.wdCollapseStart);
            Application.DoEvents();
            Word.Document docCopy = application.Documents.Add();
            Application.DoEvents();
            docCopy.TrackRevisions = false;
            docCopy.AcceptAllRevisions();
            docCopy.Select();
            Application.DoEvents();
            application.Selection.PasteAndFormat(Word.WdRecoveryType.wdUseDestinationStylesRecovery);
            Application.DoEvents();
            ClearClipboardSafely();
            Application.DoEvents();
            log.WriteLine("Number of sections: " + docCopy.Sections.Count);
            return docCopy;
        }


        //private Word.Document CopyDocumentToHtml(Word.Application application, StreamWriter log)
        //{
        //    // 元ドキュメントの全範囲を取得
        //    Document srcDoc = application.ActiveDocument;
        //    Range srcRange = srcDoc.Content;

        //    // 新規ドキュメントを作成
        //    Document docCopy = application.Documents.Add();
        //    docCopy.TrackRevisions = false;

        //    // 元ドキュメントの全範囲をコピー＆ペースト（フィールドを保持）
        //    srcRange.Copy();
        //    Range destRange = docCopy.Content;
        //    destRange.Paste();

        //    Application.DoEvents();

        //    // クリップボードをクリア（任意）
        //    ClearClipboardSafely();

        //    log.WriteLine("Number of sections: " + docCopy.Sections.Count);
        //    return docCopy;
        //}
    }
}
