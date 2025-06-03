using Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // Word.Application、Document の生成・初期設定
        private (Application, Document) CreateWordDocument(string filePath, object objMissing)
        {
            var app = new Application();
            app.DisplayAlerts = WdAlertLevel.wdAlertsNone;
            app.Options.CheckGrammarAsYouType = false;
            app.Options.CheckGrammarWithSpelling = false;
            app.Options.CheckSpellingAsYouType = false;
            app.Options.ShowReadabilityStatistics = false;
            app.Visible = false;

            object objFilePath = filePath;
            var doc = app.Documents.Open(
                ref objFilePath,    // FileName
                ref objMissing,     // ConfirmVersions
                ref objMissing,     // ReadOnly
                ref objMissing,     // AddToRecentFiles
                ref objMissing,     // PasswordDocument
                ref objMissing,     // PasswordTemplate
                ref objMissing,     // Revert
                ref objMissing,     // WritePasswordDocument
                ref objMissing,     // WritePasswordTemplate
                ref objMissing,     // Format
                ref objMissing,     // Encoding
                ref objMissing,     // Visible
                ref objMissing,     // OpenAndRepair
                ref objMissing,     // DocumentDirection
                ref objMissing,     // NoEncodingDialog
                ref objMissing      // XMLTransform
            );

            return (app, doc);
        }
    }
}
