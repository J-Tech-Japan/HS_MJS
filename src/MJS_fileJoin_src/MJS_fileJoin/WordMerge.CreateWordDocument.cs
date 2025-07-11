using System;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        private void InitializeWordAndOpenDocument(
            string strOrgDoc,
            ref Word.Application objApp,
            ref Word.Document objDocLast)
        {
            try
            {
                objApp = new Word.Application
                {
                    Visible = false,
                    DisplayAlerts = Word.WdAlertLevel.wdAlertsNone
                };

                objApp.Options.CheckGrammarAsYouType = false;
                objApp.Options.CheckGrammarWithSpelling = false;
                objApp.Options.CheckSpellingAsYouType = false;
                objApp.Options.ShowReadabilityStatistics = false;

                objDocLast = objApp.Documents.Open(
                    strOrgDoc,    // FileName
                    Type.Missing, // ConfirmVersions
                    Type.Missing, // ReadOnly
                    Type.Missing, // AddToRecentFiles
                    Type.Missing, // PasswordDocument
                    Type.Missing, // PasswordTemplate
                    Type.Missing, // Revert
                    Type.Missing, // WritePasswordDocument
                    Type.Missing, // WritePasswordTemplate
                    Type.Missing, // Format
                    Type.Missing, // Encoding
                    Type.Missing, // Visible
                    Type.Missing, // OpenAndRepair
                    Type.Missing, // DocumentDirection
                    Type.Missing, // NoEncodingDialog
                    Type.Missing  // XMLTransform
                );
            }
            catch (Exception ex)
            {
                MessageBox.Show("Wordの起動またはファイルオープンに失敗しました: " + ex.Message, "エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // 必要に応じてリソース解放
                if (objApp != null)
                {
                    objApp.Quit(Type.Missing, Type.Missing, Type.Missing);
                    objApp = null;
                }
                throw;
            }
        }
    }
}
