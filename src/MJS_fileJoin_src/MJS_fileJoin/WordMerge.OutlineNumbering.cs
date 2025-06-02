using Microsoft.Office.Interop.Word;
using MJS_fileJoin;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // Word.Application、Document の生成・初期設定
        private (Application, Document) CreateAndOpenWordDocument(string filePath, object objMissing)
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

        // Word文書のアウトライン番号（章番号や見出し番号）の書式やスタイルを設定
        public void SetOutlineNumberingFormat(Application objApp, Document objDocLast, MainForm fm)
        {
            // 章扉-タイトル（第%1章）
            var level1 = objApp.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[1];
            level1.NumberFormat = "第%1章";
            level1.TrailingCharacter = WdTrailingCharacter.wdTrailingNone;
            level1.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            level1.NumberPosition = objApp.MillimetersToPoints(0F);
            level1.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            level1.TextPosition = objApp.MillimetersToPoints(5.0F);
            level1.ResetOnHigher = 0;
            level1.StartAt = 1;
            level1.Font.Bold = 1;
            level1.Font.Italic = 0;
            level1.Font.Color = WdColor.wdColorAutomatic;
            level1.Font.Size = 60;
            level1.Font.Name = "メイリオ";
            level1.LinkedStyle = "MJS_章扉-タイトル";

            // 見出し1（%1.%2）
            var level2 = objApp.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[2];
            level2.NumberFormat = "%1.%2";
            level2.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            level2.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            level2.NumberPosition = objApp.MillimetersToPoints(1.5F);
            level2.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            level2.TextPosition = objApp.MillimetersToPoints(20.0F);
            level2.TabPosition = objApp.MillimetersToPoints(20.0F);
            level2.ResetOnHigher = 1;
            level2.StartAt = 1;
            level2.Font.Bold = 1;
            level2.Font.Italic = 0;
            level2.Font.Color = WdColor.wdColorAutomatic;
            level2.Font.Size = 16;
            level2.Font.Name = "メイリオ";
            level2.LinkedStyle = "見出し 1";

            // 見出し2（%1.%2.%3）
            var level3 = objApp.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[3];
            level3.NumberFormat = "%1.%2.%3";
            level3.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            level3.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            level3.NumberPosition = objApp.MillimetersToPoints(0.0F);
            level3.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            level3.TextPosition = objApp.MillimetersToPoints(20.0F);
            level3.TabPosition = objApp.MillimetersToPoints(20.0F);
            level3.ResetOnHigher = 2;
            level3.StartAt = 1;
            level3.Font.Bold = 1;
            level3.Font.Italic = 0;
            level3.Font.Color = (WdColor)(31 + 0x100 * 73 + 0x10000 * 125); // Color.FromArgb(31, 73, 125)
            level3.Font.Size = 14;
            level3.Font.Name = "メイリオ";
            level3.LinkedStyle = "見出し 2";

            // 見出し3（%1.%2.%3.%4）
            var level4 = objApp.ListGalleries[WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels[4];
            level4.NumberFormat = "%1.%2.%3.%4";
            level4.TrailingCharacter = WdTrailingCharacter.wdTrailingTab;
            level4.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic;
            level4.NumberPosition = objApp.MillimetersToPoints(7.0F);
            level4.Alignment = WdListLevelAlignment.wdListLevelAlignLeft;
            level4.TextPosition = objApp.MillimetersToPoints(28.0F);
            level4.TabPosition = objApp.MillimetersToPoints(28.0F);
            level4.ResetOnHigher = 3;
            level4.StartAt = 1;
            level4.Font.Name = "メイリオ";
            level4.LinkedStyle = "見出し 3";

            fm.label10.Text = "見出し番号修正中...";
            fm.progressBar1.Maximum = objDocLast.ListParagraphs.Count;
            fm.progressBar1.Value = 1;
        }
    }
}