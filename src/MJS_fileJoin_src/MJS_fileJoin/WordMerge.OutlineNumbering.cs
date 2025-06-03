using System.Collections.Generic;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.Word;
using MJS_fileJoin;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        private void ApplyOutlineNumbering(Word.Application objApp, Word.Document objDocLast, MainForm fm)
        {
            List<string> styleNames = new List<string>();
            styleNames.Add("MJS_章扉-タイトル");
            styleNames.Add("見出し 1,MJS_見出し 1");
            styleNames.Add("見出し 2,MJS_見出し 2");
            styleNames.Add("見出し 3,MJS_見出し 3");

            // スタイルのアウトライン番号を設定
            SetOutlineNumberingFormat(objApp, objDocLast, fm);

            int first = 0;
            int second = 0;
            int third = 0;
            int fourth = 0;

            for (int i = 1; i <= objDocLast.ListParagraphs.Count; i++)
            {
                fm.progressBar1.Increment(1);
                if (!Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"第.*?章") && !Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"\d\.\d")) continue;
                if (Regex.IsMatch(objDocLast.ListParagraphs[i].Range.ListFormat.ListString, @"第.*?章"))
                {
                    first++;
                    second = 0;
                    third = 0;
                    fourth = 0;
                    if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != first)
                        objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                }
                else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 2)
                {
                    second++;
                    third = 0;
                    fourth = 0;
                    if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != second)
                        objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                }
                else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 3)
                {
                    third++;
                    fourth = 0;
                    if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != third)
                        objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                }
                else if (objDocLast.ListParagraphs[i].Range.ListFormat.ListLevelNumber == 4)
                {
                    fourth++;
                    if (objDocLast.ListParagraphs[i].Range.ListFormat.ListValue != fourth)
                        objDocLast.ListParagraphs[i].Range.ListFormat.ApplyListTemplateWithLevel(objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7], true, Word.WdListApplyTo.wdListApplyToWholeList, Word.WdDefaultListBehavior.wdWord10ListBehavior);
                }
            }
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