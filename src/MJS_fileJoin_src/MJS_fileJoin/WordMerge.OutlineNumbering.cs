using System.Drawing;
using System.Text.RegularExpressions;
using MJS_fileJoin;
using Word = Microsoft.Office.Interop.Word;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // Wordのアウトライン番号書式を設定
        private void SetOutlineNumberingFormat(Word.Application objApp, Color mycolor)
        {
            var listLevels = objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7].ListLevels;

            // レベル1
            listLevels[1].NumberFormat = "第%1章";
            listLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            listLevels[1].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevels[1].NumberPosition = objApp.MillimetersToPoints(0F);
            listLevels[1].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            listLevels[1].TextPosition = objApp.MillimetersToPoints(5.0F);
            listLevels[1].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingNone;
            listLevels[1].ResetOnHigher = 0;
            listLevels[1].StartAt = 1;
            listLevels[1].Font.Bold = 1;
            listLevels[1].Font.Italic = 0;
            listLevels[1].Font.Color = Word.WdColor.wdColorAutomatic;
            listLevels[1].Font.Size = 60;
            listLevels[1].Font.Name = "メイリオ";
            listLevels[1].LinkedStyle = "MJS_章扉-タイトル";

            // レベル2
            listLevels[2].NumberFormat = "%1.%2";
            listLevels[2].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            listLevels[2].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevels[2].NumberPosition = objApp.MillimetersToPoints(1.5F);
            listLevels[2].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            listLevels[2].TextPosition = objApp.MillimetersToPoints(20.0F);
            listLevels[2].TabPosition = objApp.MillimetersToPoints(20.0F);
            listLevels[2].ResetOnHigher = 1;
            listLevels[2].StartAt = 1;
            listLevels[2].Font.Bold = 1;
            listLevels[2].Font.Italic = 0;
            listLevels[2].Font.Color = Word.WdColor.wdColorAutomatic;
            listLevels[2].Font.Size = 16;
            listLevels[2].Font.Name = "メイリオ";
            listLevels[2].LinkedStyle = "見出し 1";

            // レベル3
            listLevels[3].NumberFormat = "%1.%2.%3";
            listLevels[3].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            listLevels[3].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevels[3].NumberPosition = objApp.MillimetersToPoints(0.0F);
            listLevels[3].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            listLevels[3].TextPosition = objApp.MillimetersToPoints(20.0F);
            listLevels[3].TabPosition = objApp.MillimetersToPoints(20.0F);
            listLevels[3].ResetOnHigher = 2;
            listLevels[3].StartAt = 1;
            listLevels[3].Font.Bold = 1;
            listLevels[3].Font.Italic = 0;
            listLevels[3].Font.Color = (Word.WdColor)(mycolor.R + 0x100 * mycolor.G + 0x10000 * mycolor.B);
            listLevels[3].Font.Size = 14;
            listLevels[3].Font.Name = "メイリオ";
            listLevels[3].LinkedStyle = "見出し 2";

            // レベル4
            listLevels[4].NumberFormat = "%1.%2.%3.%4";
            listLevels[4].TrailingCharacter = Word.WdTrailingCharacter.wdTrailingTab;
            listLevels[4].NumberStyle = Word.WdListNumberStyle.wdListNumberStyleArabic;
            listLevels[4].NumberPosition = objApp.MillimetersToPoints(7.0F);
            listLevels[4].Alignment = Word.WdListLevelAlignment.wdListLevelAlignLeft;
            listLevels[4].TextPosition = objApp.MillimetersToPoints(28.0F);
            listLevels[4].TabPosition = objApp.MillimetersToPoints(28.0F);
            listLevels[4].ResetOnHigher = 3;
            listLevels[4].StartAt = 1;
            listLevels[4].Font.Name = "メイリオ";
            listLevels[4].LinkedStyle = "見出し 3";
        }

        // 段落のアウトライン番号を修正
        private void FixOutlineNumbering(Word.Document objDocLast, Word.Application objApp, MainForm form)
        {
            int paragraphCount = objDocLast.ListParagraphs.Count;
            
            using (var progress = Utils.BeginProgress(form, "見出し番号修正中...", paragraphCount))
            {
                // ListTemplateを事前にキャッシュ
                var listTemplate = objApp.ListGalleries[Word.WdListGalleryType.wdOutlineNumberGallery].ListTemplates[7];
                
                // 正規表現を事前にコンパイル
                var chapterRegex = new Regex(@"第.*?章", RegexOptions.Compiled);
                var sectionRegex = new Regex(@"\d\.\d", RegexOptions.Compiled);

                int first = 0;
                int second = 0;
                int third = 0;
                int fourth = 0;
                
                for (int i = 1; i <= paragraphCount; i++)
                {
                    // UI更新頻度を調整（10段落ごと、または最後）
                    if (i % 10 == 0 || i == paragraphCount)
                    {
                        progress.SetValue(i);
                    }

                    var paragraph = objDocLast.ListParagraphs[i];
                    var listFormat = paragraph.Range.ListFormat;
                    
                    // COM オブジェクトへのアクセスを最小限に
                    string listString = listFormat.ListString;
                    
                    bool isChapter = chapterRegex.IsMatch(listString);
                    bool hasSection = sectionRegex.IsMatch(listString);
                    
                    if (!isChapter && !hasSection) continue;

                    int levelNumber = listFormat.ListLevelNumber;
                    int currentValue = listFormat.ListValue;
                    bool needsUpdate = false;
                    
                    if (isChapter)
                    {
                        first++;
                        second = 0;
                        third = 0;
                        fourth = 0;
                        needsUpdate = (currentValue != first);
                    }
                    else if (levelNumber == 2)
                    {
                        second++;
                        third = 0;
                        fourth = 0;
                        needsUpdate = (currentValue != second);
                    }
                    else if (levelNumber == 3)
                    {
                        third++;
                        fourth = 0;
                        needsUpdate = (currentValue != third);
                    }
                    else if (levelNumber == 4)
                    {
                        fourth++;
                        needsUpdate = (currentValue != fourth);
                    }

                    // 必要な場合のみテンプレートを適用
                    if (needsUpdate)
                    {
                        listFormat.ApplyListTemplateWithLevel(
                            listTemplate,
                            true,
                            Word.WdListApplyTo.wdListApplyToWholeList,
                            Word.WdDefaultListBehavior.wdWord10ListBehavior
                        );
                    }
                }
                
                progress.Complete();
            }
        }
    }
}