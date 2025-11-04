using System.Linq;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        private void UpdateChapterFrontNumbers(Word.Document objDocLast, MainForm fm)
        {
            const string targetStyleName = "MJS_章扉-タイトル";
            const string shouTobiraStyleName = "MJS_章扉-目次1";

            // スタイルの存在確認を一度だけ実行
            bool hasTargetStyle = objDocLast.Styles.Cast<Word.Style>()
                .Any(s => s.NameLocal == targetStyleName);

            if (!hasTargetStyle)
            {
                // スタイルが存在しない場合は処理をスキップ
                return;
            }

            object shouTobiraStyle = shouTobiraStyleName;
            int allChap = objDocLast.Sections.Count;

            // 正規表現を事前にコンパイル
            Regex numberRegex = new Regex(@"^\d+?", RegexOptions.Compiled);

            for (int i = 1; i <= allChap; i++)
            {
                try
                {
                    Word.Range sectionRange = objDocLast.Sections[i].Range;
                    Word.Paragraphs paragraphs = sectionRange.Paragraphs;
                    int paragraphCount = paragraphs.Count;

                    int shou = 0;
                    bool foundChapterTitle = false;

                    // 段落を一度だけ列挙
                    for (int p = 1; p <= paragraphCount - 1; p++)
                    {
                        Word.Paragraph para = paragraphs[p];
                        string styleName = para.get_Style().NameLocal.Trim();

                        if (styleName == targetStyleName)
                        {
                            shou = para.Range.ListFormat.ListValue;
                            foundChapterTitle = true;
                        }
                        else if (foundChapterTitle && styleName.Contains(shouTobiraStyleName))
                        {
                            // 章番号が見つかった後のみ処理
                            Word.Range paraRange = para.Range;
                            string originalText = paraRange.Text;
                            string newText = numberRegex.Replace(originalText, shou.ToString());

                            if (originalText != newText)
                            {
                                paraRange.Text = newText;
                                paraRange.set_Style(ref shouTobiraStyle);
                            }
                        }
                    }
                }
                catch (System.Exception ex)
                {
                    // エラー内容をログに記録（必要に応じて）
                    System.Diagnostics.Debug.WriteLine(
                        $"Section {i} でエラー: {ex.Message}");
                }

                fm.progressBar1.Increment(1);
            }
        }

        private void UpdateChapterFrontNumbersOld(Word.Document objDocLast, MainForm fm)
        {
            string[] chapFrontItems = { "MJS_章扉-タイトル" };
            foreach (string styleName in chapFrontItems)
            {
                object styleObject = styleName;
                object shouTobira = "MJS_章扉-目次1";
                int allChap = objDocLast.Sections.Count;
                for (int i = 1; i <= allChap; i++)
                {
                    try
                    {
                        Word.Range wr = objDocLast.Sections[i].Range;
                        wr.Find.ClearFormatting();
                        if (objDocLast.Styles.Cast<Word.Style>().Any(s => s.NameLocal == styleName))
                        {
                            wr.Find.set_Style(ref styleObject);
                        }
                        else
                        {
                            // スタイルがなければスキップやログ出力
                        }
                        wr.Find.Execute();
                        if (wr.Find.Found)
                        {
                            int shou = 0;
                            for (int p = 1; p <= objDocLast.Sections[i].Range.Paragraphs.Count - 1; p++)
                            {
                                if (objDocLast.Sections[i].Range.Paragraphs[p].get_Style().NameLocal.Trim() == "MJS_章扉-タイトル")
                                    shou = objDocLast.Sections[i].Range.Paragraphs[p].Range.ListFormat.ListValue;
                                else if (objDocLast.Sections[i].Range.Paragraphs[p].get_Style().NameLocal.Trim().Contains("MJS_章扉-目次1"))
                                {
                                    objDocLast.Sections[i].Range.Paragraphs[p].Range.Text = Regex.Replace(objDocLast.Sections[i].Range.Paragraphs[p].Range.Text, @"^\d+?", shou.ToString());
                                    objDocLast.Sections[i].Range.Paragraphs[p].Range.set_Style(ref shouTobira);
                                }
                            }
                        }
                    }
                    catch
                    { }
                    fm.progressBar1.Increment(1);
                }
            }
        }
    }
}
