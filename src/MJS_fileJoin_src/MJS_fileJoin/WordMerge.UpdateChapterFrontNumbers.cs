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

            int allChap = objDocLast.Sections.Count;
            
            using (var progress = Utils.BeginProgress(fm, "章扉の項番号修正中...", allChap))
            {
                object shouTobiraStyle = shouTobiraStyleName;

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

                    progress.SetValue(i);
                }
                
                progress.Complete();
            }
        }
    }
}
