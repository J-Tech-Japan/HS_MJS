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
                        //wr.Find.Wrap = Word.WdFindWrap.wdFindStop;
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
                    {}
                    fm.progressBar1.Increment(1);
                }
            }
        }
    }
}
