using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;
using System.Linq;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        private void RemoveLastSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCntLast, ref bool last)
        {
            fm.label10.Text = "章扉章節項番号修正中...";
            fm.progressBar1.Maximum = objDocLast.Sections.Count;
            fm.progressBar1.Value = 1;

            string[] lastItems = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次" };
            foreach (string styleName in lastItems)
            {
                object styleObject = styleName;
                int allChap = objDocLast.Sections.Count;
                for (int i = allChap; i > chapCntLast; i--)
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
                            if (last)
                                last = false;
                            else
                            {
                                objDocLast.Sections[i].Range.Delete();
                                i--;
                                chapCntLast--;
                            }
                        }
                    }
                    catch
                    {
                        break;
                    }
                }
            }
        }

        private void RemoveDuplicateSectionsByStyle(Word.Document objDocLast, MainForm fm, int chapCnt, ref int chapCntLast)
        {
            fm.label10.Text = "重複箇所削除中...";
            fm.progressBar1.Maximum = 11;
            fm.progressBar1.Value = 1;

            string[] lsStyleName = { "MJS_見出し 1（項番なし）", "MJS_見出し 2（項番なし）", "MJS_マニュアルタイトル", "MJS_目次", "奥付タイトル", "索引見出し" };
            foreach (string styleName in lsStyleName)
            {
                object styleObject = styleName;
                for (int i = chapCnt + 1; i <= chapCntLast; i++)
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
                        objDocLast.Sections[i].Range.Delete();
                        i--;
                        chapCntLast--;
                    }
                }
                fm.progressBar1.Increment(1);
            }
        }
    }
}
