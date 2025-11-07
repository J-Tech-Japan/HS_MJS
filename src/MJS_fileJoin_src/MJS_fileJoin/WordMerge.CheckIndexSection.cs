using System.Linq;
using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // 末尾付近のセクションに「索引見出し」スタイルが使われているかチェック
        private bool CheckIndexSection(Word.Document objDocLast, MainForm fm, int chapCntLast)
        {
            bool last = false;
            string[] indexItems = { "索引見出し" };
            foreach (string styleName in indexItems)
            {
                object styleObject = styleName;
                int allChap = objDocLast.Sections.Count;
                for (int i = allChap; i > chapCntLast; i--)
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
                        last = true;
                        break;
                    }
                }
                fm.progressBar1.Increment(1);
            }
            return last;
        }
    }
}
