using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // 目次と索引の更新処理
        private void UpdateTocAndIndex(Word.Document doc, MainForm fm)
        {
            // 目次の更新
            using (var progress = Utils.BeginProgress(fm, "目次更新中...", 1))
            {
                if (doc.TablesOfContents.Count >= 1)
                    doc.TablesOfContents[1].Update();
                progress.Complete();
            }

            // 索引の更新
            using (var progress = Utils.BeginProgress(fm, "索引更新中...", 1))
            {
                if (doc.Indexes.Count >= 1)
                    doc.Indexes[1].Update();
                progress.Complete();
            }
        }
    }
}
