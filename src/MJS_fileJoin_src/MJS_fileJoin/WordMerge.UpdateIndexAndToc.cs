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
            if (doc.TablesOfContents.Count >= 1)
                doc.TablesOfContents[1].Update();
            fm.progressBar1.Value = 1;

            // 索引の更新
            fm.label10.Text = "索引更新中...";
            fm.progressBar1.Value = 0;
            fm.progressBar1.Maximum = 1;
            if (doc.Indexes.Count >= 1)
                doc.Indexes[1].Update();
            fm.progressBar1.Value = 1;
        }
    }
}
