using Word = Microsoft.Office.Interop.Word;
using MJS_fileJoin;

namespace DocMergerComponent
{
    public partial class DocMerger
    {
        // 目次と索引を更新する
        private void UpdateIndexAndToc(Word.Document objDocLast, MainForm fm)
        {
            // 索引セクションの重複削除
            UpdateProgress(fm, "索引更新中...");
            RemoveDuplicateIndexSections(objDocLast, "索引見出し");

            // 目次の更新
            UpdateProgress(fm, "目次更新中...");
            UpdateTableOfContents(objDocLast);
            fm.progressBar1.Value = 1;

            // 索引の更新
            UpdateProgress(fm, "索引更新中...");
            UpdateIndex(objDocLast);
            fm.progressBar1.Value = 1;
        }

        // ヘルパーメソッド：目次の更新
        private void UpdateTableOfContents(Word.Document doc)
        {
            if (doc.TablesOfContents.Count >= 1)
                doc.TablesOfContents[1].Update();
        }

        // ヘルパーメソッド：インデックスの更新
        private void UpdateIndex(Word.Document doc)
        {
            if (doc.Indexes.Count >= 1)
                doc.Indexes[1].Update();
        }

        // ヘルパーメソッド：ラベルとプログレスバーの更新
        private void UpdateProgress(MainForm fm, string labelText)
        {
            fm.label10.Text = labelText;
            fm.progressBar1.Value = 0;
            fm.progressBar1.Maximum = 1;
        }
    }
}
