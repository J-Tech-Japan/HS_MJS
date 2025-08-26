// GenerateHTMLButton.Search.cs

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private static string BuildSearchJs()
        {
            // searchBase.jsを基にsearch.jsを生成する
            // searchBase.jsファイルの先頭に "var searchWords = $('♪');" を追加するだけの簡潔な実装
            string searchJs = @"var searchWords = $('♪');" + "\n";
            
            // htmlTemplates.zipから展開されたsearchBase.jsの内容をそのまま追加
            // 実際のsearchBase.jsの内容の読み込みは、GenerateSearchFiles内で行われる
            return searchJs;
        }

    }
}
