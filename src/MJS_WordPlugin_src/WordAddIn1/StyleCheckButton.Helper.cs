using System.Collections.Generic;
using System.Linq;
using Word = Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        // ヘルパーメソッド: 無効なコメントを削除
        private void DeleteInvalidComments(Word.Comments comments)
        {
            var invalidTexts = new List<string>
                {
                    "使用できない書式です。",
                    "使用できない文字列です。",
                    "描画キャンバス外に行内配置でない画像があります。",
                    "上の段落に【MJS_手順番号リセット用】スタイルを挿入してください。",
                    "描画キャンバスが行内配置ではありません。"
                };

            foreach (Word.Comment comment in comments)
            {
                if (invalidTexts.Any(text => comment.Range.Text.Contains(text)))
                {
                    comment.Delete();
                }
            }
        }

        // ヘルパーメソッド: 選択範囲を元に戻す
        private void RestoreSelection(int start, int end)
        {
            var selection = Globals.ThisAddIn.Application.Selection;
            selection.Start = start;
            selection.End = end;
        }
    }
}
