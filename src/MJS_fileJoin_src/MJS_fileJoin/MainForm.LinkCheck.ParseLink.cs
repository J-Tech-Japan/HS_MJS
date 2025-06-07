using System.IO;
using System.Text.RegularExpressions;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void ParseLink(string file, Match m, out string targetURL, out string anchor)
        {
            // リンク先を # で分割し、ページ部分とアンカー部分に分ける
            string[] parts = m.Groups[1].Value.Split('#');

            if (parts.Length >= 2 && parts[0].Contains(".html") == false)
            {
                // ページ部分に拡張子がなければ .html を付加して絶対パスを作成
                targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + parts[0] + ".html";
            }
            else
            {
                // それ以外はページ部分をそのまま使って絶対パスを作成
                targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value.Split('#')[0];
            }

            // アンカー部分を取得
            anchor = m.Groups[1].Value.Split('#')[1];

            // アンカー部分に .html が含まれていれば除去
            if (anchor.Contains(".html"))
            {
                anchor = anchor.Replace(".html", "");
            }
        }
    }
}
