using System.Drawing;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // 指定されたリンクの検証結果をListViewに追加
        private void AddLinkTitleMatchResult(string file, Match m, string titleName)
        {
            bool isMatch = titleName == m.Groups[2].Value;
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            lvi.SubItems.Add(m.Groups[2].Value);
            lvi.SubItems.Add(isMatch ? "true" : "false");
            lvi.SubItems.Add(titleName);
            lvi.SubItems.Add("true");
            if (!isMatch)
            {
                lvi.BackColor = Color.Orange;
            }
        }

        // リンク切れやID不一致などがあった場合の結果をListViewに追加
        private void AddRefLinkBrokenOrIdMismatchResult(string file, string linkPage, Match m4, string content, int indexOfComma)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
            lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
            lvi.SubItems.Add("false");
            lvi.SubItems.Add("none");
            lvi.SubItems.Add("false");
            lvi.BackColor = Color.Red;
        }

        // 参照リンクの検証結果（正常）をListViewに追加
        private void AddRefLinkValidOrMatchedResult(string file, string linkPage, Match m4, string content, int indexOfComma, string titleName)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
            lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
            lvi.SubItems.Add("true");
            lvi.SubItems.Add(titleName);
            lvi.SubItems.Add("true");
        }

        // HTTPリンクの検証でエラーがあった場合の結果をListViewに追加
        private void AddHttpLinkErrorResult(string file, Match m)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            try
            {
                lvi.SubItems.Add(GetStatusCode(m.Groups[1].Value).ToString());
            }
            catch
            {
                lvi.SubItems.Add("取得に失敗しました。");
            }
            lvi.SubItems.Add("");
            lvi.SubItems.Add("");
            lvi.SubItems.Add("");
            lvi.BackColor = Color.Orange;
        }

        // 無効なリンクの検証結果をListViewに追加
        private void AddInvalidLinkResult(string file, Match m)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            lvi.SubItems.Add(m.Groups[2].Value);
            lvi.SubItems.Add("false");
            lvi.SubItems.Add("none");
            lvi.SubItems.Add("false");
            lvi.BackColor = Color.Red;
        }

        // 指定したURLのHTTPステータスコードを取得
        public static HttpStatusCode GetStatusCode(string url)
        {
            HttpWebRequest req = (HttpWebRequest)WebRequest.Create(url);
            HttpWebResponse res = null;
            HttpStatusCode statusCode;

            try
            {
                res = (HttpWebResponse)req.GetResponse();
                statusCode = res.StatusCode;
            }
            catch (WebException ex)
            {
                res = (HttpWebResponse)ex.Response;
                if (res != null)
                    statusCode = res.StatusCode;
                else
                    throw; // サーバ接続不可などの場合は再スロー
            }
            finally
            {
                if (res != null)
                    res.Close();
            }
            return statusCode;
        }

    }
}
