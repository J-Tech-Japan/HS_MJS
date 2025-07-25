using System;
using System.Drawing;
using System.IO;
using System.Net;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // 薄いオレンジ色と赤色を定義
        private static readonly Color LightOrange = Color.FromArgb(255, 255, 216, 93);
        private static readonly Color LightRed = Color.FromArgb(255, 255, 180, 180);

        // タイトル一致チェック
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
                lvi.BackColor = LightRed;
            }
        }

        // 内部参照判定メソッド
        private bool IsInternalReference(string sourceFile, string linkPage)
        {
            // 参照先が空なら内部参照とする
            if (string.IsNullOrEmpty(linkPage))
                return true;

            // 先頭が#なら内部参照
            if (linkPage.StartsWith("#"))
                return true;

            // 相対参照なら内部参照
            if (linkPage.StartsWith("./") || linkPage.StartsWith(@".\"))
                return true;

            // sourceFileがパスの場合はファイル名だけを使う
            string sourceFileName = Path.GetFileName(sourceFile);

            // linkPageからフラグメント（#以降）を除去し、ファイル名だけを抽出
            string linkPageWithoutFragment = linkPage.Split('#')[0];
            string linkFileName = Path.GetFileName(linkPageWithoutFragment);

            // ファイル名が_Refで始まる場合は内部参照
            if (linkFileName.StartsWith("_Ref", StringComparison.OrdinalIgnoreCase))
                return true;

            // linkPageが同じディレクトリになる場合は内部参照
            try
            {
                // sourceFileのディレクトリ
                string sourceDir = Path.GetDirectoryName(Path.GetFullPath(sourceFile));

                // linkPageの絶対パス
                string linkFullPath = Path.GetFullPath(Path.Combine(sourceDir, linkPageWithoutFragment));
                string linkDir = Path.GetDirectoryName(linkFullPath);

                if (string.Equals(sourceDir, linkDir, StringComparison.OrdinalIgnoreCase))
                    return true;
            }
            catch
            {
                // パス解決失敗時は無視
            }

            // ファイル名が「アルファベット3文字+5桁数字」形式か判定
            //var fileNamePattern = new Regex(@"^([A-Z]{3})(\d{5})\.html$", RegexOptions.IgnoreCase);

            //var sourceMatch = fileNamePattern.Match(sourceFileName);
            //var linkMatch = fileNamePattern.Match(linkFileName);

            //if (sourceMatch.Success && linkMatch.Success)
            //{
            //    // 先頭2桁の数字が一致すれば内部参照
            //    string sourcePrefix = sourceMatch.Groups[2].Value.Substring(0, 2);
            //    string linkPrefix = linkMatch.Groups[2].Value.Substring(0, 2);
            //    return sourcePrefix == linkPrefix;
            //}

            return false;
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
            lvi.BackColor = LightRed;
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
            lvi.BackColor = LightOrange;
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
            lvi.BackColor = LightRed;
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
