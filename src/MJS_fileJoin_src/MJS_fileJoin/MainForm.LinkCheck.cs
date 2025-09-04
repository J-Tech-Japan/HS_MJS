using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // HTMLファイルのリンクチェック
        private void LinkCheck(string folder)
        {
            string[] files = Directory.GetFiles(folder, "*.html", SearchOption.AllDirectories);
            ShowProgressStart(files.Length);

            foreach (string file in files)
            {
                UpdateProgress(file, folder);
                string allText = ReadAllText(file);

                var mc2 = GetAnchorMatches(allText);
                var mc3 = GetMjsRefMatches(allText);
                var mcRefPage = GetRefPageMatches(allText);
                var mcMergePage = GetMergePageMatches(allText);

                if (mcRefPage != null && mcRefPage.Count > 0)
                {
                    HandleRefPageLinks(file, mc3, mcRefPage, mcMergePage);
                }

                foreach (Match m in mc2)
                {
                    HandleAnchorLink(file, m);
                }

                logen.Clear();
                foreach (ListViewItem lvi in listView1.Items)
                    logen.Add(lvi);
            }

            // 外部参照の行を薄いオレンジに塗り直す
            foreach (ListViewItem lvi in listView1.Items)
            {
                if (lvi.BackColor == LightRed)
                {
                    string file = lvi.Text;
                    string linkPage = lvi.SubItems.Count > 1 ? lvi.SubItems[1].Text : "";

                    if (!IsInternalReference(file, linkPage))
                    {
                        lvi.BackColor = LightOrange;
                    }
                }
            }

            // 列幅を文字数に合わせる
            //UpdateListViewColumnsWidth();
        }

        // 進捗バーの初期化
        private void ShowProgressStart(int fileCount)
        {
            progressBar2.Visible = true;
            progressBar2.Value = 0;
            progressBar2.Maximum = fileCount - 1;
            label14.Visible = true;
        }

        // 進捗バーとラベルを更新
        private void UpdateProgress(string file, string folder)
        {
            label14.Text = Path.GetFileName(file) + " / " + Path.GetFileName(folder);
            label14.Refresh();
            progressBar2.Increment(1);
        }

        // 指定ファイルの全テキストをUTF-8で読み込む
        private string ReadAllText(string file)
        {
            using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
            {
                return sr.ReadToEnd();
            }
        }

        // HTML内のアンカーリンク（<a href="">）を抽出
        private MatchCollection GetAnchorMatches(string allText)
        {
            Regex r2 = new Regex(@"(?<=<a href="")([^""]*?)"">([^<]*?)(?=</a>)");
            return r2.Matches(allText);
        }

        // MJS_refクラスの参照を抽出
        private MatchCollection GetMjsRefMatches(string allText)
        {
            Regex r3 = new Regex(@"(?<=<p\sclass=""MJS_ref""><span\sname=""([^""]+)""\sclass=""ref""\s*/>([^<]*?)</p>)");
            return r3.Matches(allText);
        }

        // refPage定義部分を抽出
        private MatchCollection GetRefPageMatches(string allText)
        {
            Regex r4 = new Regex(@"refPage\s*=\s*{([\s\S]*?)}");
            return r4.Matches(allText);
        }

        // mergePage定義部分を抽出
        private MatchCollection GetMergePageMatches(string allText)
        {
            Regex r5 = new Regex(@"mergePage\s*=\s*{([\s\S]*?)}");
            return r5.Matches(allText);
        }

        // refPageリンクの検証
        private void HandleRefPageLinks(string file, MatchCollection mc3, MatchCollection mcRefPage, MatchCollection mcMergePage)
        {
            var refPage = mcRefPage[0].Groups[1].Value;
            foreach (Match m4 in mc3)
            {
                if (Regex.IsMatch(mcRefPage[0].Groups[1].Value, m4.Groups[1].Value.Trim().Replace("_ref", "")))
                {
                    MatchCollection matches = Regex.Matches(refPage, $"{m4.Groups[1].Value.Trim().Replace("_ref", "")}:(.*?])");
                    foreach (Match match in matches)
                    {
                        string content = match.Groups[1].Value.Trim().Replace("[", "").Replace("]", "");
                        string[] parts = content.Split(',');
                        int indexOfComma = content.IndexOf(',');
                        string linkPage = (mcMergePage != null && mcMergePage.Count > 0) ? GetLinkPage(mcMergePage, parts) : "";
                        string targetHtml = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html";
                        if (File.Exists(targetHtml))
                        {
                            string titleName = ReadAllText(targetHtml);
                            if (Regex.IsMatch(titleName, @"<span name=""" + m4.Groups[1].Value.Replace("_ref", "") + @""" class=""ref"" />"))
                            {
                                titleName = GetMjsRefTitleFromFile(file, m4.Groups[1].Value);
                            }
                            else titleName = "";
                            
                            // タイトル比較を正規化して行う
                            string expectedTitle = content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", "");
                            string normalizedTitleName = NormalizeTitle(titleName);
                            string normalizedExpectedTitle = NormalizeTitle(expectedTitle);
                            
                            if (normalizedTitleName.Contains(normalizedExpectedTitle) || normalizedExpectedTitle.Contains(normalizedTitleName))
                            {
                                AddRefLinkValidOrMatchedResult(file, linkPage, m4, content, indexOfComma, titleName);
                            }
                            else
                            {
                                AddRefLinkBrokenOrIdMismatchResult(file, linkPage, m4, content, indexOfComma);
                            }
                        }
                        else
                        {
                            AddRefLinkBrokenOrIdMismatchResult(file, linkPage, m4, content, indexOfComma);
                        }
                    }
                    logen.Clear();
                    foreach (ListViewItem lvi in listView1.Items)
                        logen.Add(lvi);
                }
            }
        }

        private void HandleAnchorLink(string file, Match m)
        {
            string targetURL = "";
            string anchor = "";
            if (Regex.IsMatch(m.Groups[1].Value, @"^#"))
            {
                targetURL = file;
                anchor = m.Groups[1].Value.Replace("#", "");
            }
            else if (m.Groups[1].Value.Contains("#"))
            {
                ParseLink(file, m, out targetURL, out anchor);
            }
            else
            {
                targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value;
            }
            string titleName = "";
            if (Regex.IsMatch(m.Groups[1].Value, "^http", RegexOptions.Singleline | RegexOptions.IgnoreCase))
            {
                AddHttpLinkErrorResult(file, m);
                return;
            }
            if (File.Exists(targetURL))
            {
                titleName = GetTitleFromFile(targetURL, anchor, file, m);
                titleName = Regex.Replace(titleName, @"<span[^>]*?>", "");
                titleName = Regex.Replace(titleName, @"</span>", "");
                AddLinkTitleMatchResult(file, m, titleName);
            }
            else if (File.Exists(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", "")))
            {
                titleName = GetHeading3aTitle(file, m);
                AddLinkTitleMatchResult(file, m, titleName);
            }
            else if (m.Groups[1].Value.StartsWith("_Ref"))
            {
                // _Refリンクの参照先存在判定
                if (IsRefLinkValid(file, m.Groups[1].Value))
                {
                    titleName = GetRefLinkTitle(file, m);
                    AddLinkTitleMatchResult(file, m, titleName);
                }
                else
                {
                    // 参照先が存在しない場合は赤くする
                    AddInvalidLinkResult(file, m);
                }
            }
            else
            {
                AddInvalidLinkResult(file, m);
            }
        }

        // _Refリンクの参照先が存在するかを判定
        private bool IsRefLinkValid(string file, string refLink)
        {
            // _Refの後の数字を抽出
            var refMatch = Regex.Match(refLink, @"_Ref(\d+)");
            if (!refMatch.Success)
            {
                return false; // _Refの後に数字がない場合は無効
            }

            string refNumber = refMatch.Groups[1].Value;
            string allText = ReadAllText(file);

            // 同じファイル内でrefPage定義をチェック
            var mcRefPage = GetRefPageMatches(allText);
            if (mcRefPage != null && mcRefPage.Count > 0)
            {
                string refPageContent = mcRefPage[0].Groups[1].Value;
                // refPage内に対応する参照IDが存在するかチェック
                string refPattern = @"_Ref" + refNumber + @"\s*:\s*\[";
                if (Regex.IsMatch(refPageContent, refPattern))
                {
                    return true;
                }
            }

            // 同じディレクトリ内の他のHTMLファイルでref要素をチェック
            string directory = Path.GetDirectoryName(file);
            string[] htmlFiles = Directory.GetFiles(directory, "*.html", SearchOption.TopDirectoryOnly);

            foreach (string htmlFile in htmlFiles)
            {
                string content = ReadAllText(htmlFile);
                // <span name="_Ref{数字}" class="ref" />の存在をチェック
                string spanPattern = @"<span\s+name=""_Ref" + refNumber + @"""\s+class=""ref""\s/>";
                if (Regex.IsMatch(content, spanPattern, RegexOptions.IgnoreCase))
                {
                    return true;
                }
            }
            return false; // 参照先が見つからない
        }
    }
}
