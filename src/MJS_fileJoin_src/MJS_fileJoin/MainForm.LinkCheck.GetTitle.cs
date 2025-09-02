using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // リンク元のページ名から参照すべきマージ後のページ名を特定する（リファクタリング版）
        private string GetLinkPage(MatchCollection mcMergePage, string[] parts)
        {
            // parts[0]のクリーンな値を取得
            string target = parts[0].Trim().Replace("'", "");

            if (mcMergePage == null || mcMergePage.Count == 0)
            {
                return target;
            }

            var mergePages = mcMergePage[0].Groups[1].Value.Trim().Split(',');
            foreach (var key in mergePages)
            {
                string trimmedKey = key.Trim();
                int colonIndex = trimmedKey.IndexOf(':');
                if (colonIndex < 0) continue;

                // コロンで分割し、各要素を比較
                var linkParts = trimmedKey.Split(':');
                foreach (var linkPart in linkParts)
                {
                    if (linkPart.Trim().Replace("'", "") == target)
                    {
                        // コロンの前の部分を返す
                        return trimmedKey.Substring(0, colonIndex).Trim();
                    }
                }
            }
            // 一致しない場合はデフォルトでtargetを返す
            return target;
        }

        private string GetTitleFromFile(string targetURL, string anchor, string file, Match m)
        {
            string titleName = "";
            using (StreamReader sr = new StreamReader(targetURL, Encoding.UTF8))
            {
                titleName = sr.ReadToEnd();
                if (String.IsNullOrEmpty(anchor))
                {
                    var head = new Regex(@"(?<=<title>)(.*?)(?=</title>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    var h = head.Match(titleName);
                    if (h.Success)
                    {
                        titleName = h.Groups[1].Value;
                    }
                    else titleName = "";
                }
                else
                {
                    if (titleName.Contains($@"<p class=""Heading3a"" id=""{anchor}"">"))
                    {
                        titleName = Regex.Match(titleName, $@"<p class=""Heading3a"" id=""{anchor}"">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                    }
                    else if (titleName.Contains($@"<p class=""Heading3"" id=""{anchor}"">"))
                    {
                        titleName = Regex.Match(titleName, $@"<p class=""Heading3"" id=""{anchor}"">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                        if (Regex.IsMatch(titleName, @"<span\s+name=""([^""]*)""\s+class=""ref""\s*/>") && titleName.Contains(m.Groups[2].Value))
                        {
                            titleName = m.Groups[2].Value;
                        }
                    }
                    else if (titleName.Contains($@"<p class=""Heading3"" id=""{m.Groups[1].Value.Replace(".html#", "＃")}"">"))
                    {
                        titleName = Regex.Match(titleName, $@"<p class=""Heading3"" id=""{m.Groups[1].Value.Replace(".html#", "＃")}"">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                    }
                    // ' Ver - 2023.16.08 - VyNL - ↓ - 追加'
                    else if (File.Exists(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{anchor}.html"))
                    {
                        using (StreamReader srAnchor = new StreamReader(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{anchor}.html", Encoding.UTF8))
                        {
                            titleName = srAnchor.ReadToEnd();
                            var head = new Regex(@"(?<=<title>)(.*?)(?=</title>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                            var h = head.Match(titleName);
                            if (h.Success)
                            {
                                titleName = h.Groups[1].Value;
                            }
                            else titleName = "";
                        }
                    }
                    // ' Ver - 2023.16.08 - VyNL - ↑ - 追加'
                    else titleName = "";
                }
            }
            return titleName;
        }

        private string GetHeading3aTitle(string file, Match m)
        {
            string titleName = "";
            string targetPath = Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", "");
            if (File.Exists(targetPath))
            {
                using (StreamReader sr = new StreamReader(targetPath, Encoding.UTF8))
                {
                    titleName = sr.ReadToEnd();
                    var head = new Regex(@"(?<=class=""Heading3a"" id=""" + Regex.Replace(m.Groups[1].Value, @".*?#", "") + @""">)(.*?)(?=</p>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                    var h = head.Match(titleName);
                    if (h.Success)
                    {
                        titleName = h.Groups[1].Value;
                    }
                    else
                    {
                        titleName = "";
                    }
                }
            }
            return titleName;
        }

        private string GetRefLinkTitle(string file, Match m)
        {
            string refLink = m.Groups[1].Value;
            
            // _Refリンクの場合、_Refの後ろの数字を抽出してrefPageから情報を取得
            var refMatch = Regex.Match(refLink, @"_Ref(\d+)");
            if (refMatch.Success)
            {
                string refNumber = refMatch.Groups[1].Value;
                string refId = "_Ref" + refNumber;
                
                // まずrefPageからタイトルを取得を試行
                string titleFromRefPage = GetTitleFromRefPage(file, refNumber);
                if (!string.IsNullOrEmpty(titleFromRefPage))
                {
                    return titleFromRefPage;
                }
                
                // refPageで見つからない場合、同じディレクトリ内のHTMLファイルを検索
                string titleFromHtmlFiles = GetTitleFromHtmlFiles(file, refId);
                if (!string.IsNullOrEmpty(titleFromHtmlFiles))
                {
                    return titleFromHtmlFiles;
                }
            }
            
            // 従来の方法でリンクタイトルを取得（フォールバック）
            string replaceTitleName = "";
            using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
            {
                string titleNameFile = srFile.ReadToEnd();
                replaceTitleName = Regex.Match(titleNameFile, @"<a\s+href=""" + Regex.Escape(m.Groups[1].Value) + @""">([^<]+)<\/a>").Groups[1].Value.Trim();
            }
            return replaceTitleName;
        }

        // refPageから指定されたrefNumberのタイトルを取得
        private string GetTitleFromRefPage(string file, string refNumber)
        {
            string allText = ReadAllText(file);
            var mcRefPage = GetRefPageMatches(allText);
            
            if (mcRefPage != null && mcRefPage.Count > 0)
            {
                string refPageContent = mcRefPage[0].Groups[1].Value;
                
                // refPage内から対応する参照IDの情報を取得
                string refPattern = @"_Ref" + refNumber + @"\s*:\s*\[(.*?)\]";
                var refPageMatch = Regex.Match(refPageContent, refPattern, RegexOptions.Singleline);
                
                if (refPageMatch.Success)
                {
                    string content = refPageMatch.Groups[1].Value.Trim();
                    string[] parts = content.Split(',');
                    
                    if (parts.Length >= 2)
                    {
                        // 2番目の要素（タイトル部分）を返す
                        return parts[1].Trim().Trim('\'', '"');
                    }
                }
            }
            
            return string.Empty;
        }

        // HTMLファイルから指定されたrefIdのタイトルを取得
        private string GetTitleFromHtmlFiles(string file, string refId)
        {
            string directory = Path.GetDirectoryName(file);
            string[] htmlFiles = Directory.GetFiles(directory, "*.html", SearchOption.TopDirectoryOnly);
            
            foreach (string htmlFile in htmlFiles)
            {
                string content = ReadAllText(htmlFile);
                
                // <span name="_Ref{数字}" class="ref" />の周辺のタイトルを取得
                string spanPattern = @"<span\s+name=""" + Regex.Escape(refId) + @"""\s+class=""ref""\s*/>\s*([^<]*?)(?=</p>|<|$)";
                var spanMatch = Regex.Match(content, spanPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                
                if (spanMatch.Success && !string.IsNullOrWhiteSpace(spanMatch.Groups[1].Value))
                {
                    return spanMatch.Groups[1].Value.Trim();
                }
                
                // MJS_refクラス内の参照タイトルを検索
                string mjsRefPattern = @"<p\s+class=""MJS_ref""><span\s+name=""" + Regex.Escape(refId) + @"""\s+class=""ref""\s*/>(.*?)</p>";
                var mjsRefMatch = Regex.Match(content, mjsRefPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                
                if (mjsRefMatch.Success)
                {
                    return mjsRefMatch.Groups[1].Value.Trim();
                }
                
                // より広範囲での検索：refIdを含む段落のテキストを取得
                string paragraphPattern = @"<p[^>]*>.*?<span[^>]*name=""" + Regex.Escape(refId) + @"""[^>]*>.*?</span>(.*?)</p>";
                var paragraphMatch = Regex.Match(content, paragraphPattern, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                
                if (paragraphMatch.Success && !string.IsNullOrWhiteSpace(paragraphMatch.Groups[1].Value))
                {
                    // HTMLタグを除去してテキストのみを返す
                    string text = Regex.Replace(paragraphMatch.Groups[1].Value, @"<[^>]*>", "").Trim();
                    if (!string.IsNullOrWhiteSpace(text))
                    {
                        return text;
                    }
                }
            }
            
            return string.Empty;
        }

        private string GetMjsRefTitleFromFile(string file, string refName)
        {
            using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
            {
                string titleNameFile = srFile.ReadToEnd();
                return Regex.Match(
                    titleNameFile,
                    @"<p class=""MJS_ref""><span name=""" + refName.Trim() + @""" class=""ref"" />(.*?)</p>",
                    RegexOptions.Singleline | RegexOptions.IgnoreCase
                ).Groups[1].Value.Trim();
            }
        }
    }
}