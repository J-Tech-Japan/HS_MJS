using System;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        // リンク元のページ名から、実際に参照すべきマージ後のページ名を特定する
        private string GetLinkPage(MatchCollection mcMergePage, string[] parts)
        {
            string linkPage = "";
            if (mcMergePage != null && mcMergePage.Count > 0)
            {
                var mergePage = mcMergePage[0].Groups[1].Value.Trim().Split(',');
                foreach (var key in mergePage)
                {
                    string[] link = key.Trim().Split(':');
                    foreach (var itemLink in link)
                    {
                        if (itemLink.Trim().Replace("'", "") == parts[0].Trim().Replace("'", ""))
                        {
                            int colonIndex = key.IndexOf(':');
                            linkPage = key.Substring(0, colonIndex).Trim();
                        }
                    }
                }
                if (linkPage == "")
                {
                    linkPage = parts[0].Trim().Replace("'", "");
                }
            }
            else
            {
                linkPage = parts[0].Trim().Replace("'", "");
            }
            return linkPage;
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
            string replaceTitleName = "";
            using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
            {
                string titleNameFile = srFile.ReadToEnd();
                replaceTitleName = Regex.Match(titleNameFile, @"<a\s+href=""" + m.Groups[1].Value + @""">([^<]+)<\/a>").Groups[1].Value.Trim();
            }
            return replaceTitleName;
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