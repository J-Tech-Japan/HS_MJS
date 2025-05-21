using System;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void LinkCheck(string folder)
        {
            string[] files = Directory.GetFiles(folder, "*.html", SearchOption.AllDirectories);
            progressBar2.Visible = true;
            progressBar2.Value = 0;
            progressBar2.Maximum = files.Count() - 1;
            label14.Visible = true;

            foreach (string file in files)
            {
                label14.Text = Path.GetFileName(file) + " / " + Path.GetFileName(folder);
                label14.Refresh();
                progressBar2.Increment(1);
                string allText = "";

                using (StreamReader sr = new StreamReader(file, Encoding.UTF8))
                {
                    allText = sr.ReadToEnd();
                }

                Regex r = new Regex(@"(?<=<a href="")http[^""]*?(?="")");
                MatchCollection mc = r.Matches(allText);
                
                Regex r2 = new Regex(@"(?<=<a href="")([^""]*?)"">([^<]*?)(?=</a>)");
                MatchCollection mc2 = r2.Matches(allText);

                Regex r3 = new Regex(@"(?<=<p\sclass=""MJS_ref""><span\sname=""([^""]+)""\sclass=""ref""\s*/>([^<]*?)<\/p>)");
                MatchCollection mc3 = r3.Matches(allText);

                Regex r4 = new Regex(@"refPage\s*=\s*{([\s\S]*?)}");
                MatchCollection mcRefPage = r4.Matches(allText);

                Regex r5 = new Regex(@"mergePage\s*=\s*{([\s\S]*?)}");
                MatchCollection mcMergePage = r5.Matches(allText);

                if (mcRefPage != null && mcRefPage.Count > 0)
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

                                string titleName = "";
                                string replaceTitleName = "";
                                string linkPage = "";
                                string[] parts = content.Split(',');
                                int indexOfComma = content.IndexOf(',');

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

                                if (File.Exists(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html"))
                                {
                                    using (StreamReader sr = new StreamReader(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html", Encoding.UTF8))
                                    {
                                        titleName = sr.ReadToEnd();

                                    }
                                    if (Regex.IsMatch(titleName, @"<span name=""" + m4.Groups[1].Value.Replace("_ref", "") + @""" class=""ref"" />"))
                                    {
                                        using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
                                        {
                                            string titleNameFile = srFile.ReadToEnd();
                                            replaceTitleName = Regex.Match(titleNameFile, @"<p class=""MJS_ref""><span name=""" + m4.Groups[1].Value.Trim() + @""" class=""ref"" />(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                        }
                                        titleName = replaceTitleName;
                                    }
                                    else titleName = "";
                                    if (titleName.Contains(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", "")))
                                    {
                                        ListViewItem lvi = listView1.Items.Add(file);
                                        lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
                                        lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
                                        lvi.SubItems.Add("true");
                                        lvi.SubItems.Add(titleName);
                                        lvi.SubItems.Add("true");
                                    }

                                }
                                else
                                {
                                    ListViewItem lvi = listView1.Items.Add(file);
                                    lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
                                    lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
                                    lvi.SubItems.Add("false");
                                    lvi.SubItems.Add("none");
                                    lvi.SubItems.Add("false");
                                    lvi.BackColor = Color.Red;
                                }

                            }
                            logen.Clear();
                            foreach (ListViewItem lvi in listView1.Items)
                                logen.Add(lvi);
                        }

                    }

                }
                
                foreach (Match m in mc2)
                {
                    string targetURL = "";
                    string anchor = "";
                    string replaceTitleName = "";

                    if (Regex.IsMatch(m.Groups[1].Value, @"^#"))
                    {
                        targetURL = file;
                        anchor = m.Groups[1].Value.Replace("#", "");
                    }
                    else if (m.Groups[1].Value.Contains("#"))
                    {
                        // check link
                        string[] parts = m.Groups[1].Value.Split('#');

                        if (parts.Length >= 2 && parts[0].Contains(".html") == false)
                        {
                            // targetURL need concat with extension .html => pass check with extension
                            targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + parts[0] + ".html";
                        }
                        else
                        {
                            targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value.Split('#')[0];

                        }

                        anchor = m.Groups[1].Value.Split('#')[1];
                        // check anchor have extension .html
                        if (anchor.Contains(".html"))
                        {
                            anchor = anchor.Replace(".html", "");
                        }
                    }
                    else
                    {
                        targetURL = Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + m.Groups[1].Value;
                    }
                    string titleName = "";
                    if (Regex.IsMatch(m.Groups[1].Value, "^http", RegexOptions.Singleline | RegexOptions.IgnoreCase))
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
                        lvi.BackColor = Color.Red;
                        continue;
                    }
                    if (File.Exists(targetURL))
                    {
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

                                if (titleName.Contains(@"<p class=""Heading3a"" id=""" + anchor + @""">"))
                                {
                                    titleName = Regex.Match(titleName, @"<p class=""Heading3a"" id=""" + anchor + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                }
                                else if (titleName.Contains(@"<p class=""Heading3"" id=""" + anchor + @""">"))
                                {

                                    titleName = Regex.Match(titleName, @"<p class=""Heading3"" id=""" + anchor + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();
                                    if (Regex.IsMatch(titleName, @"<span\s+name=""([^""]*)""\s+class=""ref""\s*/>") && titleName.Contains(m.Groups[2].Value))
                                    {
                                        titleName = m.Groups[2].Value;
                                    }

                                }
                                else if (titleName.Contains(@"<p class=""Heading3"" id=""" + m.Groups[1].Value.Replace(".html#", "＃") + @""">"))
                                {

                                    titleName = Regex.Match(titleName, @"<p class=""Heading3"" id=""" + m.Groups[1].Value.Replace(".html#", "＃") + @""">(.*?)</p>", RegexOptions.Singleline | RegexOptions.IgnoreCase).Groups[1].Value.Trim();

                                }
                                //' Ver - 2023.16.08 - VyNL - ↓ - 追加'
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

                                //' Ver - 2023.16.08 - VyNL - ↑ - 追加'
                                else titleName = "";
                            }
                        }
                        // using regex clear tag <span> </span> in titleName
                        titleName = Regex.Replace(titleName, @"<span[^>]*?>", "");
                        titleName = Regex.Replace(titleName, @"</span>", "");
                        // check title same with title in link
                        if (titleName != m.Groups[2].Value)
                        {
                            // if not same add to listview
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255,192,203);//#ffc0cb pink
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else if (File.Exists(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", "")))
                    {
                        using (StreamReader sr = new StreamReader(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", ""), Encoding.UTF8))
                        {
                            titleName = sr.ReadToEnd();
                            var head = new Regex(@"(?<=class=""Heading3a"" id=""" + Regex.Replace(m.Groups[1].Value, @".*?#", "") + @""">)(.*?)(?=</p>)", RegexOptions.IgnoreCase | RegexOptions.Singleline);
                            var h = head.Match(titleName);
                            if (h.Success)
                            {
                                titleName = h.Groups[1].Value;
                            }
                            else titleName = "";
                        }
                        if (titleName != m.Groups[2].Value)
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255, 192, 203);//#ffc0cb pink
                            //allCheck += file + "," + m.Groups[1].Value + "," + m.Groups[2].Value + ",false," + titleName + ",true" + "\r\n";
                            //Console.WriteLine(file + "\r\n" + m.Groups[1].Value + "\r\n" + m.Groups[2].Value + "\r\nfalse\r\n" + titleName + "\r\ntrue" + "\r\n");
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else if (m.Groups[1].Value.StartsWith("_Ref"))
                    {
                        using (StreamReader srFile = new StreamReader(file, Encoding.UTF8))
                        {
                            string titleNameFile = srFile.ReadToEnd();

                            replaceTitleName = Regex.Match(titleNameFile, @"<a\s+href=""" + m.Groups[1].Value + @""">([^<]+)<\/a>").Groups[1].Value.Trim();
                        }
                        titleName = replaceTitleName;

                        if (titleName != m.Groups[2].Value)
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("false");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                            lvi.BackColor = Color.Red;
                            //lvi.BackColor = Color.FromArgb(255, 192, 203);//#ffc0cb pink
                            //allCheck += file + "," + m.Groups[1].Value + "," + m.Groups[2].Value + ",false," + titleName + ",true" + "\r\n";
                            //Console.WriteLine(file + "\r\n" + m.Groups[1].Value + "\r\n" + m.Groups[2].Value + "\r\nfalse\r\n" + titleName + "\r\ntrue" + "\r\n");
                        }
                        else
                        {
                            ListViewItem lvi = listView1.Items.Add(file);
                            lvi.SubItems.Add(m.Groups[1].Value);
                            lvi.SubItems.Add(m.Groups[2].Value);
                            lvi.SubItems.Add("true");
                            lvi.SubItems.Add(titleName);
                            lvi.SubItems.Add("true");
                        }
                    }
                    else
                    {
                        ListViewItem lvi = listView1.Items.Add(file);
                        lvi.SubItems.Add(m.Groups[1].Value);
                        lvi.SubItems.Add(m.Groups[2].Value);
                        lvi.SubItems.Add("false");
                        lvi.SubItems.Add("none");
                        lvi.SubItems.Add("false");
                        lvi.BackColor = Color.Red;
                    }
                }
                logen.Clear();
                foreach (ListViewItem lvi in listView1.Items)
                    logen.Add(lvi);
            }
            //using (StreamWriter sw = new StreamWriter("./kekka.csv", false, Encoding.UTF8))
            //{
            //    sw.Write(allCheck);
            //}
            //Assembly myAssembly = Assembly.GetEntryAssembly();
            //string path = Path.GetDirectoryName(myAssembly.Location) + "/";
            //Process.Start(path + "./kekka.csv");
        }
    }
}
