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
                                string linkPage = "";
                                string[] parts = content.Split(',');
                                int indexOfComma = content.IndexOf(',');

                                if (mcMergePage != null && mcMergePage.Count > 0)
                                {
                                    linkPage = GetLinkPage(mcMergePage, parts);
                                }

                                if (File.Exists(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html"))
                                {
                                    using (StreamReader sr = new StreamReader(Path.GetFullPath(Path.GetDirectoryName(file)) + "/" + $"{parts[0].Trim().Replace("'", "")}.html", Encoding.UTF8))
                                    {
                                        titleName = sr.ReadToEnd();
                                    }

                                    if (Regex.IsMatch(titleName, @"<span name=""" + m4.Groups[1].Value.Replace("_ref", "") + @""" class=""ref"" />"))
                                    {
                                        titleName = GetMjsRefTitleFromFile(file, m4.Groups[1].Value);
                                    }

                                    else titleName = "";

                                    if (titleName.Contains(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", "")))
                                    {
                                        AddRefLinkCheckResult(file, linkPage, m4, content, indexOfComma, titleName);
                                    }
                                }

                                else
                                {
                                    AddRefLinkCheckErrorResult(file, linkPage, m4, content, indexOfComma);
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
                        titleName = GetTitleFromFile(targetURL, anchor, file, m);

                        // using regex clear tag <span> </span> in titleName
                        titleName = Regex.Replace(titleName, @"<span[^>]*?>", "");
                        titleName = Regex.Replace(titleName, @"</span>", "");

                        // check title same with title in link
                        AddLinkCheckResult(file, m, titleName);
                    }

                    else if (File.Exists(Path.GetDirectoryName(Path.GetFullPath(file)) + "\\" + Regex.Replace(m.Groups[1].Value, @"#.*?$", "")))
                    {
                        titleName = GetHeading3aTitle(file, m);
                        AddLinkCheckResult(file, m, titleName);
                    }

                    else if (m.Groups[1].Value.StartsWith("_Ref"))
                    {
                        titleName = GetRefLinkTitle(file, m);
                        AddLinkCheckResult(file, m, titleName);
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
