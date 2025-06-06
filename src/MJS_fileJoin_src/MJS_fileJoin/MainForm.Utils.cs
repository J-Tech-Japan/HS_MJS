using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;

namespace MJS_fileJoin
{
    // UIから独立して動作する処理（ファイル操作、データ変換、HTML生成など）
    public partial class MainForm
    {
        private void CopyDirectory(string sourceDir, string destinationDir)
        {
            Directory.CreateDirectory(destinationDir);

            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string destinationFile = Path.Combine(destinationDir, Path.GetFileName(file));
                if (Path.GetFileName(file).ToLower().Contains("image"))
                {

                }
                else if (Path.GetFileName(file).ToLower().Contains(".html"))
                {

                }
                else
                {
                    File.Copy(file, destinationFile, true);
                }
            }

            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(destinationDir, Path.GetFileName(subDir));
                CopyDirectory(subDir, destSubDir);
            }
        }


        public static void CopyDirectory(string sourceDir, string destinationDir, bool overwrite)
        {
            // コピー先のディレクトリがなければ作成する
            if (!Directory.Exists(destinationDir))
            {
                Directory.CreateDirectory(destinationDir);
                File.SetAttributes(destinationDir, File.GetAttributes(sourceDir));
                overwrite = true;
            }

            // コピー元のディレクトリにあるすべてのファイルをコピーする
            if (overwrite)
            {
                foreach (string copyFrom in Directory.GetFiles(sourceDir))
                {
                    string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                    File.Copy(copyFrom, copyTo, true);
                }
            }
            else
            {
                foreach (string copyFrom in Directory.GetFiles(sourceDir))
                {
                    string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                    if (!File.Exists(copyTo))
                    {
                        File.Copy(copyFrom, copyTo, false);
                    }
                }
            }

            // コピー元のディレクトリをすべてコピーする (再帰)
            foreach (string copyFrom in Directory.GetDirectories(sourceDir))
            {
                string copyTo = Path.Combine(destinationDir, Path.GetFileName(copyFrom));
                CopyDirectory(copyFrom, copyTo, overwrite);
            }
        }
        
        private bool addHtmlDir(string dirPath)
        {
            string headerDir = Path.Combine(Path.GetDirectoryName(dirPath), "headerFile");
            if (!Directory.Exists(headerDir))
            {
                MessageBox.Show($"「{headerDir}」フォルダが存在しません。");
                return false;
            }

            string[] bibFiles = Directory.GetFiles(headerDir, "???.txt");
            if (bibFiles.Length == 0)
            {
                MessageBox.Show($"「{headerDir}」に書誌情報ファイルが存在しません。");
                return false;
            }

            var table = new System.Data.DataTable();
            table.Columns.Add("Column1", typeof(bool));
            table.Columns.Add("Column2", typeof(string));
            table.Columns.Add("Column3", typeof(string));
            table.Columns.Add("Column4", typeof(string));
            table.Columns.Add("Column5", typeof(string));

            bool hasHtml = false;
            try
            {
                using (var sr = new StreamReader(bibFiles[0]))
                {
                    while (!sr.EndOfStream)
                    {
                        string[] fields = sr.ReadLine().Split('\t');
                        if (fields.Length < 3) continue;
                        if (fields[2].Contains("#")) continue;

                        string htmlPath = Path.Combine(dirPath, fields[2] + ".html");
                        if (fields.Length > 3 && !string.IsNullOrEmpty(fields[3]))
                            table.Rows.Add(true, fields[0], fields[1], fields[2], fields[3]);
                        else
                            table.Rows.Add(true, fields[0], fields[1], fields[2]);

                        if (!hasHtml && File.Exists(htmlPath))
                            hasHtml = true;
                        if (!File.Exists(htmlPath))
                            MessageBox.Show($"{htmlPath}が存在しません。\r\nwebHelpフォルダ内にHTMLを配置後、結合を実行してください。");
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("書誌情報ファイルの読み込み中にエラーが発生しました: " + ex.Message);
                return false;
            }

            if (!hasHtml)
            {
                MessageBox.Show("書誌情報ファイルに紐づくHTMLファイルが見つかりません。\nフォルダをご確認ください。");
                return false;
            }

            bookInfo[dirPath] = table;
            lbHtmlList.Items.Add(dirPath);
            lbHtmlList.SelectedIndex = lbHtmlList.Items.Count - 1;
            return true;
        }

        public static void ListViewToCSV(ListView listView, string filePath, bool includeHidden)
        {
            //make header string
            StringBuilder result = new StringBuilder();
            WriteCSVRow(result, listView.Columns.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listView.Columns[i].Text);

            //export data rows
            foreach (ListViewItem listItem in listView.Items)
                WriteCSVRow(result, listView.Columns.Count, i => includeHidden || listView.Columns[i].Width > 0, i => listItem.SubItems[i].Text);

            File.WriteAllText(filePath, result.ToString(), Encoding.GetEncoding("Shift-JIS"));
        }

        private static void WriteCSVRow(StringBuilder result, int itemsCount, Func<int, bool> isColumnNeeded, Func<int, string> columnValue)
        {
            bool isFirstTime = true;
            for (int i = 0; i < itemsCount; i++)
            {
                if (!isColumnNeeded(i))
                    continue;

                if (!isFirstTime)
                    result.Append(",");
                isFirstTime = false;

                result.Append(String.Format("\"{0}\"", columnValue(i)));
            }
            result.AppendLine();
        }

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

        private void mergeHeaderFile()
        {
            string mergeText = "";
            string headerFilePath = "";
            foreach (string file in lbHtmlList.Items)
            {
                string[] files = Directory.GetFiles(file, "*.html");
                string pathName = "";

                foreach (string f in files)
                {
                    if (Regex.IsMatch(Path.GetFileName(f), @"^[A-Z]{3}\d+\.html$"))
                    {
                        pathName = Regex.Replace(Path.GetFileName(f), @"\d+\.html$", "");
                        break;
                    }
                }
                using (StreamReader sr = new StreamReader(Path.Combine(Path.Combine(Path.GetDirectoryName(file), "headerFile"), pathName + ".txt")))
                {
                    mergeText += sr.ReadToEnd();
                }
                if (!headerFilePath.Contains(pathName + "_")) headerFilePath += pathName + "_";
            }
            List<string> ls = new List<string>();

            if (!Directory.Exists(Path.Combine(tbOutputDir.Text, "headerFile"))) Directory.CreateDirectory(Path.Combine(tbOutputDir.Text, "headerFile"));
            using (StreamWriter sw = new StreamWriter(Path.Combine(tbOutputDir.Text, "headerFile\\" + Regex.Replace(headerFilePath, @"_$", "")) + ".txt"))
            using (StringReader sr = new StringReader(mergeText))
            {
                while (sr.Peek() > 0)
                {
                    string lineText = sr.ReadLine();
                    if (!ls.Contains(Regex.Replace(lineText, @"^.*?\t.*?\t(.*?)$", "$1")))
                    {
                        ls.Add(Regex.Replace(lineText, @"^.*?\t.*?\t(.*?)$", "$1"));
                        sw.WriteLine(lineText);
                    }
                }
            }
        }

        private void CreateToc(XmlNode objToc)
        {
            string htmlToc = "";
            foreach (XmlNode toc in objToc.SelectNodes("item"))
            {
                if (htmlToc != "")
                {
                    htmlToc = htmlToc + ",";
                }

                htmlToc = htmlToc + @"{""type"":""";

                if (toc.SelectNodes("item").Count != 0)
                {
                    htmlToc = htmlToc + "book";
                }
                else
                {
                    htmlToc = htmlToc + "item";
                }

                htmlToc += @""",""name"":""" + ((XmlElement)toc).GetAttribute("title") + @"""";

                if (toc.SelectNodes("item").Count != 0)
                {
                    htmlToc += @",""key"":""toc" + (toc.SelectNodes("preceding::item[boolean(item)]").Count + toc.SelectNodes("ancestor-or-self::item").Count) + @"""";
                }

                if (((XmlElement)toc).GetAttribute("href") != "")
                {
                    htmlToc += @",""url"":""" + ((XmlElement)toc).GetAttribute("href") + @".html""";
                }

                htmlToc += "}";

                if (toc.SelectNodes("item").Count != 0)
                {
                    CreateToc(toc);
                }
            }

            if (htmlToc != "")
            {

                if (Regex.IsMatch(htmlToc, @"""url""\s*:\s*""([^""]*)#([^""]*)"""))
                {
                    htmlToc = Regex.Replace(htmlToc, @"""url""\s*:\s*""([^""]*)#([^""]*)""", match =>
                    {
                        string url = match.Groups[1].Value;
                        string hash = match.Groups[2].Value;

                        return $@"""url"": ""{url}.html#{hash}""";
                    });
                }

                int itemCount = objToc.SelectNodes("preceding::item[boolean(item)]").Count + objToc.SelectNodes("ancestor-or-self::item").Count;
                StreamWriter sw = new StreamWriter(tbOutputDir.Text + "\\" + exportDir + "\\whxdata\\toc" + ((itemCount != 0) ? itemCount.ToString() : "") + ".new.js", false, Encoding.UTF8);
                sw.WriteLine("(function() {");
                sw.WriteLine("var toc =  [" + htmlToc + "];");
                sw.WriteLine("window.rh.model.publish(rh.consts('KEY_TEMP_DATA'), toc, { sync:true });");
                sw.WriteLine("})();");
                sw.Close();
            }
        }
    }
}
