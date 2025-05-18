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
        private void CopyDirectory(string sourceDir, string destDir)
        {
            Directory.CreateDirectory(destDir);

            foreach (string file in Directory.GetFiles(sourceDir))
            {
                string destFile = Path.Combine(destDir, Path.GetFileName(file));
                if (Path.GetFileName(file).ToLower().Contains("image"))
                {

                }
                else if (Path.GetFileName(file).ToLower().Contains(".html"))
                {

                }
                else
                {
                    File.Copy(file, destFile, true);
                }
            }

            foreach (string subDir in Directory.GetDirectories(sourceDir))
            {
                string destSubDir = Path.Combine(destDir, Path.GetFileName(subDir));
                CopyDirectory(subDir, destSubDir);
            }
        }

        public static void CopyDirectory(string stSourcePath, string stDestPath, bool bOverwrite)
        {
            // コピー先のディレクトリがなければ作成する
            if (!Directory.Exists(stDestPath))
            {
                Directory.CreateDirectory(stDestPath);
                File.SetAttributes(stDestPath, File.GetAttributes(stSourcePath));
                bOverwrite = true;
            }

            // コピー元のディレクトリにあるすべてのファイルをコピーする
            if (bOverwrite)
            {
                foreach (string stCopyFrom in Directory.GetFiles(stSourcePath))
                {
                    string stCopyTo = Path.Combine(stDestPath, Path.GetFileName(stCopyFrom));
                    File.Copy(stCopyFrom, stCopyTo, true);
                }

                // 上書き不可能な場合は存在しない時のみコピーする
            }
            else
            {
                foreach (string stCopyFrom in Directory.GetFiles(stSourcePath))
                {
                    string stCopyTo = Path.Combine(stDestPath, Path.GetFileName(stCopyFrom));

                    if (!File.Exists(stCopyTo))
                    {
                        File.Copy(stCopyFrom, stCopyTo, false);
                    }
                }
            }

            // コピー元のディレクトリをすべてコピーする (再帰)
            foreach (string stCopyFrom in Directory.GetDirectories(stSourcePath))
            {
                string stCopyTo = Path.Combine(stDestPath, Path.GetFileName(stCopyFrom));
                CopyDirectory(stCopyFrom, stCopyTo, bOverwrite);
            }
        }

        private bool addHtmlDir(string dirPath)
        {
            if (Directory.Exists(Path.Combine(Path.GetDirectoryName(dirPath), "headerFile")))
            {
                string[] listFile = Directory.GetFiles(Path.Combine(Path.GetDirectoryName(dirPath), "headerFile"), "???.txt");
                if (listFile.Length == 0)
                {
                    MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(dirPath), "headerFile") + "」に書誌情報ファイルが存在しません。");
                }
                else
                {
                    bookInfo[dirPath] = new System.Data.DataTable();
                    bookInfo[dirPath].Columns.Add("Column1", typeof(bool));
                    bookInfo[dirPath].Columns.Add("Column2", typeof(string));
                    bookInfo[dirPath].Columns.Add("Column3", typeof(string));
                    bookInfo[dirPath].Columns.Add("Column4", typeof(string));
                    bookInfo[dirPath].Columns.Add("Column5", typeof(string));

                    bool isHtmlDir = false;
                    using (StreamReader sr = new StreamReader(listFile[0]))
                    {
                        while (!sr.EndOfStream)
                        {
                            string[] lineStr = (sr.ReadLine()).Split('\t');
                            string htmlName = "";
                            if (lineStr[2].Contains("#"))
                                continue;
                            else
                                htmlName = Path.Combine(dirPath, lineStr[2] + ".html");
                            if (lineStr.Length > 3 && !String.IsNullOrEmpty(lineStr[3]))
                            {
                                bookInfo[dirPath].Rows.Add(true, lineStr[0], lineStr[1], lineStr[2], lineStr[3]);
                            }
                            else
                            {
                                bookInfo[dirPath].Rows.Add(true, lineStr[0], lineStr[1], lineStr[2]);

                            }


                            if (!isHtmlDir &&
                                File.Exists(htmlName))
                            {
                                isHtmlDir = true;
                            }
                            if (!File.Exists(htmlName)) MessageBox.Show(htmlName + "が存在しません。\r\nwebHelpフォルダ内にHTMLを配置後、結合を実行してください。");
                        }
                    }
                    if (!isHtmlDir)
                    {
                        MessageBox.Show("書誌情報ファイルに紐づくHTMLファイルが見つかりません。\nフォルダをご確認ください。");
                        bookInfo.Remove(dirPath);
                    }
                    else
                    {
                        lbHtmlList.Items.Add(dirPath);
                        lbHtmlList.SelectedIndex = lbHtmlList.Items.Count - 1;
                    }
                }
            }
            else
            {
                MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(dirPath), "headerFile") + "」フォルダが存在しません。");
            }
            return false;
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

        private void createToc(XmlNode objToc)
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
                    createToc(toc);
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
