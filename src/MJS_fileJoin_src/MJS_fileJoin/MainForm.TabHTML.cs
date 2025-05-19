using System;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void btnSelectJoinList_Click(object sender, EventArgs e)
        {
            try
            {
                if (Directory.Exists(Path.GetDirectoryName(tbSelectJoinList.Text)))
                {
                    openFileDialog1.InitialDirectory = Path.GetDirectoryName(tbSelectJoinList.Text);
                }
            }
            catch (Exception ex)
            {
            }
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            tbSelectJoinList.Text = openFileDialog1.FileName;
        }

        private void tbSelectJoinList_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            for (int i = 0; i < s.Length; i++)
            {
                if (File.Exists(s[i]))
                {
                    tbSelectJoinList.Text = s[i];
                    break;
                }
            }
        }

        private void tbSelectJoinList_TextChanged(object sender, EventArgs e)
        {
            if (File.Exists(tbSelectJoinList.Text))
            {
                XmlDocument list = new XmlDocument();
                try
                {
                    list.Load(tbSelectJoinList.Text);
                    if (list.SelectSingleNode(".//changeTitle") != null)
                    {
                        chbChangeTitle.Checked = true;
                        tbChangeTitle.Enabled = true;
                        tbChangeTitle.Text = list.SelectSingleNode(".//changeTitle").InnerText;
                    }
                    if (list.SelectSingleNode(".//addTopLevel") != null)
                    {
                        chbAddTop.Checked = true;
                        tbAddTop.Enabled = true;
                        tbAddTop.Text = list.SelectSingleNode(".//addTopLevel").InnerText;
                    }

                    lbHtmlList.Items.Clear();
                    bookInfo.Clear();

                    foreach (XmlNode htmlList in list.SelectNodes(".//htmlList/item"))
                    {
                        string path = ((XmlElement)htmlList).GetAttribute("src");
                        int lbListItemIndex = lbHtmlList.Items.Add(path);

                        if (Directory.Exists(Path.Combine(Path.GetDirectoryName(path), "headerFile")))
                        {
                            string[] listFile = Directory.GetFiles(Path.Combine(Path.GetDirectoryName(path), "headerFile"), "???.txt");
                            if (listFile.Length == 0)
                            {
                                MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(path), "headerFile") + "」に書誌情報ファイルが存在しません。");
                            }
                            else
                            {
                                bookInfo[path] = new System.Data.DataTable();

                                bookInfo[path].Columns.Add("Column1", typeof(bool));
                                bookInfo[path].Columns.Add("Column2", typeof(string));
                                bookInfo[path].Columns.Add("Column3", typeof(string));
                                bookInfo[path].Columns.Add("Column4", typeof(string));

                                using (StreamReader sr = new StreamReader(listFile[0]))
                                {
                                    while (!sr.EndOfStream)
                                    {
                                        string[] lineStr = (sr.ReadLine()).Split('\t');
                                        if (lineStr[2].Contains("#")) continue;
                                        bookInfo[path].Rows.Add(((htmlList.SelectSingleNode(".//checked[@id='" + lineStr[2] + "']") != null) ? true : false), lineStr[0], lineStr[1], lineStr[2]);
                                    }
                                }
                            }
                        }
                        else
                        {
                            MessageBox.Show("「" + Path.Combine(Path.GetDirectoryName(path), "headerFile") + "」フォルダが存在しません。");
                            lbHtmlList.Items.RemoveAt(lbListItemIndex);
                        }
                    }

                    if (list.SelectSingleNode(".//outputDir") != null)
                    {
                        tbOutputDir.Text = list.SelectSingleNode(".//outputDir/@src").InnerText;
                    }
                }
                catch (XmlException xmlex)
                {
                    if (Regex.IsMatch(tbSelectJoinList.Text, @"\.xml$"))
                    {
                        MessageBox.Show("結合リストが破損しています。");
                    }
                    else
                    {
                        MessageBox.Show("XMLファイルを選択してください。");
                    }
                    tbSelectJoinList.Text = "";
                }
                catch (XPathException xpathex)
                {
                    MessageBox.Show("結合リストが破損しています。");
                    tbSelectJoinList.Text = "";
                }
            }
        }

        private void chbChangeTitle_CheckedChanged(object sender, EventArgs e)
        {
            if (chbChangeTitle.Checked)
            {
                tbChangeTitle.Enabled = true;
            }
            else
            {
                tbChangeTitle.Enabled = false;
            }
        }

        private void chbAddTop_CheckedChanged(object sender, EventArgs e)
        {
            if (chbAddTop.Checked)
            {
                tbAddTop.Enabled = true;
            }
            else
            {
                tbAddTop.Enabled = false;
            }
        }

        private void btnHtmlListUp_Click(object sender, EventArgs e)
        {
            if ((lbHtmlList.SelectedIndex != -1) && (lbHtmlList.SelectedIndex != 0))
            {
                object tmp = lbHtmlList.Items[lbHtmlList.SelectedIndex - 1];
                lbHtmlList.Items[lbHtmlList.SelectedIndex - 1] = lbHtmlList.Items[lbHtmlList.SelectedIndex];
                lbHtmlList.Items[lbHtmlList.SelectedIndex] = tmp;
                lbHtmlList.SelectedIndex = lbHtmlList.SelectedIndex - 1;
            }
        }

        private void btnHtmlListDown_Click(object sender, EventArgs e)
        {
            if ((lbHtmlList.SelectedIndex != -1) && (lbHtmlList.SelectedIndex != (lbHtmlList.Items.Count - 1)))
            {
                object tmp = lbHtmlList.Items[lbHtmlList.SelectedIndex + 1];
                lbHtmlList.Items[lbHtmlList.SelectedIndex + 1] = lbHtmlList.Items[lbHtmlList.SelectedIndex];
                lbHtmlList.Items[lbHtmlList.SelectedIndex] = tmp;
                lbHtmlList.SelectedIndex = lbHtmlList.SelectedIndex + 1;
            }
        }

        private void btnHtmlListDel_Click(object sender, EventArgs e)
        {
            if (lbHtmlList.SelectedIndex != -1)
            {
                bookInfo.Remove(lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString());
                lbHtmlList.Items.RemoveAt(lbHtmlList.SelectedIndex);
            }
        }

        private void lbHtmlList_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lbHtmlList.SelectedIndex == -1)
            {
                dataGridView1.DataSource = null;
            }
            else if (bookInfo.ContainsKey(lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString()))
            {
                dataGridView1.AutoGenerateColumns = true;
                dataGridView1.DataSource = bookInfo[lbHtmlList.Items[lbHtmlList.SelectedIndex].ToString()];
                dataGridView1.Columns[0].Width = 32;
                dataGridView1.Columns[1].Width = 45;
                dataGridView1.Columns[2].Width = 180;
                dataGridView1.Columns[3].Width = 70;
            }
        }

        private void btnOutputDir_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                tbOutputDir.Text = folderBrowserDialog1.SelectedPath;
            }
        }

        private void tbOutputDir_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            for (int i = 0; i < s.Length; i++)
            {
                if (Directory.Exists(s[i]))
                {
                    tbOutputDir.Text = s[i];
                    break;
                }
            }
        }

        private void checkBox3_CheckedChanged(object sender, EventArgs e)
        {
            if (!checkBox3.Checked) textBox2.Text = "webHelp";
            if (checkBox3.Checked) textBox2.Enabled = true;
            else textBox2.Enabled = false;
        }
    }
}
