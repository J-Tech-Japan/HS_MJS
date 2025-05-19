using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using System.Xml.XPath;
using Microsoft.Office.Interop.Word;
using DocMergerComponent;
using Word = Microsoft.Office.Interop.Word;




namespace MJS_fileJoin
{
    public partial class MainForm : Form
    {
        public Dictionary<string, System.Data.DataTable> bookInfo = new Dictionary<string, System.Data.DataTable>();
        public string exportDir = "";
        public List<ListViewItem> logen = new List<ListViewItem>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void tbSelectJoinList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btnHtmlListFile_Click(object sender, EventArgs e)
        {
            if ((folderBrowserDialog1.ShowDialog() == DialogResult.OK) && !lbHtmlList.Items.Contains(folderBrowserDialog1.SelectedPath))
            {
                addHtmlDir(folderBrowserDialog1.SelectedPath);
            }
        }
        
        private void lbHtmlList_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            List<string> webHelpFol = new List<string>();
            foreach (string folder in s)
            {
                if (File.Exists(folder)) continue;
                if (Path.GetFileName(folder) == "webHelp")
                    webHelpFol.Add(folder);
                else
                {
                    string[] fol = Directory.GetDirectories(folder, "webHelp", SearchOption.AllDirectories);
                    foreach (string webhelp in fol) webHelpFol.Add(webhelp);
                }
            }

            for (int i = 0; i < webHelpFol.Count; i++)
            {
                if (!addHtmlDir(webHelpFol[i])) continue;
            }
        }

        private void lbHtmlList_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        

        

        

        private void tbOutputDir_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btnClear_Click(object sender, EventArgs e)
        {
            tbSelectJoinList.Text = "";
            chbChangeTitle.Checked = false;
            tbChangeTitle.Text = "";
            tbChangeTitle.Enabled = false;
            chbAddTop.Checked = false;
            tbAddTop.Text = "";
            tbAddTop.Enabled = false;
            while (lbHtmlList.Items.Count != 0)
            {
                lbHtmlList.Items.RemoveAt(0);
            }
            dataGridView1.DataSource = null;
            bookInfo.Clear();
            tbOutputDir.Text = "";
            chbListOutput.Checked = false;
        }

        private void btnAddDoc_Click(object sender, EventArgs e)
        {
            openFileDialog2.FileName = "";
            openFileDialog2.Filter = "docファイル(*.doc)|*.doc|すべてのファイル(*.*)|*.*";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                if (Path.GetExtension(openFileDialog2.FileName) != ".doc")
                {
                    MessageBox.Show("docファイルを選択してください。");
                    return;
                }
                listBox1.Items.Add(openFileDialog2.FileName);
            }
            checkItems();
        }

        private void btnRemoveDoc_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
            checkItems();
        }

        private void btnDocUp_Click(object sender, EventArgs e)
        {
            if ((listBox1.SelectedIndex != -1) && (listBox1.SelectedIndex != 0))
            {
                object tmp = listBox1.Items[listBox1.SelectedIndex - 1];
                listBox1.Items[listBox1.SelectedIndex - 1] = listBox1.Items[listBox1.SelectedIndex];
                listBox1.Items[listBox1.SelectedIndex] = tmp;
                listBox1.SelectedIndex = listBox1.SelectedIndex - 1;
            }
        }

        private void btnDocDown_Click(object sender, EventArgs e)
        {
            if ((listBox1.SelectedIndex != -1) && (listBox1.SelectedIndex != (listBox1.Items.Count - 1)))
            {
                object tmp = listBox1.Items[listBox1.SelectedIndex + 1];
                listBox1.Items[listBox1.SelectedIndex + 1] = listBox1.Items[listBox1.SelectedIndex];
                listBox1.Items[listBox1.SelectedIndex] = tmp;
                listBox1.SelectedIndex = listBox1.SelectedIndex + 1;
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string folder in s)
            {
                if (File.Exists(folder) && Path.GetExtension(folder) == ".doc")
                {
                    if (!folder.Contains("~$") && !listBox1.Items.Contains(folder))
                        listBox1.Items.Add(folder);
                }
                else if (Directory.Exists(folder))
                {
                    string[] fol = Directory.GetFiles(folder, "*.doc", SearchOption.AllDirectories);
                    foreach (string webhelp in fol)
                    {
                        if (!webhelp.Contains("~$") && !listBox1.Items.Contains(webhelp))
                            listBox1.Items.Add(webhelp);
                    }
                }
            }
            checkItems();
        }

        private void btnSelectOutputFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog2.SelectedPath;
            }
        }
        private void checkItems()
        {
            if (listBox1.Items.Count > 1) checkBox4.Checked = true;
            else checkBox4.Checked = false;
            if (listBox1.Items.Count == 1) checkBox5.Checked = true;
        }

        private void btnClearDocList_Click(object sender, EventArgs e)
        {
            while (listBox1.Items.Count != 0)
            {
                listBox1.Items.RemoveAt(0);
            }
            textBox1.Text = "";
            checkBox4.Checked = false;
            checkBox5.Checked = false;
        }

        private void btnJoinAndExport_Click(object sender, EventArgs e)
        {
            if (!checkBox4.Checked && !checkBox5.Checked)
            {
                MessageBox.Show("出力するファイル種別を選択してください。");
                return;
            }
            //if(checkBox4.Checked && listBox1.Items.Count <= 1)
            //{
            //    MessageBox.Show("結合するWordファイルを複数指定してください。");
            //    return;
            //}
            //if(textBox1.Text == "")
            //{
            //    MessageBox.Show("出力先を指定してください。");
            //    return;
            //}
            if (checkBox4.Checked)
                saveFileDialog2.Title = "結合済みWordファイル保存";
            else
                saveFileDialog2.Title = "PDFファイル保存";

            saveFileDialog2.InitialDirectory = Path.GetDirectoryName(listBox1.Items[0].ToString());
            if (!checkBox4.Checked)
                saveFileDialog2.FileName = Path.GetFileName(listBox1.Items[0].ToString()).Replace(".doc", ".pdf");
            else
                saveFileDialog2.FileName = Path.GetFileName(listBox1.Items[0].ToString());

            if (checkBox4.Checked)
                saveFileDialog2.Filter = "Word ファイル|*.doc|すべてのファイル|*.*";
            else
                saveFileDialog2.Filter = "PDF ファイル|*.pdf|すべてのファイル|*.*";

            string strOutputDir = "";
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                strOutputDir = saveFileDialog2.FileName.Replace(".pdf", ".doc");
            else
                return;

            groupBox1.Visible = true;
            if (checkBox4.Checked)
                label10.Text = "Word結合中...";
            else
                label10.Text = "PDF出力中...";

            string strOrigFile = "";
            List<string> strCopiesDir = new List<string>();
            bool bl = false;
            foreach (string file in listBox1.Items)
            {
                if (!bl)
                {
                    strOrigFile = file;
                    bl = true;
                    continue;
                }
                strCopiesDir.Add(file);
            }
            //string strOutputDir = textBox1.Text + "\\" + Path.GetFileName(listBox1.Items[0].ToString());
            DocMerger objMerger = new DocMerger();
            bool cntItems = false;
            if (listBox1.Items.Count == 1) cntItems = true;
            objMerger.Merge(strOrigFile, strCopiesDir, strOutputDir, this, checkBox4.Checked, checkBox5.Checked, cntItems);
            groupBox1.Visible = false;
            if (checkBox4.Checked)
            {
                MessageBox.Show("Wordの結合が完了しました。");
                DialogResult selectMess = MessageBox.Show(strOutputDir + "\r\nが出力されました。\r\n出力したWordを表示しますか？", "Word結合成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectMess == DialogResult.Yes)
                {
                    try
                    {
                        Process.Start(strOutputDir);
                    }
                    catch
                    {
                        MessageBox.Show("Wordの出力に失敗しました。", "Word出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
            if (checkBox5.Checked)
            {
                MessageBox.Show("PDFの出力が完了しました。");
                DialogResult selectMess = MessageBox.Show(strOutputDir.Replace(".doc", ".pdf") + "\r\nが出力されました。\r\n出力したPDFを表示しますか？", "PDF結合成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                if (selectMess == DialogResult.Yes)
                {
                    try
                    {
                        Process.Start(strOutputDir.Replace(".doc", ".pdf"));
                    }
                    catch
                    {
                        MessageBox.Show("PDFの出力に失敗しました。", "PDF出力失敗", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void listBox2_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void listBox2_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string folder in s)
            {
                if (File.Exists(folder)) continue;
                bool bl = false;
                foreach (string lbi in listBox2.Items)
                    if (lbi == folder) bl = true;
                if (bl == false) listBox2.Items.Add(folder);
            }
        }

        private void btnListBox2Up_Click(object sender, EventArgs e)
        {
            if ((listBox2.SelectedIndex != -1) && (listBox2.SelectedIndex != 0))
            {
                object tmp = listBox2.Items[listBox2.SelectedIndex - 1];
                listBox2.Items[listBox2.SelectedIndex - 1] = listBox2.Items[listBox2.SelectedIndex];
                listBox2.Items[listBox2.SelectedIndex] = tmp;
                listBox2.SelectedIndex = listBox2.SelectedIndex - 1;
            }
        }

        private void btnListBox2Down_Click(object sender, EventArgs e)
        {
            if ((listBox2.SelectedIndex != -1) && (listBox2.SelectedIndex != (listBox2.Items.Count - 1)))
            {
                object tmp = listBox2.Items[listBox2.SelectedIndex + 1];
                listBox2.Items[listBox2.SelectedIndex + 1] = listBox2.Items[listBox2.SelectedIndex];
                listBox2.Items[listBox2.SelectedIndex] = tmp;
                listBox2.SelectedIndex = listBox2.SelectedIndex + 1;
            }
        }

        private void btnRemoveListBox2Item_Click(object sender, EventArgs e)
        {
            if (listBox2.SelectedIndex != -1)
            {
                listBox2.Items.RemoveAt(listBox2.SelectedIndex);
            }
        }

        private void btnCheckLinksInFolders_Click(object sender, EventArgs e)
        {
            listView1.Items.Clear();
            button13.Enabled = false;
            foreach (string folder in listBox2.Items)
            {
                if (Directory.Exists(folder)) LinkCheck(folder);
                label14.Visible = false;
                progressBar2.Visible = false;
            }
            foreach (ListViewItem lvi in listView1.Items)
                if (lvi.BackColor == Color.Red) button13.Enabled = true;
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1) this.Width = 800;
            else this.Width = 439;
        }

        private void listView1_DoubleClick(object sender, EventArgs e)
        {
            //if (Regex.IsMatch(Path.GetFileName(listView1.SelectedItems[0].Text), @"^\w{3}\d{5}.html$"))
            //{
            //    System.Diagnostics.Process.Start(Path.GetDirectoryName(listView1.SelectedItems[0].Text) + @"\index.html", "#t=" + Path.GetFileName(listView1.SelectedItems[0].Text));
            //}
            System.Diagnostics.Process.Start(listView1.SelectedItems[0].Text);
        }

        private void addFolderToListBox2_Click(object sender, EventArgs e)
        {
            if ((folderBrowserDialog3.ShowDialog() == DialogResult.OK) && !listBox2.Items.Contains(folderBrowserDialog3.SelectedPath))
            {
                listBox2.Items.Add(folderBrowserDialog3.SelectedPath);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked)
            {
                List<ListViewItem> lo = new List<ListViewItem>();
                foreach (ListViewItem lvi in listView1.Items)
                    if (lvi.BackColor == Color.Red || lvi.BackColor == Color.FromArgb(255, 192, 203))
                        lo.Add(lvi);
                listView1.Items.Clear();
                foreach (ListViewItem o in lo)
                    listView1.Items.Add(o);
            }
            else
            {
                if (logen.Count > 0)
                {
                    listView1.Items.Clear();
                    foreach (ListViewItem lvi in logen)
                        listView1.Items.Add(lvi);
                }
            }
        }

        private void exportListViewToCSV_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
                ListViewToCSV(listView1, saveFileDialog1.FileName, true);
        }

        
    }
}
