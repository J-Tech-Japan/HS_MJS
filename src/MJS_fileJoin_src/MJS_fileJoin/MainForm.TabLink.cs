using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
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
