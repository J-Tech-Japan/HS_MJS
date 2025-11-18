using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using DocMergerComponent;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void btnAddDoc_Click(object sender, EventArgs e)
        {
            openFileDialog2.FileName = "";
            openFileDialog2.Filter = "Wordファイル(*.doc;*.docx)|*.doc;*.docx|docファイル(*.doc)|*.doc|docxファイル(*.docx)|*.docx|すべてのファイル(*.*)|*.*";
            if (openFileDialog2.ShowDialog() == DialogResult.OK)
            {
                string extension = Path.GetExtension(openFileDialog2.FileName).ToLower();
                if (extension != ".doc" && extension != ".docx")
                {
                    MessageBox.Show("Wordファイル（.docまたは.docx）を選択してください。");
                    return;
                }
                listBox1.Items.Add(openFileDialog2.FileName);
            }
        }

        private void btnRemoveDoc_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex != -1)
            {
                listBox1.Items.RemoveAt(listBox1.SelectedIndex);
            }
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

        private void listBox1_DragDrop(object sender, DragEventArgs e)
        {
            string[] s = (string[])e.Data.GetData(DataFormats.FileDrop, false);
            foreach (string folder in s)
            {
                if (File.Exists(folder))
                {
                    string extension = Path.GetExtension(folder).ToLower();
                    if ((extension == ".doc" || extension == ".docx") &&
                        !folder.Contains("~$") && !listBox1.Items.Contains(folder))
                    {
                        listBox1.Items.Add(folder);
                    }
                }
                else if (Directory.Exists(folder))
                {
                    // .docと.docxの両方のファイルを取得
                    var docFiles = Directory.GetFiles(folder, "*.doc", SearchOption.AllDirectories);
                    var docxFiles = Directory.GetFiles(folder, "*.docx", SearchOption.AllDirectories);

                    // 両方の配列を結合
                    var allWordFiles = new string[docFiles.Length + docxFiles.Length];
                    docFiles.CopyTo(allWordFiles, 0);
                    docxFiles.CopyTo(allWordFiles, docFiles.Length);

                    foreach (string wordFile in allWordFiles)
                    {
                        if (!wordFile.Contains("~$") && !listBox1.Items.Contains(wordFile))
                            listBox1.Items.Add(wordFile);
                    }
                }
            }
        }

        private void listBox1_DragEnter(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effect = DragDropEffects.Copy;
            }
        }

        private void btnSelectOutputFolder_Click(object sender, EventArgs e)
        {
            if (folderBrowserDialog2.ShowDialog() == DialogResult.OK)
            {
                textBox1.Text = folderBrowserDialog2.SelectedPath;
            }
        }

        private void btnClearDocList_Click(object sender, EventArgs e)
        {
            while (listBox1.Items.Count != 0)
            {
                listBox1.Items.RemoveAt(0);
            }
            textBox1.Text = "";
        }

        private void btnJoinAndExport_Click(object sender, EventArgs e)
        {
            if (listBox1.Items.Count == 0)
            {
                MessageBox.Show("結合するWordファイルを指定してください。");
                return;
            }

            saveFileDialog2.Title = "結合済みWordファイル保存";
            saveFileDialog2.InitialDirectory = Path.GetDirectoryName(listBox1.Items[0].ToString());
            saveFileDialog2.FileName = Path.GetFileName(listBox1.Items[0].ToString());
            saveFileDialog2.Filter = "Word ファイル|*.doc|すべてのファイル|*.*";

            string strOutputDir = "";
            if (saveFileDialog2.ShowDialog() == DialogResult.OK)
                strOutputDir = saveFileDialog2.FileName;
            else
                return;

            groupBox1.Visible = true;
            
            // 単一ドキュメントかどうかで表示メッセージを変更
            if (listBox1.Items.Count == 1)
            {
                label10.Text = "ハイパーリンク更新中...";
            }
            else
            {
                label10.Text = "Word結合中...";
            }

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

            DocMerger objMerger = new DocMerger();
            bool cntItems = false;
            if (listBox1.Items.Count == 1) cntItems = true;

            objMerger.MergeFromFolders(strOrigFile, strCopiesDir, strOutputDir, this, true, cntItems);
            
            groupBox1.Visible = false;
            
            // 単一ドキュメントかどうかで完了メッセージを変更
            if (listBox1.Items.Count == 1)
            {
                MessageBox.Show("ハイパーリンクの更新が完了しました。");
                DialogResult selectMess = MessageBox.Show(strOutputDir + "\r\nが出力されました。\r\n出力したWordを表示しますか？", "ハイパーリンク更新成功", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
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
            else
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
        }
    }
}
