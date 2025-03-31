using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Web;

namespace WordAddIn1
{
    public partial class setLink : Form
    {
        private string initText = "";
        public setLink()
        {
            InitializeComponent();

            Word.Application thisApp = WordAddIn1.Globals.ThisAddIn.Application;

            // set last using data
            tbFilePath.Text = WordAddIn1.Properties.Settings.Default.tbFilePathLast;      

            try
            {
                Word.Range selectionRange = thisApp.Selection.Range;
                tbDisplayStr.Text = selectionRange.Text;
                initText = selectionRange.Text;
                //tbDisplayStr.Text = Regex.Replace(selectionRange.Text, @"^[\s　]*(?:\d+\.)*\d+[\s　]+", "");
            }
            catch (Exception)
            {
            }
        }
        
        private void tbFilePath_TextChanged(object sender, EventArgs e)
        {
            if (File.Exists(tbFilePath.Text))
            {
                tbURL.Text = "";
                dataGridView1.Rows.Clear();
                dataGridView1.Enabled = true;
                using (StreamReader sr = new StreamReader(tbFilePath.Text))
                {
                    while(!sr.EndOfStream)
                    {
                        string[] lineStr = (sr.ReadLine()).Split('\t');
                        dataGridView1.Rows.Add();
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].ReadOnly = false;
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[0].Value = false;
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[1].Value = lineStr[0];
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[2].Value = lineStr[1];
                        dataGridView1.Rows[dataGridView1.Rows.Count - 1].Cells[3].Value = lineStr[2];
                    }
                }

                //System.Threading.Thread.Sleep(2000);
                Application.DoEvents();
            }
        }

        private void btnFileSelect_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "txt";

            try
            {
                if (Directory.Exists(Path.GetDirectoryName(tbFilePath.Text)))
                {
                    openFileDialog1.InitialDirectory = Path.GetDirectoryName(tbFilePath.Text);
                }
            }
            catch (Exception)
            {
            }

            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            tbFilePath.Text = openFileDialog1.FileName;
        }

        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.Index != e.RowIndex) row.Cells[0].Value = false;
                else row.Cells[0].Value = true;
            }
            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            Uri u1 = new Uri(Path.GetDirectoryName(Uri.UnescapeDataString(thisDocument.FullName)) + "\\");
            Uri u2 = new Uri(Path.GetDirectoryName(Path.GetDirectoryName(Uri.UnescapeDataString(tbFilePath.Text))) + "\\");
            string relativePath = u1.MakeRelativeUri(u2).ToString();
            //if (String.IsNullOrEmpty(relativePath)) relativePath = "";
            relativePath = Uri.UnescapeDataString(relativePath);
            relativePath = Regex.Replace(relativePath, @"@[^/]*?/", "/");
            if (Regex.IsMatch(Path.GetFileName(dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString()), @"^[A-Z]{3}0+$"))
            {
                tbURL.Text = Uri.UnescapeDataString(relativePath + "index.html");
                if(String.IsNullOrEmpty(initText))
                    tbDisplayStr.Text = Regex.Replace(thisDocument.Name, @"^[A-Z]{3}_([^_]*?)_.*?$", "$1");
            }
            else
            {
                if (Regex.IsMatch(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString(), @"^\d+$") && e.RowIndex + 1 < dataGridView1.RowCount - 1)
                {
                    if (dataGridView1.Rows[e.RowIndex + 1].Cells[3].Value.ToString().Contains("#"))
                        tbURL.Text = Uri.UnescapeDataString(relativePath + Regex.Replace(dataGridView1.Rows[e.RowIndex + 1].Cells[3].Value.ToString(), @"^(.*?)#(.*?)$", "$1.html#$2"));
                    else
                        tbURL.Text = Uri.UnescapeDataString(relativePath + dataGridView1.Rows[e.RowIndex + 1].Cells[3].Value.ToString() + ".html");
                }
                else
                {
                    if (dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString().Contains("#"))
                        tbURL.Text = Uri.UnescapeDataString(relativePath + Regex.Replace(dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString(), @"^(.*?)#(.*?)$", "$1.html#$2"));
                    else
                        tbURL.Text = Uri.UnescapeDataString(relativePath + dataGridView1.Rows[e.RowIndex].Cells[3].Value.ToString() + ".html");

                }
                if (String.IsNullOrEmpty(initText))
                {
                    if (!String.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString()))
                        tbDisplayStr.Text = dataGridView1.Rows[e.RowIndex].Cells[1].Value.ToString() + "　" + dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                    else
                        tbDisplayStr.Text = dataGridView1.Rows[e.RowIndex].Cells[2].Value.ToString();
                }
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (String.IsNullOrEmpty(tbDisplayStr.Text) && String.IsNullOrEmpty(tbURL.Text))
            {
                MessageBox.Show("表示文字列とURLを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (String.IsNullOrEmpty(tbDisplayStr.Text))
            {
                MessageBox.Show("表示文字列を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            else if (String.IsNullOrEmpty(tbURL.Text))
            {
                MessageBox.Show("URLを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            Word.Application thisApp = WordAddIn1.Globals.ThisAddIn.Application;
            Word.Document thisDocument = WordAddIn1.Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range selectionRange = thisApp.Selection.Range;
            selectionRange.Text = tbDisplayStr.Text;
            if (selectionRange.Hyperlinks.Count > 0) foreach (Word.Hyperlink hp in selectionRange.Hyperlinks) hp.Delete();
            thisDocument.Hyperlinks.Add(Anchor: selectionRange, Address: tbURL.Text);

            // save information
            WordAddIn1.Properties.Settings.Default.tbFilePathLast = tbFilePath.Text;
            WordAddIn1.Properties.Settings.Default.Save();

            this.Close();
        }

        private void tbDisplayStr_TextChanged(object sender, EventArgs e)
        {

        }
    }
}
