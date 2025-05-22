using System;
using System.ComponentModel;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

// リファクタリング済
namespace WordAddIn1
{
    public partial class SetLink : Form
    {
        private string initText = string.Empty;

        public SetLink()
        {
            InitializeComponent();
            InitializeForm();
        }

        // 初期化処理
        private void InitializeForm()
        {
            var thisApp = Globals.ThisAddIn.Application;
            tbFilePath.Text = Properties.Settings.Default.tbFilePathLast;
            try
            {
                var selectionRange = thisApp.Selection.Range;
                tbDisplayStr.Text = selectionRange.Text;
                initText = selectionRange.Text;
            }
            catch { /* 無視 */ }
        }

        // ファイルパス変更時の処理
        private void tbFilePath_TextChanged(object sender, EventArgs e)
        {
            if (!File.Exists(tbFilePath.Text)) return;
            tbURL.Text = string.Empty;
            dataGridView1.Rows.Clear();
            dataGridView1.Enabled = true;
            LoadFileToGrid(tbFilePath.Text);
            Application.DoEvents();
        }

        // ファイル内容をDataGridViewにロード
        private void LoadFileToGrid(string filePath)
        {
            using (var sr = new StreamReader(filePath))
            {
                while (!sr.EndOfStream)
                {
                    var lineStr = sr.ReadLine()?.Split('\t');
                    if (lineStr == null || lineStr.Length < 3) continue;
                    int idx = dataGridView1.Rows.Add();
                    var row = dataGridView1.Rows[idx];
                    row.Cells[0].ReadOnly = false;
                    row.Cells[0].Value = false;
                    row.Cells[1].Value = lineStr[0];
                    row.Cells[2].Value = lineStr[1];
                    row.Cells[3].Value = lineStr[2];
                }
            }
        }

        // ファイル選択ボタン
        private void btnFileSelect_Click(object sender, EventArgs e)
        {
            openFileDialog1.DefaultExt = "txt";
            try
            {
                var dir = Path.GetDirectoryName(tbFilePath.Text);
                if (!string.IsNullOrEmpty(dir) && Directory.Exists(dir))
                    openFileDialog1.InitialDirectory = dir;
            }
            catch { /* 無視 */ }
            openFileDialog1.ShowDialog();
        }

        // ファイルダイアログOK時
        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            tbFilePath.Text = openFileDialog1.FileName;
        }

        // DataGridViewクリック時
        private void dataGridView1_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            SetRowSelection(e.RowIndex);
            var thisDocument = Globals.ThisAddIn.Application.ActiveDocument;
            string relativePath = GetRelativePath(thisDocument.FullName, tbFilePath.Text);
            string fileName = dataGridView1.Rows[e.RowIndex].Cells[3].Value?.ToString() ?? string.Empty;
            string displayStr = dataGridView1.Rows[e.RowIndex].Cells[1].Value?.ToString() ?? string.Empty;
            string subStr = dataGridView1.Rows[e.RowIndex].Cells[2].Value?.ToString() ?? string.Empty;
            string url = GenerateUrl(e.RowIndex, relativePath, fileName, displayStr);
            tbURL.Text = url;
            if (string.IsNullOrEmpty(initText))
            {
                tbDisplayStr.Text = !string.IsNullOrEmpty(displayStr)
                    ? displayStr + "　" + subStr
                    : subStr;
            }
        }

        // 選択行のチェックボックスをON
        private void SetRowSelection(int selectedIndex)
        {
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                row.Cells[0].Value = row.Index == selectedIndex;
            }
        }

        // 相対パス計算
        private string GetRelativePath(string docFullPath, string filePath)
        {
            var docDir = Path.GetDirectoryName(Uri.UnescapeDataString(docFullPath)) + "\\";
            var fileDir = Path.GetDirectoryName(Path.GetDirectoryName(Uri.UnescapeDataString(filePath))) + "\\";
            var u1 = new Uri(docDir);
            var u2 = new Uri(fileDir);
            string relativePath = u1.MakeRelativeUri(u2).ToString();
            relativePath = Uri.UnescapeDataString(relativePath);
            return Regex.Replace(relativePath, "@[^/]*?/", "/");
        }

        // URL生成
        private string GenerateUrl(int rowIndex, string relativePath, string fileName, string displayStr)
        {
            if (Regex.IsMatch(Path.GetFileName(fileName), @"^[A-Z]{3}0+$"))
            {
                return Uri.UnescapeDataString(relativePath + "index.html");
            }
            else
            {
                bool isNumber = Regex.IsMatch(displayStr, @"^\d+$");
                bool hasNext = rowIndex + 1 < dataGridView1.RowCount - 1;
                string nextFile = hasNext ? dataGridView1.Rows[rowIndex + 1].Cells[3].Value?.ToString() ?? string.Empty : string.Empty;
                if (isNumber && hasNext)
                {
                    return nextFile.Contains("#")
                        ? Uri.UnescapeDataString(relativePath + Regex.Replace(nextFile, @"^(.*?)#(.*?)$", "$1.html#$2"))
                        : Uri.UnescapeDataString(relativePath + nextFile + ".html");
                }
                else
                {
                    return fileName.Contains("#")
                        ? Uri.UnescapeDataString(relativePath + Regex.Replace(fileName, @"^(.*?)#(.*?)$", "$1.html#$2"))
                        : Uri.UnescapeDataString(relativePath + fileName + ".html");
                }
            }
        }

        // キャンセルボタン
        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        // OKボタン
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(tbDisplayStr.Text) && string.IsNullOrEmpty(tbURL.Text))
            {
                MessageBox.Show("表示文字列とURLを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(tbDisplayStr.Text))
            {
                MessageBox.Show("表示文字列を入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            if (string.IsNullOrEmpty(tbURL.Text))
            {
                MessageBox.Show("URLを入力してください。", "入力エラー", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }
            var thisApp = Globals.ThisAddIn.Application;
            var thisDocument = thisApp.ActiveDocument;
            var selectionRange = thisApp.Selection.Range;
            selectionRange.Text = tbDisplayStr.Text;
            if (selectionRange.Hyperlinks.Count > 0)
            {
                foreach (Word.Hyperlink hp in selectionRange.Hyperlinks) hp.Delete();
            }
            thisDocument.Hyperlinks.Add(Anchor: selectionRange, Address: tbURL.Text);
            Properties.Settings.Default.tbFilePathLast = tbFilePath.Text;
            Properties.Settings.Default.Save();
            Close();
        }

        private void tbDisplayStr_TextChanged(object sender, EventArgs e)
        {
            // 必要に応じて実装
        }
    }
}
