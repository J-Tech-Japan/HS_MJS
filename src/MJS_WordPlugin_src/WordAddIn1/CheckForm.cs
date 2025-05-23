using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

// リファクタリング済
namespace WordAddIn1
{
    public partial class CheckForm : Form
    {
        public RibbonMJS ribbon1;
        private List<CheckInfo> showResult;

        // コンストラクタ
        public CheckForm(RibbonMJS _ribbon1)
        {
            InitializeComponent();

            DialogResult = DialogResult.No;

            ribbon1 = _ribbon1;

            List<CheckInfo> checkResult = ribbon1.checkResult;

            showResult = new List<CheckInfo>();

            foreach (CheckInfo info in checkResult)
            {
                showResult.Add(info);
                //if (!string.IsNullOrEmpty(info.old_num) || !string.IsNullOrEmpty(info.new_num))
                //{
                //    showResult.Add(info);
                //}
            }

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = showResult;

        }

        // 「更新」ボタンのイベントハンドラ
        private void UpdateButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[6, i].Style.BackColor == Color.LightPink)
                {
                    MessageBox.Show("IDに不具合があります。");
                    return;
                }
            }
            // 重複チェック
            //    foreach (CheckInfo sr in showResult)
            //{
            //    if (string.IsNullOrEmpty(sr.new_id_show))
            //    {
            //        continue;
            //    }

            //    List<CheckInfo> checks = showResult.Where(p => p.new_id_show.Equals(sr.new_id_show)).ToList();

            //    if (checks.Count > 1)
            //    {
            //        MessageBox.Show(string.Format("重複するID{0}が存在します。ご確認ください。", sr.new_id_show));
            //        return;
            //    }
            //}            

            DialogResult = DialogResult.OK;
            Close();
        }

        // 「キャンセルボタン」のイベントハンドラ
        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        //private void CheckForm_SizeChanged(object sender, EventArgs e)
        //{
        //    Control form = (Control)sender;
        //    this.dataGridView1.Width = form.Size.Width - 38;
        //    this.dataGridView1.Height = form.Size.Height - 90;
        //}

    }
}
