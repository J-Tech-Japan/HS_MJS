using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CheckForm : Form
    {
        public RibbonMJS ribbon1;

        private List<CheckInfo> showResult;

        public CheckForm(RibbonMJS _ribbon1)
        {
            InitializeComponent();

            this.DialogResult = DialogResult.No;

            ribbon1 = _ribbon1;

            List<CheckInfo> checkResult = ribbon1.checkResult;

            showResult = new List<CheckInfo>();

            foreach (CheckInfo info in checkResult)
            {
                showResult.Add(info);
            }

            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = showResult;

        }

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

            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
