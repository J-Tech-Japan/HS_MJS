using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class bookInfo : Form
    {
        public bookInfo()
        {
            InitializeComponent();
        }

        private void btnEnter_Click(object sender, EventArgs e)
        {
            if (Regex.IsMatch(tbxDefaultValue.Text, "[０-９]"))
            {
                for (int i = 0; i < 10; i++)
                {
                    tbxDefaultValue.Text = Regex.Replace(tbxDefaultValue.Text, ((char)('０' + i)).ToString(), i.ToString());
                }
            }

            if (Regex.IsMatch(tbxDefaultValue.Text, @"^\d\d$"))
            {
                this.DialogResult = DialogResult.OK;
            }
            else if (Regex.IsMatch(tbxDefaultValue.Text, @"^\d$"))
            {
                tbxDefaultValue.Text = "0" + tbxDefaultValue.Text;
                this.DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("数字2桁でご指定ください。");
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            tbxDefaultValue.Text = "";
            this.DialogResult = DialogResult.Cancel;
        }

        private void bookInfo_Load_1(object sender, EventArgs e)
        {
            this.Visible = true;
        }

        private void tbxDefaultValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnEnter_Click(null, null);
            }

            //0～9と、バックスペース以外の時は、イベントをキャンセルする
            if ((e.KeyChar < '0' || '9' < e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }
    }
}
