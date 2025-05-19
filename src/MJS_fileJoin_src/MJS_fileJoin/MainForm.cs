using System;
using System.Collections.Generic;
using System.Windows.Forms;

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

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1) this.Width = 800;
            else this.Width = 439;
        }
    }
}
