using System;
using System.Collections.Generic;
using System.Reflection;
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

            var version = Assembly.GetExecutingAssembly().GetName().Version;
            this.Text = $"MJSファイル結合ツール [V{version.Major}.{version.Minor}.{version.Build}]";
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (tabControl1.SelectedIndex == 1) this.Width = 800;
            else this.Width = 439;
        }
    }
}
