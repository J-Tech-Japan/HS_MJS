using System.ComponentModel;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class loader : Form
    {
        public loader()
        {
            InitializeComponent();
        }

        private void backgroundWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            label1.Text = "aaaaa";
        }
    }
}
