// alert.cs

using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class Alert : Form
    {
        public Alert()
        {
            InitializeComponent();
            pictureBox1.WaitOnLoad = false;
            pictureBox1.LoadAsync("C:/Users/y-yonekura/Desktop/1amw.gif");
        }
    }
}
