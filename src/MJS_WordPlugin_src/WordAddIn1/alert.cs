using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class alert : Form
    {
        public alert()
        {
            InitializeComponent();
            pictureBox1.WaitOnLoad = false;
            pictureBox1.LoadAsync("C:/Users/y-yonekura/Desktop/1amw.gif");
        }
    }
}
