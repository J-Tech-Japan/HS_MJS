using Microsoft.Office.Tools.Ribbon;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            loader load = new loader();
            load.Visible = false;
            if (!makeBookInfo(load))
            {
                load.Close();
                load.Dispose();
                return;
            }

            MessageBox.Show("出力が終了しました。");

            button4.Enabled = true;
            button2.Enabled = true;
            button5.Enabled = true;
        }
    }
}
