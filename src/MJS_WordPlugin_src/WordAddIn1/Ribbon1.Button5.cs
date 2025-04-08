using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class Ribbon1
    {
        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            setLink stLink = new setLink();
            stLink.Show();
        }
    }
}
