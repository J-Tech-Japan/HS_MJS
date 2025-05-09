using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void SetLinkButton(object sender, RibbonControlEventArgs e)
        {
            setLink stLink = new setLink();
            stLink.Show();
        }
    }
}
