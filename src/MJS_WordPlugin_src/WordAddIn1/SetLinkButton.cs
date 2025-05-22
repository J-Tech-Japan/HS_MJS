using Microsoft.Office.Tools.Ribbon;

namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void SetLinkButton(object sender, RibbonControlEventArgs e)
        {
            // 「リンク設定」のフォームを表示
            SetLink stLink = new SetLink();
            stLink.Show();
        }
    }
}
