using System.Reflection;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;

// リファクタリング済
namespace WordAddIn1
{
    public partial class RibbonMJS
    {
        private void BookInfoButton(object sender, RibbonControlEventArgs e)
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

        private void RibbonMJS_Load(object sender, RibbonUIEventArgs e)
        {
            // アセンブリのバージョンを取得
            var version = Assembly.GetExecutingAssembly().GetName().Version;
            string versionText = version.ToString(3); // "1.0.0" 形式で取得

            // labelVersion はリボンデザイナで追加したラベルの名前
            versionFileJoin.Label = $"\nバージョン\n{versionText}";
        }
    }
}
