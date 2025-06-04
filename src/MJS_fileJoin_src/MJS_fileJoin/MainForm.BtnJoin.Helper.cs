using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private bool ValidateInput()
        {
            if (tbOutputDir.Text == "")
            {
                MessageBox.Show("出力ディレクトリをご指定ください。");
                return false;
            }
            if (!Directory.Exists(tbOutputDir.Text))
            {
                MessageBox.Show("出力ディレクトリが存在しません。");
                return false;
            }
            if (String.IsNullOrEmpty(textBox2.Text))
            {
                MessageBox.Show("格納フォルダをご指定ください。");
                return false;
            }
            if (lbHtmlList.Items.Count == 0)
            {
                MessageBox.Show("変換したHTMLファイルが格納されているフォルダーが登録されていません。");
                return false;
            }
            int fileCount = 0;
            foreach (string htmlDir in lbHtmlList.Items)
            {
                fileCount += bookInfo[htmlDir].Select("Column1 = true").Count();
            }
            if (fileCount == 0)
            {
                MessageBox.Show("コンテンツが選択されていません。");
                return false;
            }
            foreach (string htmlDir in lbHtmlList.Items)
            {
                if (!Directory.Exists(htmlDir))
                {
                    MessageBox.Show("「" + htmlDir + "」は削除されたか、名前が変更されています。");
                    return false;
                }
            }
            return true;
        }
    }
}
