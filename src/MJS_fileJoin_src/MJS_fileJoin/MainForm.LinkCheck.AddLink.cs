using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace MJS_fileJoin
{
    public partial class MainForm
    {
        private void AddLinkCheckResult(string file, Match m, string titleName)
        {
            bool isMatch = titleName == m.Groups[2].Value;
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            lvi.SubItems.Add(m.Groups[2].Value);
            lvi.SubItems.Add(isMatch ? "true" : "false");
            lvi.SubItems.Add(titleName);
            lvi.SubItems.Add("true");
            if (!isMatch)
            {
                lvi.BackColor = Color.Red;
                //lvi.BackColor = Color.FromArgb(255, 192, 203);//#ffc0cb pink
                //allCheck += file + "," + m.Groups[1].Value + "," + m.Groups[2].Value + ",false," + titleName + ",true" + "\r\n";
                //Console.WriteLine(file + "\r\n" + m.Groups[1].Value + "\r\n" + m.Groups[2].Value + "\r\nfalse\r\n" + titleName + "\r\ntrue" + "\r\n");
            }
        }

        private void AddRefLinkCheckErrorResult(string file, string linkPage, Match m4, string content, int indexOfComma)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
            lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
            lvi.SubItems.Add("false");
            lvi.SubItems.Add("none");
            lvi.SubItems.Add("false");
            lvi.BackColor = Color.Red;
        }

        private void AddRefLinkCheckResult(string file, string linkPage, Match m4, string content, int indexOfComma, string titleName)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add($"{linkPage}.html#{m4.Groups[1].Value.Replace("_ref", "")}");
            lvi.SubItems.Add(content.Substring(indexOfComma + 1).Trim('\'', ' ').Replace("'", ""));
            lvi.SubItems.Add("true");
            lvi.SubItems.Add(titleName);
            lvi.SubItems.Add("true");
        }

        private void AddHttpLinkErrorResult(string file, Match m)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            try
            {
                lvi.SubItems.Add(GetStatusCode(m.Groups[1].Value).ToString());
            }
            catch
            {
                lvi.SubItems.Add("取得に失敗しました。");
            }
            lvi.SubItems.Add("");
            lvi.SubItems.Add("");
            lvi.SubItems.Add("");
            lvi.BackColor = Color.Red;
        }

        private void AddInvalidLinkResult(string file, Match m)
        {
            ListViewItem lvi = listView1.Items.Add(file);
            lvi.SubItems.Add(m.Groups[1].Value);
            lvi.SubItems.Add(m.Groups[2].Value);
            lvi.SubItems.Add("false");
            lvi.SubItems.Add("none");
            lvi.SubItems.Add("false");
            lvi.BackColor = Color.Red;
        }
    }
}
