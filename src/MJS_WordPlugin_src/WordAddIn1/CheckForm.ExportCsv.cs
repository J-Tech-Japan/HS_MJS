using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

// リファクタリング済
namespace WordAddIn1
{
    public partial class CheckForm
    {
        // 「CSV出力」ボタンのイベントハンドラ
        private void ExportCsvButton_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
            {
                string folderName = folderBrowserDialog1.SelectedPath;
                ExportToCsvFile(folderName, showResult);
                MessageBox.Show("CSVファイルを作成しました。");
                System.Diagnostics.Process.Start(folderName);
            }
        }

        // CSVファイルへのエクスポート処理
        private void ExportToCsvFile(string folderName, List<CheckInfo> result)
        {
            string filePath = Path.Combine(folderName, "書誌情報比較_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            using (StreamWriter docinfo = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                docinfo.WriteLine("旧.項番,旧.タイトル,旧.ID,新.項番,新.タイトル,新.ID,新.ID（修正候補）,差異内容,修正処理（候補）,新規追加");
                foreach (CheckInfo info in result)
                {
                    docinfo.WriteLine(BuildCsvLine(info));
                }
            }
        }

        // CSV行を生成するヘルパーメソッド
        private string BuildCsvLine(CheckInfo info)
        {
            StringBuilder sb = new StringBuilder();
            if (info.old_num != null)
            {
                sb.Append(info.old_num);
            }
            sb.Append(",");
            if (info.old_title != null)
            {
                sb.Append(info.old_title);
            }
            sb.Append(",");
            if (info.old_id != null)
            {
                sb.Append(info.old_id);
            }
            sb.Append(",");
            if (info.new_num != null)
            {
                sb.Append(info.new_num);
            }
            sb.Append(",");
            if (info.new_title != null)
            {
                sb.Append(info.new_title);
            }
            sb.Append(",");
            if (info.new_id != null)
            {
                sb.Append(info.new_id);
            }
            sb.Append(",");
            if (info.new_id_show != null)
            {
                sb.Append(info.new_id_show);
            }
            sb.Append(",");
            if (info.diff != null)
            {
                sb.Append(info.diff);
            }
            sb.Append(",");
            if (info.editshow != null)
            {
                sb.Append(info.editshow);
            }
            sb.Append(",");
            if (info.edit != null)
            {
                sb.Append(info.edit);
            }
            return sb.ToString();
        }
    }
}
