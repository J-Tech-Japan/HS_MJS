// CheckForm.ExportCsv.cs

using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CheckForm
    {
        // 「CSV出力」ボタンのイベントハンドラ
        // フォルダ選択ダイアログを表示し、CSVファイルのエクスポートを実行
        private void ExportCsvButton_Click(object sender, EventArgs e)
        {
            // フォルダ選択ダイアログを表示
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
            {
                // 選択されたフォルダパスを取得
                string folderName = folderBrowserDialog1.SelectedPath;
                
                // CSVファイルにエクスポート実行
                ExportToCsvFile(folderName, showResult);
                
                MessageBox.Show("CSVファイルを作成しました。");
                
                // エクスプローラーで出力フォルダを開く
                System.Diagnostics.Process.Start(folderName);
            }
        }

        // CSVファイルへのエクスポート処理
        // 指定されたフォルダに書誌情報比較結果をCSV形式で出力
        private void ExportToCsvFile(string folderName, List<CheckInfo> result)
        {
            // タイムスタンプ付きのファイル名を生成
            string filePath = Path.Combine(folderName, "書誌情報比較_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv");
            
            // UTF-8エンコーディングでCSVファイルを作成
            using (StreamWriter docinfo = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                // CSVヘッダー行を出力
                docinfo.WriteLine("旧.項番,旧.タイトル,旧.ID,新.項番,新.タイトル,新.ID,新.ID（修正候補）,差異内容,修正処理（候補）,新規追加");
                
                // 各チェック結果を行として出力
                foreach (CheckInfo info in result)
                {
                    docinfo.WriteLine(BuildCsvLine(info));
                }
            }
        }

        // CSV行を生成するヘルパーメソッド
        // CheckInfoオブジェクトの各プロパティをカンマ区切りの文字列に変換
        private string BuildCsvLine(CheckInfo info)
        {
            StringBuilder sb = new StringBuilder();
            
            // 旧.項番
            if (info.old_num != null)
            {
                sb.Append(info.old_num);
            }
            sb.Append(",");
            
            // 旧.タイトル
            if (info.old_title != null)
            {
                sb.Append(info.old_title);
            }
            sb.Append(",");
            
            // 旧.ID
            if (info.old_id != null)
            {
                sb.Append(info.old_id);
            }
            sb.Append(",");
            
            // 新.項番
            if (info.new_num != null)
            {
                sb.Append(info.new_num);
            }
            sb.Append(",");
            
            // 新.タイトル
            if (info.new_title != null)
            {
                sb.Append(info.new_title);
            }
            sb.Append(",");
            
            // 新.ID
            if (info.new_id != null)
            {
                sb.Append(info.new_id);
            }
            sb.Append(",");
            
            // 新.ID（修正候補）
            if (info.new_id_show != null)
            {
                sb.Append(info.new_id_show);
            }
            sb.Append(",");
            
            // 差異内容
            if (info.diff != null)
            {
                sb.Append(info.diff);
            }
            sb.Append(",");
            
            // 修正処理（候補）
            if (info.editshow != null)
            {
                sb.Append(info.editshow);
            }
            sb.Append(",");
            
            // 新規追加
            if (info.edit != null)
            {
                sb.Append(info.edit);
            }
            
            return sb.ToString();
        }
    }
}
