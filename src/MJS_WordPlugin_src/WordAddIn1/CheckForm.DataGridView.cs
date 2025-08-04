// CheckForm.DataGridView.cs

using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CheckForm
    {
        // DataGridViewのセルクリックイベントハンドラ
        // セルの内容に応じて修正処理の状態を切り替え、ID番号の再計算を行う
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string syori = ""; // 選択された処理内容を格納

            // クリック対象が8列目（修正処理列）または7列目（差異内容列）以外の場合
            // または既に選択済み（●マーク）または空の場合は処理しない
            if ((e.ColumnIndex != 8 && e.ColumnIndex != 7) ||
                Regex.IsMatch(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), @"^●") ||
                String.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()))
            {
                return;
            }
            // 8列目（修正処理列）がクリックされた場合
            else if (e.ColumnIndex == 8)
            {
                DataGridViewLinkCell cell = (DataGridViewLinkCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                if (cell.Value == null)
                {
                    return;
                }
                else
                {
                    if (!String.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        // ○を●に変更（選択状態にする）
                        cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
                        cell.VisitedLinkColor = Color.Black;

                        // 選択されたセルの背景色を灰色に変更
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.BackColor = Color.LightGray;

                        // 7列目（差異内容列）の表示を未選択状態に戻す
                        dataGridView1[7, cell.RowIndex].Style.BackColor = Color.White;
                        dataGridView1[7, cell.RowIndex].Style.ForeColor = Color.Blue;
                        dataGridView1[7, cell.RowIndex].Value = Regex.Replace(dataGridView1[7, cell.RowIndex].Value.ToString(), "^●", "○");
                    }
                }

                syori = cell.Value.ToString();
            }
            // 7列目（差異内容列）がクリックされた場合
            else
            {
                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                DataGridViewLinkCell cell_right = (DataGridViewLinkCell)dataGridView1.Rows[e.RowIndex].Cells[8];

                // タイトル変更以外の場合は処理しない
                if (cell.Value == null || !cell.Value.ToString().Contains("タイトル変更"))
                {
                    return;
                }
                else
                {
                    if (!String.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        // ○を●に変更（選択状態にする）
                        cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
                        
                        // 選択されたセルの表示を変更
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.BackColor = Color.LightGray;
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.ForeColor = Color.Black;
                        
                        // 8列目（修正処理列）の表示を未選択状態に戻す
                        dataGridView1[8, cell.RowIndex].Style.BackColor = Color.White;
                        cell_right.VisitedLinkColor = Color.Blue;
                        dataGridView1[8, cell.RowIndex].Value = Regex.Replace(dataGridView1[8, cell.RowIndex].Value.ToString(), "^●", "○");
                    }
                }

                syori = cell.Value.ToString();
            }

            int rowIndex = e.RowIndex;
            int bibMaxNum = 0; // 書誌番号の最大値

            CheckInfo infoNow = showResult[rowIndex];

            // 選択された処理内容に応じてnew_id_showを設定
            if (syori.Contains("タイトル変更・結合追加"))
            {
                // 結合追加の場合：旧IDに結合部分を追加
                string[] idnewshows = infoNow.new_id_show.Split(new char[] { ' ' });
                infoNow.new_id_show = infoNow.old_id + (idnewshows.Length == 2 ? " " + idnewshows[1] : "");
            }
            else if (syori.Contains("タイトル変更・結合解除"))
            {
                // 結合解除の場合：旧IDの最初の部分のみを使用
                string[] old_ids = infoNow.old_id.Split(new char[] { ' ' });
                infoNow.new_id_show = old_ids[0];
            }
            else if (syori.Contains("タイトル変更"))
            {
                // タイトル変更の場合：旧IDをそのまま使用
                infoNow.new_id_show = infoNow.old_id;
            }

            // 現在の最大書誌番号を計算
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                CheckInfo info = showResult[i];

                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[7];
                
                // 新規追加・タイトル変更・見出しレベル変更以外の行から最大番号を取得
                if (!string.IsNullOrEmpty(info.new_id_show) && cell.Value.ToString() != "新規追加" && cell.Value.ToString().IndexOf("○タイトル変更") == -1 && cell.Value.ToString() != "見出しレベル変更")
                {
                    string idWithOutMerge = info.new_id_show.Split(new char[] { ' ' })[0];
                    int bibNum = int.Parse(idWithOutMerge.Substring(idWithOutMerge.Length - 3, 3));
                    if (bibMaxNum < bibNum)
                    {
                        bibMaxNum = bibNum;
                    }
                }
            }

            // 新規追加・タイトル変更・見出しレベル変更の行に新しい番号を割り当て
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[7];
                    if (cell.Value.ToString() == "新規追加" || cell.Value.ToString().IndexOf("○タイトル変更") == 0 || cell.Value.ToString() == "見出しレベル変更")
                    {
                        CheckInfo info = showResult[i];
                        bibMaxNum++; // 書誌番号をインクリメント
                        
                        // ハッシュ記号が含まれていない場合（通常のID）
                        if (!dataGridView1[6, i].Value.ToString().Contains("#"))
                        {
                            string[] ids = dataGridView1[6, i].Value.ToString().Split(new char[] { ' ' });
                            // 新しい書誌番号（3桁）で更新
                            info.new_id_show = ids[0].Substring(0, ids[0].Length - 3) + bibMaxNum.ToString("000") + (ids.Length == 2 ? " " + ids[1] : "");
                        }
                        // ハッシュ記号が含まれている場合（サブアイテム）
                        else
                        {
                            // 前の行から基準となるIDを探す
                            for (int l = i - 1; l >= 0; l--)
                            {
                                if (!String.IsNullOrEmpty(dataGridView1[6, l].Value.ToString()) && !dataGridView1[6, l].Value.ToString().Contains("#"))
                                {
                                    info.new_id_show = dataGridView1[6, l].Value.ToString() + "#" + dataGridView1[6, i].Value.ToString().Substring(dataGridView1[6, i].Value.ToString().Length - 8, 5) + bibMaxNum.ToString("000");
                                    break;
                                }
                            }
                        }
                    }
                }
                catch
                {
                    // エラーが発生した場合は次の行に進む
                    continue;
                }
            }

            // ハッシュ記号を含むIDの基準部分を更新
            string numChange = "";
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[6];
                if (!String.IsNullOrEmpty(cell.Value.ToString()))
                {
                    // ハッシュ記号が含まれていない場合は基準IDとして記録
                    if (!cell.Value.ToString().Contains(@"#"))
                        numChange = cell.Value.ToString().Split(new char[] { ' ' })[0];
                    // ハッシュ記号が含まれている場合は基準IDと組み合わせて更新
                    else
                    {
                        CheckInfo info = showResult[i];
                        info.new_id_show = numChange + "#" + cell.Value.ToString().Split('#')[1];
                    }
                }
            }

            // DataGridViewのデータソースを更新して再描画
            this.dataGridView1.DataSource = showResult;
            this.dataGridView1.Refresh();

            // セルの色設定を更新
            SetColor();
        }

        // DataGridViewのデータバインド完了イベントハンドラ
        // データバインド後にセルの色設定を適用
        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            SetColor();
        }

        /// DataGridViewのセル値変更イベントハンドラ
        /// セル値が変更された際にセルの色設定を更新
        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            SetColor();
        }
    }
}
