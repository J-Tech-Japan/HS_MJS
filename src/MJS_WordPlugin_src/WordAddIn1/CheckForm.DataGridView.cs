using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// リファクタリング済
namespace WordAddIn1
{
    public partial class CheckForm
    {
        // 列インデックスの定数
        private const int TitleChangeColumnIndex = 7;
        private const int LinkColumnIndex = 8;

        // DataGridViewのセルクリックイベント
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (!IsValidCellClick(e)) return;

            string syori = "";
            int rowIndex = e.RowIndex;

            if (e.ColumnIndex == LinkColumnIndex)
            {
                syori = HandleLinkCellClick(rowIndex);
            }
            else if (e.ColumnIndex == TitleChangeColumnIndex)
            {
                syori = HandleTitleChangeCellClick(rowIndex);
            }
            else
            {
                return;
            }

            UpdateNewIdShow(syori, rowIndex);
            int bibMaxNum = GetMaxBibNumber();
            UpdateNewIdShowForSpecialRows(ref bibMaxNum);
            UpdateMergedIdShow();

            dataGridView1.DataSource = showResult;
            dataGridView1.Refresh();
            SetColor();
        }

        // セルクリックが有効か判定
        private bool IsValidCellClick(DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0 || e.ColumnIndex < 0) return false;
            var cellValue = dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value?.ToString();
            if ((e.ColumnIndex != LinkColumnIndex && e.ColumnIndex != TitleChangeColumnIndex) ||
                string.IsNullOrEmpty(cellValue) ||
                Regex.IsMatch(cellValue, @"^●"))
            {
                return false;
            }
            return true;
        }

        // リンクセルクリック時の処理
        private string HandleLinkCellClick(int rowIndex)
        {
            var cell = (DataGridViewLinkCell)dataGridView1.Rows[rowIndex].Cells[LinkColumnIndex];
            if (cell.Value == null || string.IsNullOrEmpty(cell.Value.ToString())) return string.Empty;
            cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
            cell.VisitedLinkColor = Color.Black;
            SetCellStyle(rowIndex, LinkColumnIndex, Color.LightGray, Color.Black);
            SetCellStyle(rowIndex, TitleChangeColumnIndex, Color.White, Color.Blue);
            ReplaceCellValue(rowIndex, TitleChangeColumnIndex, "^●", "○");
            return cell.Value.ToString();
        }

        // タイトル変更セルクリック時の処理
        private string HandleTitleChangeCellClick(int rowIndex)
        {
            var cell = (DataGridViewTextBoxCell)dataGridView1.Rows[rowIndex].Cells[TitleChangeColumnIndex];
            var cellRight = (DataGridViewLinkCell)dataGridView1.Rows[rowIndex].Cells[LinkColumnIndex];
            if (cell.Value == null || !cell.Value.ToString().Contains("タイトル変更")) return string.Empty;
            cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
            SetCellStyle(rowIndex, TitleChangeColumnIndex, Color.LightGray, Color.Black);
            SetCellStyle(rowIndex, LinkColumnIndex, Color.White, Color.Blue);
            cellRight.VisitedLinkColor = Color.Blue;
            ReplaceCellValue(rowIndex, LinkColumnIndex, "^●", "○");
            return cell.Value.ToString();
        }

        // セルの色・文字色を設定
        private void SetCellStyle(int rowIndex, int colIndex, Color backColor, Color foreColor)
        {
            dataGridView1[colIndex, rowIndex].Style.BackColor = backColor;
            dataGridView1[colIndex, rowIndex].Style.ForeColor = foreColor;
        }

        // セル値の置換
        private void ReplaceCellValue(int rowIndex, int colIndex, string pattern, string replacement)
        {
            var val = dataGridView1[colIndex, rowIndex].Value?.ToString();
            if (!string.IsNullOrEmpty(val))
            {
                dataGridView1[colIndex, rowIndex].Value = Regex.Replace(val, pattern, replacement);
            }
        }

        // new_id_showの更新
        private void UpdateNewIdShow(string syori, int rowIndex)
        {
            CheckInfo infoNow = showResult[rowIndex];
            if (syori.Contains("タイトル変更・結合追加"))
            {
                string[] idnewshows = infoNow.new_id_show.Split(new char[] { ' ' });
                infoNow.new_id_show = infoNow.old_id + (idnewshows.Length == 2 ? " " + idnewshows[1] : "");
            }
            else if (syori.Contains("タイトル変更・結合解除"))
            {
                string[] old_ids = infoNow.old_id.Split(new char[] { ' ' });
                infoNow.new_id_show = old_ids[0];
            }
            else if (syori.Contains("タイトル変更"))
            {
                infoNow.new_id_show = infoNow.old_id;
            }
        }

        // 最大bib番号の取得
        private int GetMaxBibNumber()
        {
            int bibMaxNum = 0;
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                CheckInfo info = showResult[i];
                var cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[TitleChangeColumnIndex];
                var cellValue = cell.Value?.ToString();
                if (!string.IsNullOrEmpty(info.new_id_show) &&
                    cellValue != "新規追加" &&
                    !cellValue.StartsWith("○タイトル変更") &&
                    cellValue != "見出しレベル変更")
                {
                    string idWithOutMerge = info.new_id_show.Split(new char[] { ' ' })[0];
                    if (idWithOutMerge.Length >= 3 && int.TryParse(idWithOutMerge.Substring(idWithOutMerge.Length - 3, 3), out int bibNum))
                    {
                        if (bibMaxNum < bibNum)
                        {
                            bibMaxNum = bibNum;
                        }
                    }
                }
            }
            return bibMaxNum;
        }

        // 新規追加・タイトル変更・見出しレベル変更行のnew_id_show更新
        private void UpdateNewIdShowForSpecialRows(ref int bibMaxNum)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    var cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[TitleChangeColumnIndex];
                    var cellValue = cell.Value?.ToString();
                    if (cellValue == "新規追加" || cellValue.StartsWith("○タイトル変更") || cellValue == "見出しレベル変更")
                    {
                        CheckInfo info = showResult[i];
                        bibMaxNum++;
                        var idCellValue = dataGridView1[6, i].Value?.ToString();
                        if (string.IsNullOrEmpty(idCellValue)) continue;
                        if (!idCellValue.Contains("#"))
                        {
                            string[] ids = idCellValue.Split(new char[] { ' ' });
                            info.new_id_show = ids[0].Substring(0, ids[0].Length - 3) + bibMaxNum.ToString("000") + (ids.Length == 2 ? " " + ids[1] : "");
                        }
                        else
                        {
                            for (int l = i - 1; l >= 0; l--)
                            {
                                var prevIdCellValue = dataGridView1[6, l].Value?.ToString();
                                if (!string.IsNullOrEmpty(prevIdCellValue) && !prevIdCellValue.Contains("#"))
                                {
                                    info.new_id_show = prevIdCellValue + "#" + idCellValue.Substring(idCellValue.Length - 8, 5) + bibMaxNum.ToString("000");
                                    break;
                                }
                            }
                        }
                    }
                }
                catch { continue; }
            }
        }

        // マージIDのnew_id_show更新
        private void UpdateMergedIdShow()
        {
            string numChange = "";
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                var cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[6];
                var cellValue = cell.Value?.ToString();
                if (!string.IsNullOrEmpty(cellValue))
                {
                    if (!cellValue.Contains("#"))
                        numChange = cellValue.Split(new char[] { ' ' })[0];
                    else
                    {
                        CheckInfo info = showResult[i];
                        info.new_id_show = numChange + "#" + cellValue.Split('#')[1];
                    }
                }
            }
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            SetColor();
        }

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            SetColor();
        }
    }
}
