using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CheckForm
    {
        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            string syori = "";

            if ((e.ColumnIndex != 8 && e.ColumnIndex != 7) ||
                Regex.IsMatch(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString(), @"^●") ||
                String.IsNullOrEmpty(dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString()))
            {
                return;
            }
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
                        cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
                        cell.VisitedLinkColor = Color.Black;
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.BackColor = Color.LightGray;
                        dataGridView1[7, cell.RowIndex].Style.BackColor = Color.White;
                        dataGridView1[7, cell.RowIndex].Style.ForeColor = Color.Blue;
                        dataGridView1[7, cell.RowIndex].Value = Regex.Replace(dataGridView1[7, cell.RowIndex].Value.ToString(), "^●", "○");
                    }
                }

                syori = cell.Value.ToString();
            }
            else
            {
                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[e.RowIndex].Cells[e.ColumnIndex];
                DataGridViewLinkCell cell_right = (DataGridViewLinkCell)dataGridView1.Rows[e.RowIndex].Cells[8];

                if (cell.Value == null || !cell.Value.ToString().Contains("タイトル変更"))
                {
                    return;
                }
                else
                {
                    if (!String.IsNullOrEmpty(cell.Value.ToString()))
                    {
                        cell.Value = Regex.Replace(cell.Value.ToString(), "^○", "●");
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.BackColor = Color.LightGray;
                        dataGridView1[cell.ColumnIndex, cell.RowIndex].Style.ForeColor = Color.Black;
                        dataGridView1[8, cell.RowIndex].Style.BackColor = Color.White;
                        cell_right.VisitedLinkColor = Color.Blue;
                        dataGridView1[8, cell.RowIndex].Value = Regex.Replace(dataGridView1[8, cell.RowIndex].Value.ToString(), "^●", "○");
                    }
                }

                syori = cell.Value.ToString();
            }

            int rowIndex = e.RowIndex;
            int bibMaxNum = 0;

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

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                CheckInfo info = showResult[i];

                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[7];
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

            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                try
                {
                    DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[7];
                    if (cell.Value.ToString() == "新規追加" || cell.Value.ToString().IndexOf("○タイトル変更") == 0 || cell.Value.ToString() == "見出しレベル変更")
                    {
                        CheckInfo info = showResult[i];
                        bibMaxNum++;
                        if (!dataGridView1[6, i].Value.ToString().Contains("#"))
                        {
                            string[] ids = dataGridView1[6, i].Value.ToString().Split(new char[] { ' ' });
                            info.new_id_show = ids[0].Substring(0, ids[0].Length - 3) + bibMaxNum.ToString("000") + (ids.Length == 2 ? " " + ids[1] : "");
                        }
                        else
                        {
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
                    continue;
                }
            }

            string numChange = "";
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                DataGridViewTextBoxCell cell = (DataGridViewTextBoxCell)dataGridView1.Rows[i].Cells[6];
                if (!String.IsNullOrEmpty(cell.Value.ToString()))
                {
                    if (!cell.Value.ToString().Contains(@"#"))
                        numChange = cell.Value.ToString().Split(new char[] { ' ' })[0];
                    else
                    {
                        CheckInfo info = showResult[i];
                        info.new_id_show = numChange + "#" + cell.Value.ToString().Split('#')[1];
                    }
                }
            }

            this.dataGridView1.DataSource = showResult;
            this.dataGridView1.Refresh();

            SetColor();
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
