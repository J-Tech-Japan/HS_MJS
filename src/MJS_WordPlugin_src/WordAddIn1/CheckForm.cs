using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CheckForm : Form
    {
        public RibbonMJS ribbon1;
        private List<CheckInfo> showResult;

        // コンストラクタ
        public CheckForm(RibbonMJS _ribbon1)
        {
            InitializeComponent();

            DialogResult = DialogResult.No;

            ribbon1 = _ribbon1;

            List<CheckInfo> checkResult = ribbon1.checkResult;

            showResult = new List<CheckInfo>();

            foreach (CheckInfo info in checkResult)
            {
                showResult.Add(info);
                //if (!string.IsNullOrEmpty(info.old_num) || !string.IsNullOrEmpty(info.new_num))
                //{
                //    showResult.Add(info);
                //}
            }

            dataGridView1.AutoGenerateColumns = false;
            dataGridView1.DataSource = showResult;

        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            SetColor();
        }

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
                //infoNow.new_id_show = infoNow.old_id;
            }
            else if (syori.Contains("タイトル変更・結合解除"))
            {
                string[] old_ids = infoNow.old_id.Split(new char[] { ' ' });
                infoNow.new_id_show = old_ids[0];
                //infoNow.new_id_show = infoNow.old_id;
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

            //if (syori.Contains("新規追加"))
            //{
            //    int bibMaxNumNow = bibMaxNum + 1;

            //    CheckInfo infoNow = showResult[rowIndex];

            //    string id_show_mae = infoNow.new_id_show;

            //    string id_show_new = infoNow.new_id_show.Substring(0, infoNow.new_id_show.Length - 3) + bibMaxNumNow.ToString("000");

            //    infoNow.new_id_show = id_show_new;

            //    if (!id_show_mae.Contains("#"))
            //    {
            //        for (int i = rowIndex + 1; i < showResult.Count; i++)
            //        {
            //            CheckInfo info = showResult[i];

            //            if (info.new_id_show.Contains("#"))
            //            {
            //                info.new_id_show = info.new_id_show.Replace(id_show_mae, id_show_new);
            //            }
            //        }
            //    }

            //    for (int i = rowIndex + 1; i < showResult.Count; i++)
            //    {
            //        CheckInfo info = showResult[i];

            //        if (!string.IsNullOrEmpty(info.new_id_show))
            //        {
            //            int bibNum = int.Parse(info.new_id_show.Substring(info.new_id_show.Length - 3, 3));
            //            if (bibMaxNum < bibNum || dataGridView1.Rows[i].Cells[7].Value.ToString() == "新規追加" || dataGridView1.Rows[i].Cells[7].Value.ToString() == "○タイトル変更")
            //            {
            //                bibMaxNumNow++;

            //                id_show_mae = info.new_id_show;

            //                id_show_new = id_show_mae.Substring(0, id_show_mae.Length - 3) + bibMaxNumNow.ToString("000");

            //                info.new_id_show = id_show_new;

            //                if (!id_show_mae.Contains("#"))
            //                {
            //                    for (int j = i + 1; j < showResult.Count; j++)
            //                    {
            //                        CheckInfo infoJ = showResult[j];

            //                        if (infoJ.new_id_show.Contains("#"))
            //                        {
            //                            infoJ.new_id_show = infoJ.new_id_show.Replace(id_show_mae, id_show_new);
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }

            //    this.dataGridView1.DataSource = showResult;
            //    this.dataGridView1.Refresh();

            //    setColor();
            //}
            //else if (syori.Contains("タイトル変更"))
            //{
            //    int bibMaxNumNow = bibMaxNum;

            //    CheckInfo infoNow = showResult[rowIndex];

            //    string id_show_mae = infoNow.new_id_show;

            //    string id_show_new = infoNow.new_id;

            //    infoNow.new_id_show = id_show_new;

            //    if (!id_show_mae.Contains("#"))
            //    {
            //        for (int i = rowIndex + 1; i < showResult.Count; i++)
            //        {
            //            CheckInfo info = showResult[i];

            //            if (info.new_id_show.Contains("#"))
            //            {
            //                info.new_id_show = info.new_id_show.Replace(id_show_mae, id_show_new);
            //            }
            //        }
            //    }

            //    bibMaxNumNow = 0;
            //    for (int i = rowIndex + 1; i < showResult.Count; i++)
            //    {
            //        CheckInfo info = showResult[i];
            //        if (info.diff == "新規追加") continue;
            //        int bibNum = int.Parse(info.new_id_show.Substring(info.new_id_show.Length - 3, 3));
            //        if (bibMaxNumNow < bibNum) bibMaxNumNow = bibNum;
            //    }
            //    bibMaxNum = bibMaxNumNow;

            //    for (int i = rowIndex + 1; i < showResult.Count; i++)
            //    {
            //        CheckInfo info = showResult[i];

            //        if (!string.IsNullOrEmpty(info.new_id_show))
            //        {
            //            int bibNum = int.Parse(info.new_id_show.Substring(info.new_id_show.Length - 3, 3));
            //            if (bibMaxNum < bibNum || info.diff == "新規追加")
            //            {
            //                bibMaxNumNow++;

            //                id_show_mae = info.new_id_show;

            //                id_show_new = id_show_mae.Substring(0, id_show_mae.Length - 3) + bibMaxNumNow.ToString("000");

            //                info.new_id_show = id_show_new;

            //                if (!id_show_mae.Contains("#"))
            //                {
            //                    for (int j = i + 1; j < showResult.Count; j++)
            //                    {
            //                        CheckInfo infoJ = showResult[j];

            //                        if (infoJ.new_id_show.Contains("#"))
            //                        {
            //                            infoJ.new_id_show = infoJ.new_id_show.Replace(id_show_mae, id_show_new);
            //                        }
            //                    }
            //                }
            //            }
            //        }
            //    }

            //    this.dataGridView1.DataSource = showResult;
            //    this.dataGridView1.Refresh();

            //    setColor();
            //}

            dataGridView1.DataSource = showResult;
            dataGridView1.Refresh();

            SetColor();
        }

        private void UpdateButton_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if (dataGridView1[6, i].Style.BackColor == Color.LightPink)
                {
                    MessageBox.Show("IDに不具合があります。");
                    return;
                }
            }
            // 重複チェック
            //    foreach (CheckInfo sr in showResult)
            //{
            //    if (string.IsNullOrEmpty(sr.new_id_show))
            //    {
            //        continue;
            //    }

            //    List<CheckInfo> checks = showResult.Where(p => p.new_id_show.Equals(sr.new_id_show)).ToList();

            //    if (checks.Count > 1)
            //    {
            //        MessageBox.Show(string.Format("重複するID{0}が存在します。ご確認ください。", sr.new_id_show));
            //        return;
            //    }
            //}            

            DialogResult = DialogResult.OK;
            Close();
        }

        // キャンセルボタンのクリックイベントハンドラ（旧: button3_Click）
        private void CancelButton_Click(object sender, EventArgs e)
        {
            Close();
        }

        //private void CheckForm_SizeChanged(object sender, EventArgs e)
        //{
        //    Control form = (Control)sender;
        //    this.dataGridView1.Width = form.Size.Width - 38;
        //    this.dataGridView1.Height = form.Size.Height - 90;
        //}

        private void dataGridView1_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            SetColor();
        }
    }
}
