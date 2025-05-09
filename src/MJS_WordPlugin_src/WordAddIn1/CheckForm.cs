using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class CheckForm : Form
    {
        public Ribbon1 ribbon1;

        private List<CheckInfo> showResult;

        public CheckForm(Ribbon1 _ribbon1)
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
            setColor();
        }

        private void setColor()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                CheckInfo check = showResult[i];
                dataGridView1[0, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[1, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[2, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[3, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[4, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[5, i].Style.ForeColor = Color.DarkGray;
                if (!String.IsNullOrEmpty(dataGridView1[0, i].Value.ToString()) &&
                    !String.IsNullOrEmpty(dataGridView1[1, i].Value.ToString()) &&
                    !String.IsNullOrEmpty(dataGridView1[2, i].Value.ToString()) &&
                    dataGridView1[0, i].Value.ToString() == dataGridView1[3, i].Value.ToString() &&
                    dataGridView1[1, i].Value.ToString() == dataGridView1[4, i].Value.ToString() &&
                    dataGridView1[2, i].Value.ToString() == dataGridView1[5, i].Value.ToString())
                {
                    dataGridView1[0, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[1, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[2, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[3, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[4, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[5, i].Style.ForeColor = Color.DarkGray;
                }
                if (!String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                    dataGridView1[5, i].Value.ToString() == dataGridView1[6, i].Value.ToString()
                    )
                {
                    dataGridView1[6, i].Style.ForeColor = Color.DarkGray;
                    dataGridView1[6, i].Style.BackColor = Color.LightYellow;
                }
                if (!String.IsNullOrEmpty(dataGridView1[5, i].Value.ToString()) &&
                    !String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                    dataGridView1[5, i].Value.ToString() != dataGridView1[6, i].Value.ToString()
                    )
                {
                    dataGridView1[6, i].Style.BackColor = Color.LightYellow;
                }
                if (dataGridView1[7, i].Value.ToString().Contains("●タイトル変更"))
                {
                    dataGridView1[7, i].Style.BackColor = Color.LightGray;
                }
                if (!String.IsNullOrEmpty(dataGridView1[0, i].Value.ToString()) &&
                    !String.IsNullOrEmpty(dataGridView1[3, i].Value.ToString()) &&
                    dataGridView1[0, i].Value.ToString() != dataGridView1[3, i].Value.ToString()
                    )
                {
                    dataGridView1[0, i].Style.ForeColor = Color.Black;
                    dataGridView1[3, i].Style.ForeColor = Color.Red;
                }
                if (dataGridView1[0, i].Value == null || "".Equals(dataGridView1[0, i].Value))
                {
                    dataGridView1[0, i].Style.BackColor = Color.LightGray;
                }
                if (dataGridView1[1, i].Value == null || "".Equals(dataGridView1[1, i].Value))
                {
                    dataGridView1[1, i].Style.BackColor = Color.LightGray;
                }
                if (dataGridView1[2, i].Value == null || "".Equals(dataGridView1[2, i].Value))
                {
                    dataGridView1[2, i].Style.BackColor = Color.LightGray;
                }

                if (dataGridView1[3, i].Value == null || "".Equals(dataGridView1[3, i].Value))
                {
                    dataGridView1[3, i].Style.BackColor = Color.LightGray;
                    if (!String.IsNullOrEmpty(dataGridView1[0, i].Value.ToString()))
                        dataGridView1[0, i].Style.ForeColor = Color.Black;
                }
                if (dataGridView1[4, i].Value == null || "".Equals(dataGridView1[4, i].Value))
                {
                    dataGridView1[4, i].Style.BackColor = Color.LightGray;
                    if (!String.IsNullOrEmpty(dataGridView1[1, i].Value.ToString()))
                        dataGridView1[1, i].Style.ForeColor = Color.Black;
                }
                if (dataGridView1[5, i].Value == null || "".Equals(dataGridView1[5, i].Value))
                {
                    dataGridView1[5, i].Style.BackColor = Color.LightGray;
                    if (!String.IsNullOrEmpty(dataGridView1[2, i].Value.ToString()))
                        dataGridView1[2, i].Style.ForeColor = Color.Black;
                }
                if (dataGridView1[6, i].Value == null || "".Equals(dataGridView1[6, i].Value))
                {
                    dataGridView1[6, i].Style.BackColor = Color.LightGray;
                }
               
                if ("red".Equals(check.new_num_color))
                {
                    dataGridView1[3, i].Style.ForeColor = Color.Red;
                    dataGridView1[0, i].Style.ForeColor = Color.Black;
                }
                else if ("blue".Equals(check.new_num_color))
                {
                    dataGridView1[3, i].Style.ForeColor = Color.Blue;
                }
                else
                {
                    //dataGridView1[3, i].Style.ForeColor = Color.Black;
                }
                if ("red".Equals(check.new_title_color))
                {
                    dataGridView1[4, i].Style.ForeColor = Color.Red;
                    dataGridView1[1, i].Style.ForeColor = Color.Black;
                }
                else if ("blue".Equals(check.new_title_color))
                {
                    dataGridView1[4, i].Style.ForeColor = Color.Blue;
                }
                else
                {
                    //dataGridView1[4, i].Style.ForeColor = Color.Black;
                }
                if ("red".Equals(check.new_id_color))
                {
                    dataGridView1[5, i].Style.ForeColor = Color.Red;
                    dataGridView1[2, i].Style.ForeColor = Color.Black;
                }
                else if ("blue".Equals(check.new_id_color))
                {
                    dataGridView1[5, i].Style.ForeColor = Color.Blue;
                }
                else
                {
                    //dataGridView1[5, i].Style.ForeColor = Color.Black;
                }

                if (check.new_id != null && check.new_id_show != null
                    && !check.new_id.Equals(check.new_id_show))
                {
                    dataGridView1[6, i].Style.ForeColor = Color.Red;
                }
                else
                {
                    //dataGridView1[6, i].Style.ForeColor = Color.Black;
                }

                if ("red".Equals(check.diff_color))
                {
                    dataGridView1[7, i].Style.ForeColor = Color.Red;
                }
                else if ("blue".Equals(check.diff_color))
                {
                    dataGridView1[7, i].Style.ForeColor = Color.Blue;
                }
                if ("タイトル変更".Equals(dataGridView1[7, i].Value))
                {
                    dataGridView1[7, i].Style.ForeColor = Color.Blue;
                    dataGridView1[7, i].Style.Font = new Font("Ariel", 8, FontStyle.Underline);
                }

                if ("red".Equals(check.edit_color))
                {
                    dataGridView1[9, i].Style.ForeColor = Color.Red;
                }
                else if ("blue".Equals(check.edit_color))
                {
                    dataGridView1[9, i].Style.ForeColor = Color.Blue;
                }

                if (!string.IsNullOrEmpty(check.new_id_show))
                {
                    var checklist = showResult.Where(p => p.new_id_show.Equals(check.new_id_show)).ToList();
                    if (checklist != null && checklist.Count > 1)
                    {
                        dataGridView1[6, i].Style.BackColor = Color.LightPink;
                    }
                    //    else
                    //    {
                    //        dataGridView1[6, i].Style.BackColor = Color.White;
                    //    }
                    //}
                    //else
                    //{
                    //    dataGridView1[6, i].Style.BackColor = Color.White;
                }
            }

            string idBook = "";
            string idNum = "";
            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                idBook = Regex.Replace(dataGridView1[2, i].Value.ToString(), @"^([A-Z]{3}).*?$", "$1");
                idNum = Regex.Replace(dataGridView1[2, i].Value.ToString(), @"^[A-Z]{3}(\d{2}).*?$", "$1");
                if (!String.IsNullOrEmpty(idBook) && !String.IsNullOrEmpty(idNum)) break;
            }

            for (int i = 1; i < dataGridView1.Rows.Count; i++)
            {
                if (!String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                    !String.IsNullOrEmpty(idBook) && !String.IsNullOrEmpty(idNum) &&
                    !Regex.IsMatch(dataGridView1[6, i].Value.ToString().Split(new char[] { '(' })[0].Trim(), @"^" + idBook + idNum + @"\d{3}$") &&
                    !Regex.IsMatch(dataGridView1[6, i].Value.ToString().Split(new char[] { '(' })[0].Trim(), @"^" + idBook + idNum + @"\d{3}" + "#" + idBook + idNum + @"\d{3}$"))
                    dataGridView1[6, i].Style.BackColor = Color.LightPink;

                if (i == 0 && dataGridView1[6, i].Value.ToString().Contains("#"))
                {
                    dataGridView1[6, i].Style.BackColor = Color.LightPink;
                }
                if (i > 0 && dataGridView1[6, i].Value.ToString().Contains("#"))
                {
                    for (int l = i; l >= 0; l--)
                    {
                        if (!String.IsNullOrEmpty(dataGridView1[6, l].Value.ToString()) && !dataGridView1[6, l].Value.ToString().Contains("#"))
                        {
                            if (Regex.Replace(dataGridView1[6, i].Value.ToString(), "#.*$", "") != Regex.Replace(dataGridView1[6, l].Value.ToString(), " *\\(.*\\)", ""))
                            {
                                dataGridView1[6, i].Style.BackColor = Color.LightPink;
                            }
                            else  break;
                        }
                    }
                }
            }
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

            setColor();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (DialogResult.OK == folderBrowserDialog1.ShowDialog())
            {
                string folderName = folderBrowserDialog1.SelectedPath;

                using (StreamWriter docinfo = new StreamWriter(folderName + "\\" + "書誌情報比較_" + DateTime.Now.ToString("yyyyMMddHHmmss") + ".csv", false, Encoding.UTF8))
                {
                    docinfo.WriteLine("旧.項番,旧.タイトル,旧.ID,新.項番,新.タイトル,新.ID,新.ID（修正候補）,差異内容,修正処理（候補）,新規追加");

                    foreach (CheckInfo info in showResult)
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

                        docinfo.WriteLine(sb.ToString());
                    }
                }

                MessageBox.Show("CSVファイルを作成しました。");

                System.Diagnostics.Process.Start(folderName);
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                if(dataGridView1[6, i].Style.BackColor == Color.LightPink)
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

        private void button3_Click(object sender, EventArgs e)
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
            setColor();
        }
    }
}
