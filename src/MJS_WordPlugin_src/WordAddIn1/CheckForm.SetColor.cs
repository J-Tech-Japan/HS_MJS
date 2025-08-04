// CheckForm.SetColor.cs

using System;
using System.Drawing;
using System.Linq;
using System.Text.RegularExpressions;

namespace WordAddIn1
{
    public partial class CheckForm
    {
        private void SetColor()
        {
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                CheckInfo check = showResult[i];
                SetDefaultColors(i);
                SetSameValueGray(i);
                SetIdShowYellow(i);
                SetIdShowMismatchYellow(i);
                SetTitleChangeGray(i);
                SetNumDiffRed(i);
                SetNullGray(i);
                SetNewNumColor(i, check);
                SetNewTitleColor(i, check);
                SetNewIdColor(i, check);
                SetNewIdShowRed(i, check);
                SetDiffColor(i, check);
                SetTitleChangeBlueUnderline(i);
                SetEditColor(i, check);
                SetDuplicateIdShowPink(i, check);
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
                SetIdShowFormatCheck(i, idBook, idNum);
                SetIdShowHashCheck(i);
            }
        }

        // SetColorのヘルパーメソッド
        private void SetDefaultColors(int i)
        {
            for (int col = 0; col <= 5; col++)
                dataGridView1[col, i].Style.ForeColor = Color.DarkGray;
        }

        private void SetSameValueGray(int i)
        {
            if (!String.IsNullOrEmpty(dataGridView1[0, i].Value.ToString()) &&
                !String.IsNullOrEmpty(dataGridView1[1, i].Value.ToString()) &&
                !String.IsNullOrEmpty(dataGridView1[2, i].Value.ToString()) &&
                dataGridView1[0, i].Value.ToString() == dataGridView1[3, i].Value.ToString() &&
                dataGridView1[1, i].Value.ToString() == dataGridView1[4, i].Value.ToString() &&
                dataGridView1[2, i].Value.ToString() == dataGridView1[5, i].Value.ToString())
            {
                for (int col = 0; col <= 5; col++)
                    dataGridView1[col, i].Style.ForeColor = Color.DarkGray;
            }
        }

        private void SetIdShowYellow(int i)
        {
            if (!String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                dataGridView1[5, i].Value.ToString() == dataGridView1[6, i].Value.ToString())
            {
                dataGridView1[6, i].Style.ForeColor = Color.DarkGray;
                dataGridView1[6, i].Style.BackColor = Color.LightYellow;
            }
        }

        private void SetIdShowMismatchYellow(int i)
        {
            if (!String.IsNullOrEmpty(dataGridView1[5, i].Value.ToString()) &&
                !String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                dataGridView1[5, i].Value.ToString() != dataGridView1[6, i].Value.ToString())
            {
                dataGridView1[6, i].Style.BackColor = Color.LightYellow;
            }
        }

        private void SetTitleChangeGray(int i)
        {
            if (dataGridView1[7, i].Value.ToString().Contains("●タイトル変更"))
            {
                dataGridView1[7, i].Style.BackColor = Color.LightGray;
            }
        }
        
        private void SetNumDiffRed(int i)
        {
            if (!String.IsNullOrEmpty(dataGridView1[0, i].Value.ToString()) &&
                !String.IsNullOrEmpty(dataGridView1[3, i].Value.ToString()) &&
                dataGridView1[0, i].Value.ToString() != dataGridView1[3, i].Value.ToString())
            {
                dataGridView1[0, i].Style.ForeColor = Color.Black;
                dataGridView1[3, i].Style.ForeColor = Color.Red;
            }
        }
        private void SetNullGray(int i)
        {
            for (int col = 0; col <= 2; col++)
            {
                if (dataGridView1[col, i].Value == null || "".Equals(dataGridView1[col, i].Value))
                    dataGridView1[col, i].Style.BackColor = Color.LightGray;
            }
            for (int col = 3; col <= 5; col++)
            {
                if (dataGridView1[col, i].Value == null || "".Equals(dataGridView1[col, i].Value))
                {
                    dataGridView1[col, i].Style.BackColor = Color.LightGray;
                    if (!String.IsNullOrEmpty(dataGridView1[col - 3, i].Value.ToString()))
                        dataGridView1[col - 3, i].Style.ForeColor = Color.Black;
                }
            }
            if (dataGridView1[6, i].Value == null || "".Equals(dataGridView1[6, i].Value))
                dataGridView1[6, i].Style.BackColor = Color.LightGray;
        }

        private void SetNewNumColor(int i, CheckInfo check)
        {
            if ("red".Equals(check.new_num_color))
            {
                dataGridView1[3, i].Style.ForeColor = Color.Red;
                dataGridView1[0, i].Style.ForeColor = Color.Black;
            }
            else if ("blue".Equals(check.new_num_color))
            {
                dataGridView1[3, i].Style.ForeColor = Color.Blue;
            }
        }

        private void SetNewTitleColor(int i, CheckInfo check)
        {
            if ("red".Equals(check.new_title_color))
            {
                dataGridView1[4, i].Style.ForeColor = Color.Red;
                dataGridView1[1, i].Style.ForeColor = Color.Black;
            }
            else if ("blue".Equals(check.new_title_color))
            {
                dataGridView1[4, i].Style.ForeColor = Color.Blue;
            }
        }

        private void SetNewIdColor(int i, CheckInfo check)
        {
            if ("red".Equals(check.new_id_color))
            {
                dataGridView1[5, i].Style.ForeColor = Color.Red;
                dataGridView1[2, i].Style.ForeColor = Color.Black;
            }
            else if ("blue".Equals(check.new_id_color))
            {
                dataGridView1[5, i].Style.ForeColor = Color.Blue;
            }
        }

        private void SetNewIdShowRed(int i, CheckInfo check)
        {
            if (check.new_id != null && check.new_id_show != null && !check.new_id.Equals(check.new_id_show))
            {
                dataGridView1[6, i].Style.ForeColor = Color.Red;
            }
        }

        private void SetDiffColor(int i, CheckInfo check)
        {
            if ("red".Equals(check.diff_color))
                dataGridView1[7, i].Style.ForeColor = Color.Red;
            else if ("blue".Equals(check.diff_color))
                dataGridView1[7, i].Style.ForeColor = Color.Blue;
        }

        private void SetTitleChangeBlueUnderline(int i)
        {
            if ("タイトル変更".Equals(dataGridView1[7, i].Value))
            {
                dataGridView1[7, i].Style.ForeColor = Color.Blue;
                dataGridView1[7, i].Style.Font = new Font("Ariel", 8, FontStyle.Underline);
            }
        }

        private void SetEditColor(int i, CheckInfo check)
        {
            if ("red".Equals(check.edit_color))
                dataGridView1[9, i].Style.ForeColor = Color.Red;
            else if ("blue".Equals(check.edit_color))
                dataGridView1[9, i].Style.ForeColor = Color.Blue;
        }

        private void SetDuplicateIdShowPink(int i, CheckInfo check)
        {
            if (!string.IsNullOrEmpty(check.new_id_show))
            {
                var checklist = showResult.Where(p => p.new_id_show.Equals(check.new_id_show)).ToList();
                if (checklist != null && checklist.Count > 1)
                {
                    dataGridView1[6, i].Style.BackColor = Color.LightPink;
                }
            }
        }

        private void SetIdShowFormatCheck(int i, string idBook, string idNum)
        {
            if (!String.IsNullOrEmpty(dataGridView1[6, i].Value.ToString()) &&
                !String.IsNullOrEmpty(idBook) && !String.IsNullOrEmpty(idNum) &&
                !Regex.IsMatch(dataGridView1[6, i].Value.ToString().Split(new char[] { '(' })[0].Trim(), @"^" + idBook + idNum + @"\d{3}$") &&
                !Regex.IsMatch(dataGridView1[6, i].Value.ToString().Split(new char[] { '(' })[0].Trim(), @"^" + idBook + idNum + @"\d{3}" + "#" + idBook + idNum + @"\d{3}$"))
                dataGridView1[6, i].Style.BackColor = Color.LightPink;
        }

        private void SetIdShowHashCheck(int i)
        {
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
                        else break;
                    }
                }
            }
        }
    }
}
