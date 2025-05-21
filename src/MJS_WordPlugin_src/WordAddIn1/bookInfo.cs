using System;
using System.Text.RegularExpressions;
using System.Windows.Forms;

// リファクタリング済
namespace WordAddIn1
{
    public partial class BookInfo : Form
    {
        // 正規表現パターンを定数として定義
        private const string TwoDigitPattern = @"^\d{2}$";
        private const string SingleDigitPattern = @"^\d$";
        private const string FullWidthDigitPattern = "[０-９]";

        public BookInfo()
        {
            InitializeComponent();
        }

        // 「Enter」ボタンがクリックされたときの処理
        private void btnEnter_Click(object sender, EventArgs e)
        {
            // 全角数字を半角数字に変換
            ConvertFullWidthDigits();

            // 入力が2桁の数字かどうかを確認
            if (IsValidTwoDigitNumber(tbxDefaultValue.Text))
            {
                DialogResult = DialogResult.OK;
            }
            // 入力が1桁の数字かどうかを確認
            else if (IsValidSingleDigitNumber(tbxDefaultValue.Text))
            {
                tbxDefaultValue.Text = "0" + tbxDefaultValue.Text;
                DialogResult = DialogResult.OK;
            }
            else
            {
                MessageBox.Show("数字2桁でご指定ください。");
            }
        }

        // 「キャンセル」ボタンがクリックされたときの処理
        private void btnCancel_Click(object sender, EventArgs e)
        {
            tbxDefaultValue.Text = string.Empty;
            DialogResult = DialogResult.Cancel;
        }

        // フォームがロードされたときの処理
        private void bookInfo_Load(object sender, EventArgs e)
        {
            Visible = true;
        }

        // テキストボックスでキーが押されたときの処理
        private void tbxDefaultValue_KeyPress(object sender, KeyPressEventArgs e)
        {
            // Enterキーが押された場合、「Enter」ボタンのクリック処理を呼び出す
            if (e.KeyChar == (char)Keys.Enter)
            {
                btnEnter_Click(null, null);
            }

            // 0～9とバックスペース以外のキー入力をキャンセルする
            if (!char.IsDigit(e.KeyChar) && e.KeyChar != '\b')
            {
                e.Handled = true;
            }
        }

        // 全角数字を半角数字に変換するメソッド
        private void ConvertFullWidthDigits()
        {
            if (Regex.IsMatch(tbxDefaultValue.Text, FullWidthDigitPattern))
            {
                for (int i = 0; i < 10; i++)
                {
                    tbxDefaultValue.Text = Regex.Replace(tbxDefaultValue.Text, ((char)('０' + i)).ToString(), i.ToString());
                }
            }
        }

        // 入力が2桁の数字かどうかを確認するメソッド
        private bool IsValidTwoDigitNumber(string input)
        {
            return Regex.IsMatch(input, TwoDigitPattern);
        }

        // 入力が1桁の数字かどうかを確認するメソッド
        private bool IsValidSingleDigitNumber(string input)
        {
            return Regex.IsMatch(input, SingleDigitPattern);
        }
    }
}