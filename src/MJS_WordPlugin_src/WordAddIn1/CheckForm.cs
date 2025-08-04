// CheckForm.cs

using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;

namespace WordAddIn1
{
    // 書誌情報チェック結果を表示・確認するダイアログフォーム
    // 旧書誌情報と新書誌情報の比較結果をDataGridViewで表示し、
    // ユーザーが修正内容を確認・承認できる
    public partial class CheckForm : Form
    {
        public RibbonMJS ribbon1;

        // 表示用のチェック結果リスト
        private List<CheckInfo> showResult;

        // CheckFormのコンストラクタ
        // リボンクラスからチェック結果を受け取り、DataGridViewにバインドする
        public CheckForm(RibbonMJS _ribbon1)
        {
            InitializeComponent();

            // デフォルトのダイアログ結果をNoに設定
            this.DialogResult = DialogResult.No;

            // リボンクラスの参照を保存
            ribbon1 = _ribbon1;

            // チェック結果を取得
            List<CheckInfo> checkResult = ribbon1.checkResult;

            // 表示用リストを初期化
            showResult = new List<CheckInfo>();

            // チェック結果をコピーして表示用リストに追加
            foreach (CheckInfo info in checkResult)
            {
                showResult.Add(info);
            }

            // DataGridViewの設定
            this.dataGridView1.AutoGenerateColumns = false;
            this.dataGridView1.DataSource = showResult;
        }

        // 「更新」ボタンのクリックイベントハンドラ
        // IDに不具合（LightPinkの背景色）がないかチェックし、
        // 問題がなければダイアログを承認して閉じる
        private void UpdateButton_Click(object sender, EventArgs e)
        {
            // DataGridViewの全行をチェック
            for (int i = 0; i < dataGridView1.Rows.Count; i++)
            {
                // 6列目（ID列）の背景色がLightPinkの場合はエラー
                if (dataGridView1[6, i].Style.BackColor == Color.LightPink)
                {
                    MessageBox.Show("IDに不具合があります。");
                    return;
                }
            }

            // 問題がない場合はダイアログ結果をOKに設定して閉じる
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void CancelButton_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
