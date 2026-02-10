// SettingsForm.cs

using System;
using System.Windows.Forms;

namespace WordAddIn1
{
    /// <summary>
    /// アプリケーション設定を変更するダイアログフォーム
    /// </summary>
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        /// <summary>
        /// 現在の設定を読み込んでフォームに反映
        /// </summary>
        private void LoadSettings()
        {
            try
            {
                chkExtractHighQualityImages.Checked = ApplicationSettings.GetExtractHighQualityImagesSetting();
                chkBetaMode.Checked = ApplicationSettings.GetBetaModeSetting();
                
                numOutputScale.Value = (decimal)ApplicationSettings.GetOutputScaleMultiplierSetting();
                numTableImageScale.Value = (decimal)ApplicationSettings.GetTableImageScaleMultiplierSetting();
                numColumnImageScale.Value = (decimal)ApplicationSettings.GetColumnImageScaleMultiplierSetting();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"設定の読み込みに失敗しました。{Environment.NewLine}{ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// OKボタンクリック時の処理
        /// </summary>
        private void btnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // 設定を保存
                ApplicationSettings.SetExtractHighQualityImagesSetting(chkExtractHighQualityImages.Checked);
                ApplicationSettings.SetBetaModeSetting(chkBetaMode.Checked);
                
                if (!ApplicationSettings.SetOutputScaleMultiplierSetting((float)numOutputScale.Value))
                {
                    MessageBox.Show("通常画像スケール倍率の値が範囲外です。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!ApplicationSettings.SetTableImageScaleMultiplierSetting((float)numTableImageScale.Value))
                {
                    MessageBox.Show("表内画像スケール倍率の値が範囲外です。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!ApplicationSettings.SetColumnImageScaleMultiplierSetting((float)numColumnImageScale.Value))
                {
                    MessageBox.Show("コラム内画像スケール倍率の値が範囲外です。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"設定の保存に失敗しました。{Environment.NewLine}{ex.Message}",
                    "エラー",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
        }

        /// <summary>
        /// キャンセルボタンクリック時の処理
        /// </summary>
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        /// <summary>
        /// デフォルト値に戻すボタンクリック時の処理
        /// </summary>
        private void btnResetDefaults_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show(
                "すべての設定をデフォルト値に戻しますか？",
                "確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                chkExtractHighQualityImages.Checked = true;
                chkBetaMode.Checked = false;
                numOutputScale.Value = 1.4m;
                numTableImageScale.Value = 1.2m;
                numColumnImageScale.Value = 1.2m;
            }
        }
    }
}
