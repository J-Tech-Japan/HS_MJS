// SettingsForm.cs

using System;
using System.Windows.Forms;

namespace WordAddIn1
{
    // アプリケーション設定を変更するダイアログフォーム
    public partial class SettingsForm : Form
    {
        public SettingsForm()
        {
            InitializeComponent();
            LoadSettings();
        }

        // 現在の設定を読み込んでフォームに反映
        private void LoadSettings()
        {
            try
            {
                chkExtractHighQualityImages.Checked = ApplicationSettings.GetExtractHighQualityImagesSetting();
                chkBetaMode.Checked = ApplicationSettings.GetBetaModeSetting();
                numOutputScale.Value = (decimal)ApplicationSettings.GetOutputScaleMultiplierSetting();
                numTableImageScale.Value = (decimal)ApplicationSettings.GetTableImageScaleMultiplierSetting();
                numColumnImageScale.Value = (decimal)ApplicationSettings.GetColumnImageScaleMultiplierSetting();
                numMaxOutputWidth.Value = ApplicationSettings.GetMaxOutputWidthSetting();
                numMaxOutputHeight.Value = ApplicationSettings.GetMaxOutputHeightSetting();
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

        // OKボタンクリック時の処理
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
                
                if (!ApplicationSettings.SetMaxOutputWidthSetting((int)numMaxOutputWidth.Value))
                {
                    MessageBox.Show("出力画像の最大幅の値が範囲外です。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }
                
                if (!ApplicationSettings.SetMaxOutputHeightSetting((int)numMaxOutputHeight.Value))
                {
                    MessageBox.Show("出力画像の最大高さの値が範囲外です。", "警告", MessageBoxButtons.OK, MessageBoxIcon.Warning);
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
        
        // キャンセルボタンクリック時の処理
        private void btnCancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        // デフォルト値に戻すボタンクリック時の処理
        private void btnResetDefaults_Click(object sender, EventArgs e)
        {
            // ApplicationSettingsからデフォルト値を取得して設定
            var defaults = ApplicationSettings.GetDefaultValues();

            chkExtractHighQualityImages.Checked = defaults.ExtractHighQualityImages;
            chkBetaMode.Checked = defaults.IsBetaMode;
            numOutputScale.Value = (decimal)defaults.OutputScaleMultiplier;
            numTableImageScale.Value = (decimal)defaults.TableImageScaleMultiplier;
            numColumnImageScale.Value = (decimal)defaults.ColumnImageScaleMultiplier;
            numMaxOutputWidth.Value = defaults.MaxOutputWidth;
            numMaxOutputHeight.Value = defaults.MaxOutputHeight;
        }
    }
}
