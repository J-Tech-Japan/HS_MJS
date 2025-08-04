// CoverSelectionItem.cs

using System;
using System.ComponentModel;
using System.Drawing;
using System.Windows.Forms;

namespace WordAddIn1
{
    // カバー選択用のユーザーコントロール
    // サムネイル画像とキャプション、選択状態を管理
    public partial class CoverSelectionItem : UserControl
    {
        // カバーのキャプション（タイトル）を取得・設定
        [Category("Flash")]
        [Description("The caption of the cover")]
        public string Caption
        {
            get
            {
                return LblCaption.Text;
            }
            set
            {
                LblCaption.Text = value;
                Invalidate();
            }
        }

        // カバーのサムネイル画像を取得・設定
        [Category("Flash")]
        [Description("The thumbnail image of the cover")]
        public Image CoverThumbnail
        {
            get
            {
                return ImgCover.Image;
            }
            set
            {
                ImgCover.Image = value;
                Invalidate();
            }
        }

        // 選択状態を保持するフィールド
        private bool selected = false;

        // カバーが選択されているかどうかを取得・設定
        [Category("Flash")]
        [Description("The selection status")]
        public bool Selected
        {
            get
            {
                return selected;
            }
            set
            {
                selected = value;

                UpdateSelectionStatusDisplay(selected);
            }
        }

        // 選択状態に応じて背景色などの表示を更新
        private void UpdateSelectionStatusDisplay(bool selected)
        {
            BackColor = selected ? Color.FromArgb(255, 255, 153) : Color.Transparent;

            Invalidate();
        }

        // 選択状態が変更されたときに発生するイベント
        [Browsable(true)]
        [Category("Action")]
        [Description("Invoked when user select")]
        public event EventHandler OnSelectedStatusChanged;

        // コントロールのコンストラクタ
        public CoverSelectionItem()
        {
            InitializeComponent();

            LblCaption.BackColor = Color.Transparent;
        }

        // コントロールがクリックされたときに選択状態を切り替え、イベントを発火
        private void CoverSelectionItem_MouseClick(object sender, EventArgs e)
        {
            selected = !selected;
            UpdateSelectionStatusDisplay(selected);

            if (OnSelectedStatusChanged != null)
            {
                OnSelectedStatusChanged(this, new EventArgs());
            }
        }
    }
}
