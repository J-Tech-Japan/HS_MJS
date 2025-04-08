
using System;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CoverSelectionForm : Form
    {
        // 表紙テンプレートの列挙型
        public enum CoverTemplateEnum
        {
            None,
            EasyCloud,
            EdgeTracker,
            GeneralPattern1,
            GeneralPattern2,
            GeneralPattern3
        }

        // コンストラクタ: フォームの初期化を行います。
        public CoverSelectionForm()
        {
            InitializeComponent();
        }

        // 選択された表紙テンプレートを保持する変数
        private CoverTemplateEnum selectedCoverTemplate = CoverTemplateEnum.None;

        // 選択された表紙テンプレートを取得するプロパティ
        public CoverTemplateEnum SelectedCoverTemplate
        {
            get
            {
                return this.selectedCoverTemplate;
            }
        }

        // 表紙選択アイテムの選択状態が変更されたときに呼び出されるイベントハンドラ
        private void CoverSelectionItem_OnSelectedStatusChanged(object sender, EventArgs e)
        {
            // senderがnullの場合は処理を終了
            if (sender == null) return;

            CoverSelectionItem item = (CoverSelectionItem)sender;

            // アイテムが選択された場合の処理
            if (item.Selected)
            {
                // 選択されたアイテムに応じてselectedCoverTemplateを設定
                UpdateSelectedCoverTemplate(item);
                // 他のアイテムの選択状態を解除
                DeselectOtherItems(item);
            }
            else
            {
                // アイテムが選択解除された場合、selectedCoverTemplateをNoneに設定
                this.selectedCoverTemplate = CoverTemplateEnum.None;
            }
        }

        // 選択されたアイテムに応じてselectedCoverTemplateを設定するメソッド
        private void UpdateSelectedCoverTemplate(CoverSelectionItem item)
        {
            if (item == this.CoverSelectionItemEasyCloud)
            {
                this.selectedCoverTemplate = CoverTemplateEnum.EasyCloud;
            }
            else if (item == this.CoverSelectionItemEdgeTracker)
            {
                this.selectedCoverTemplate = CoverTemplateEnum.EdgeTracker;
            }
            else if (item == this.CoverSelectionItemGeneralPattern1)
            {
                this.selectedCoverTemplate = CoverTemplateEnum.GeneralPattern1;
            }
            else if (item == this.CoverSelectionItemGeneralPattern2)
            {
                this.selectedCoverTemplate = CoverTemplateEnum.GeneralPattern2;
            }
            else
            {
                this.selectedCoverTemplate = CoverTemplateEnum.None;
            }
        }

        // 他のアイテムの選択状態を解除するメソッド
        private void DeselectOtherItems(CoverSelectionItem selectedItem)
        {
            if (selectedItem != this.CoverSelectionItemEasyCloud)
            {
                this.CoverSelectionItemEasyCloud.Selected = false;
            }

            if (selectedItem != this.CoverSelectionItemEdgeTracker)
            {
                this.CoverSelectionItemEdgeTracker.Selected = false;
            }

            if (selectedItem != this.CoverSelectionItemGeneralPattern1)
            {
                this.CoverSelectionItemGeneralPattern1.Selected = false;
            }

            if (selectedItem != this.CoverSelectionItemGeneralPattern2)
            {
                this.CoverSelectionItemGeneralPattern2.Selected = false;
            }

            //if (selectedItem != this.CoverSelectionItemGeneralPattern3)
            //{
            //    this.CoverSelectionItemGeneralPattern3.Selected = false;
            //}
        }

        // OKボタンがクリックされたときのイベントハンドラ
        private void BtnOK_Click(object sender, EventArgs e)
        {
            // 表紙テンプレートが選択されていない場合の処理
            if (this.selectedCoverTemplate == CoverTemplateEnum.None)
            {
                MessageBox.Show("表紙のパターンをを選択してください。");
            }
            // 汎用パターン3が選択された場合の処理
            else if (this.selectedCoverTemplate == CoverTemplateEnum.GeneralPattern3)
            {
                MessageBox.Show("[汎用パターン3]テンプレートはまもなく登場します。");
            }
            else
            {
                // 選択されたテンプレートが有効な場合、ダイアログを閉じる
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        // キャンセルボタンがクリックされたときのイベントハンドラ
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            // ダイアログをキャンセルとして閉じる
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

    }
}

