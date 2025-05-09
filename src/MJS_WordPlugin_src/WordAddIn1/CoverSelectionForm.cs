
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
                return selectedCoverTemplate;
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
                selectedCoverTemplate = CoverTemplateEnum.None;
            }
        }

        // 選択されたアイテムに応じてselectedCoverTemplateを設定するメソッド
        private void UpdateSelectedCoverTemplate(CoverSelectionItem item)
        {
            if (item == CoverSelectionItemEasyCloud)
            {
                selectedCoverTemplate = CoverTemplateEnum.EasyCloud;
            }
            else
            {
                selectedCoverTemplate = item == CoverSelectionItemEdgeTracker
                    ? CoverTemplateEnum.EdgeTracker
                    : item == CoverSelectionItemGeneralPattern1
                                    ? CoverTemplateEnum.GeneralPattern1
                                    : item == CoverSelectionItemGeneralPattern2 ? CoverTemplateEnum.GeneralPattern2 : CoverTemplateEnum.None;
            }
        }

        // 他のアイテムの選択状態を解除するメソッド
        private void DeselectOtherItems(CoverSelectionItem selectedItem)
        {
            if (selectedItem != CoverSelectionItemEasyCloud)
            {
                CoverSelectionItemEasyCloud.Selected = false;
            }

            if (selectedItem != CoverSelectionItemEdgeTracker)
            {
                CoverSelectionItemEdgeTracker.Selected = false;
            }

            if (selectedItem != CoverSelectionItemGeneralPattern1)
            {
                CoverSelectionItemGeneralPattern1.Selected = false;
            }

            if (selectedItem != CoverSelectionItemGeneralPattern2)
            {
                CoverSelectionItemGeneralPattern2.Selected = false;
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
            if (selectedCoverTemplate == CoverTemplateEnum.None)
            {
                MessageBox.Show("表紙のパターンをを選択してください。");
            }
            // 汎用パターン3が選択された場合の処理
            else if (selectedCoverTemplate == CoverTemplateEnum.GeneralPattern3)
            {
                MessageBox.Show("[汎用パターン3]テンプレートはまもなく登場します。");
            }
            else
            {
                // 選択されたテンプレートが有効な場合、ダイアログを閉じる
                DialogResult = DialogResult.OK;
                Close();
            }
        }

        // キャンセルボタンがクリックされたときのイベントハンドラ
        private void BtnCancel_Click(object sender, EventArgs e)
        {
            // ダイアログをキャンセルとして閉じる
            DialogResult = DialogResult.Cancel;
            Close();
        }

    }
}

