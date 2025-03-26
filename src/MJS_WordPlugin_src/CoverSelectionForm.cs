using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WindowsFormsApp1
{
    public partial class FormSelectCoverTemplate : Form
    {
        public FormSelectCoverTemplate()
        {
            InitializeComponent();
        }

        private CoverTemplateEnum selectedCoverTemplate = CoverTemplateEnum.None;

        public CoverTemplateEnum SelectedCoverTemplate
        {
            get
            {
                return this.selectedCoverTemplate;
            }
        }

        private void CoverSelectionItem_OnSelectedStatusChanged(object sender, EventArgs e)
        {
            CoverSelectionItem item = (CoverSelectionItem)sender;

            if (sender == null)
            {
                return;
            }

            if (item.Selected)
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
                else if (item == this.CoverSelectionItemGeneralPattern3)
                {
                    this.selectedCoverTemplate = CoverTemplateEnum.GeneralPattern3;
                }
                else
                {
                    this.selectedCoverTemplate = CoverTemplateEnum.None;
                }

                if (item != this.CoverSelectionItemEasyCloud)
                {
                    this.CoverSelectionItemEasyCloud.Selected = false;
                }

                if (item != this.CoverSelectionItemEdgeTracker)
                {
                    this.CoverSelectionItemEdgeTracker.Selected = false;
                }

                if (item != this.CoverSelectionItemGeneralPattern1)
                {
                    this.CoverSelectionItemGeneralPattern1.Selected = false;
                }

                if (item != this.CoverSelectionItemGeneralPattern2)
                {
                    this.CoverSelectionItemGeneralPattern2.Selected = false;
                }

                if (item != this.CoverSelectionItemGeneralPattern3)
                {
                    this.CoverSelectionItemGeneralPattern3.Selected = false;
                }
            }
            else
            {
                this.selectedCoverTemplate = CoverTemplateEnum.None;
            }
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            if (this.selectedCoverTemplate == CoverTemplateEnum.None)
            {
                MessageBox.Show("Please select a cover template!");
            }
            else
            {
            this.DialogResult = DialogResult.OK;
                this.Close();
            }
        }

        private void BtnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        public enum CoverTemplateEnum
        {
            None,
            EasyCloud,
            EdgeTracker,
            GeneralPattern1,
            GeneralPattern2,
            GeneralPattern3
        }
    }
}
