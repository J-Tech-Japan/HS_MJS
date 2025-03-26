using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace WordAddIn1
{
    public partial class CoverSelectionItem : UserControl
    {
        [Category("Flash")]
        [Description("The caption of the cover")]
        public string Caption
        {
            get
            {
                return this.LblCaption.Text;
            }
            set
            {
                this.LblCaption.Text = value;
                Invalidate();
            }
        }

        [Category("Flash")]
        [Description("The thumbnail image of the cover")]
        public Image CoverThumbnail
        {
            get
            {
                return this.ImgCover.Image;
            }
            set
            {
                this.ImgCover.Image = value;
                Invalidate();
            }
        }

        private bool selected = false;

        [Category("Flash")]
        [Description("The selection status")]
        public bool Selected
        {
            get
            {
                return this.selected;
            }
            set
            {
                this.selected = value;

                this.UpdateSelectionStatusDisplay(this.selected);
            }
        }

        private void UpdateSelectionStatusDisplay(bool selected)
        {
            if (selected)
            {
                this.BackColor = Color.FromArgb(255, 255, 153);
            }
            else
            {
                this.BackColor = Color.Transparent;
            }

            Invalidate();
        }

        [Browsable(true)]
        [Category("Action")]
        [Description("Invoked when user select")]
        public event EventHandler OnSelectedStatusChanged;

        public CoverSelectionItem()
        {
            InitializeComponent();

            this.LblCaption.BackColor = Color.Transparent;
        }

        private void CoverSelectionItem_MouseClick(object sender, EventArgs e)
        {
            this.selected = !this.selected;
            this.UpdateSelectionStatusDisplay(this.selected);

            if (this.OnSelectedStatusChanged != null)
            {
                this.OnSelectedStatusChanged(this, new EventArgs());
            }
        }
    }
}
