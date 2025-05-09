using System;
using System.ComponentModel;
using System.Drawing;
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
                return LblCaption.Text;
            }
            set
            {
                LblCaption.Text = value;
                Invalidate();
            }
        }

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

        private bool selected = false;

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

        private void UpdateSelectionStatusDisplay(bool selected)
        {
            BackColor = selected ? Color.FromArgb(255, 255, 153) : Color.Transparent;

            Invalidate();
        }

        [Browsable(true)]
        [Category("Action")]
        [Description("Invoked when user select")]
        public event EventHandler OnSelectedStatusChanged;

        public CoverSelectionItem()
        {
            InitializeComponent();

            LblCaption.BackColor = Color.Transparent;
        }

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
