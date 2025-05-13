namespace WordAddIn1
{
    partial class RibbonMJS : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// 必要なデザイナー変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public RibbonMJS()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

        /// <summary> 
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースを破棄する場合は true を指定し、その他の場合は false を指定します。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region コンポーネント デザイナーで生成されたコード

        /// <summary>
        /// デザイナーのサポートに必要なメソッドです。
        /// このメソッドの内容をコード エディターで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.tab1 = this.Factory.CreateRibbonTab();
            this.group1 = this.Factory.CreateRibbonGroup();
            this.toggleButton1 = this.Factory.CreateRibbonToggleButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.button6 = this.Factory.CreateRibbonButton();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button7 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button8 = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.button9 = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button10 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group4.SuspendLayout();
            this.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "MJSワードプラグイン";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.toggleButton1);
            this.group1.Items.Add(this.button1);
            this.group1.Name = "group1";
            this.group1.Visible = false;
            // 
            // toggleButton1
            // 
            this.toggleButton1.Checked = true;
            this.toggleButton1.Image = global::WordAddIn1.Properties.Resources.edit;
            this.toggleButton1.Label = "編集モード";
            this.toggleButton1.Name = "toggleButton1";
            this.toggleButton1.ShowImage = true;
            // 
            // button1
            // 
            this.button1.Image = global::WordAddIn1.Properties.Resources.output;
            this.button1.Label = "編集承認";
            this.button1.Name = "button1";
            this.button1.ShowImage = true;
            this.button1.Visible = false;
            // 
            // group2
            // 
            this.group2.Items.Add(this.button6);
            this.group2.Items.Add(this.button4);
            this.group2.Items.Add(this.button7);
            this.group2.Items.Add(this.button5);
            this.group2.Items.Add(this.button8);
            this.group2.Name = "group2";
            // 
            // button6
            // 
            this.button6.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button6.Enabled = false;
            this.button6.Image = global::WordAddIn1.Properties.Resources.touka;
            this.button6.Label = " ";
            this.button6.Name = "button6";
            this.button6.ShowImage = true;
            // 
            // button4
            // 
            this.button4.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button4.Image = global::WordAddIn1.Properties.Resources.headerOutput;
            this.button4.Label = "書誌情報\n出力";
            this.button4.Name = "button4";
            this.button4.ShowImage = true;
            this.button4.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.BookInfoButton);
            // 
            // button7
            // 
            this.button7.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button7.Enabled = false;
            this.button7.Image = global::WordAddIn1.Properties.Resources.touka;
            this.button7.Label = " ";
            this.button7.Name = "button7";
            this.button7.ShowImage = true;
            // 
            // button5
            // 
            this.button5.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button5.Enabled = false;
            this.button5.Image = global::WordAddIn1.Properties.Resources.setLink;
            this.button5.Label = "リンク設定";
            this.button5.Name = "button5";
            this.button5.ShowImage = true;
            this.button5.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.SetLinkButton);
            // 
            // button8
            // 
            this.button8.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button8.Enabled = false;
            this.button8.Image = global::WordAddIn1.Properties.Resources.touka;
            this.button8.Label = " ";
            this.button8.Name = "button8";
            this.button8.ShowImage = true;
            // 
            // group4
            // 
            this.group4.Items.Add(this.button9);
            this.group4.Items.Add(this.button2);
            this.group4.Items.Add(this.button10);
            this.group4.Items.Add(this.button3);
            this.group4.Name = "group4";
            // 
            // button9
            // 
            this.button9.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button9.Enabled = false;
            this.button9.Image = global::WordAddIn1.Properties.Resources.touka;
            this.button9.Label = " ";
            this.button9.Name = "button9";
            this.button9.ShowImage = true;
            // 
            // button2
            // 
            this.button2.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button2.Enabled = false;
            this.button2.Image = global::WordAddIn1.Properties.Resources.styleCheck;
            this.button2.Label = "スタイルチェック";
            this.button2.Name = "button2";
            this.button2.ShowImage = true;
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.StyleCheckButton);
            // 
            // button10
            // 
            this.button10.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button10.Enabled = false;
            this.button10.Image = global::WordAddIn1.Properties.Resources.touka;
            this.button10.Label = " ";
            this.button10.Name = "button10";
            this.button10.ShowImage = true;
            // 
            // button3
            // 
            this.button3.ControlSize = Microsoft.Office.Core.RibbonControlSize.RibbonControlSizeLarge;
            this.button3.Enabled = false;
            this.button3.Image = global::WordAddIn1.Properties.Resources.htmlOutput;
            this.button3.Label = "HTML\n出力";
            this.button3.Name = "button3";
            this.button3.ShowImage = true;
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.GenerateHTMLButton);
            // 
            // RibbonMJS
            // 
            this.Name = "RibbonMJS";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonToggleButton toggleButton1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button7;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button8;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button9;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button10;
        //internal Microsoft.Office.Tools.Ribbon.RibbonButton button11;
    }

    partial class ThisRibbonCollection
    {
        internal RibbonMJS Ribbon1
        {
            get { return this.GetRibbon<RibbonMJS>(); }
        }
    }
}
