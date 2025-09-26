namespace WordAddIn1
{
    partial class CoverSelectionForm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tableLayoutPanel1 = new System.Windows.Forms.TableLayoutPanel();
            this.FlowLayoutPanel1 = new System.Windows.Forms.FlowLayoutPanel();
            this.CoverSelectionItemEasyCloud = new WordAddIn1.CoverSelectionItem();
            this.CoverSelectionItemEdgeTracker = new WordAddIn1.CoverSelectionItem();
            this.CoverSelectionItemLucaTechGX = new WordAddIn1.CoverSelectionItem();
            this.FlowLayoutPanel2 = new System.Windows.Forms.FlowLayoutPanel();
            this.CoverSelectionItemGeneralPattern1 = new WordAddIn1.CoverSelectionItem();
            this.CoverSelectionItemGeneralPattern2 = new WordAddIn1.CoverSelectionItem();
            this.FlowLayoutPanel3 = new System.Windows.Forms.FlowLayoutPanel();
            this.BtnCancel = new System.Windows.Forms.Button();
            this.BtnOK = new System.Windows.Forms.Button();
            this.tableLayoutPanel1.SuspendLayout();
            this.FlowLayoutPanel1.SuspendLayout();
            this.FlowLayoutPanel2.SuspendLayout();
            this.FlowLayoutPanel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // tableLayoutPanel1
            // 
            this.tableLayoutPanel1.AutoSize = true;
            this.tableLayoutPanel1.ColumnCount = 1;
            this.tableLayoutPanel1.ColumnStyles.Add(new System.Windows.Forms.ColumnStyle(System.Windows.Forms.SizeType.Percent, 100F));
            this.tableLayoutPanel1.Controls.Add(this.FlowLayoutPanel1, 0, 0);
            this.tableLayoutPanel1.Controls.Add(this.FlowLayoutPanel2, 0, 1);
            this.tableLayoutPanel1.Controls.Add(this.FlowLayoutPanel3, 0, 2);
            this.tableLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.tableLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.tableLayoutPanel1.Name = "tableLayoutPanel1";
            this.tableLayoutPanel1.RowCount = 3;
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle());
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 23F));
            this.tableLayoutPanel1.RowStyles.Add(new System.Windows.Forms.RowStyle(System.Windows.Forms.SizeType.Absolute, 20F));
            this.tableLayoutPanel1.Size = new System.Drawing.Size(465, 346);
            this.tableLayoutPanel1.TabIndex = 0;
            // 
            // FlowLayoutPanel1
            // 
            this.FlowLayoutPanel1.AutoSize = true;
            this.FlowLayoutPanel1.Controls.Add(this.CoverSelectionItemEasyCloud);
            this.FlowLayoutPanel1.Controls.Add(this.CoverSelectionItemEdgeTracker);
            this.FlowLayoutPanel1.Controls.Add(this.CoverSelectionItemLucaTechGX);
            this.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FlowLayoutPanel1.Location = new System.Drawing.Point(0, 0);
            this.FlowLayoutPanel1.Margin = new System.Windows.Forms.Padding(0);
            this.FlowLayoutPanel1.Name = "FlowLayoutPanel1";
            this.FlowLayoutPanel1.Padding = new System.Windows.Forms.Padding(5, 5, 0, 0);
            this.FlowLayoutPanel1.Size = new System.Drawing.Size(465, 155);
            this.FlowLayoutPanel1.TabIndex = 0;
            // 
            // CoverSelectionItemEasyCloud
            // 
            this.CoverSelectionItemEasyCloud.BackColor = System.Drawing.Color.Transparent;
            this.CoverSelectionItemEasyCloud.Caption = "かんたんクラウド";
            this.CoverSelectionItemEasyCloud.CoverThumbnail = global::WordAddIn1.Properties.Resources.easy_cloud;
            this.CoverSelectionItemEasyCloud.Location = new System.Drawing.Point(5, 5);
            this.CoverSelectionItemEasyCloud.Margin = new System.Windows.Forms.Padding(0);
            this.CoverSelectionItemEasyCloud.Name = "CoverSelectionItemEasyCloud";
            this.CoverSelectionItemEasyCloud.Selected = false;
            this.CoverSelectionItemEasyCloud.Size = new System.Drawing.Size(150, 150);
            this.CoverSelectionItemEasyCloud.TabIndex = 0;
            this.CoverSelectionItemEasyCloud.OnSelectedStatusChanged += new System.EventHandler(this.CoverSelectionItem_OnSelectedStatusChanged);
            // 
            // CoverSelectionItemEdgeTracker
            // 
            this.CoverSelectionItemEdgeTracker.BackColor = System.Drawing.Color.Transparent;
            this.CoverSelectionItemEdgeTracker.Caption = "Edge Tracker";
            this.CoverSelectionItemEdgeTracker.CoverThumbnail = global::WordAddIn1.Properties.Resources.edge_tracker;
            this.CoverSelectionItemEdgeTracker.Location = new System.Drawing.Point(155, 5);
            this.CoverSelectionItemEdgeTracker.Margin = new System.Windows.Forms.Padding(0);
            this.CoverSelectionItemEdgeTracker.Name = "CoverSelectionItemEdgeTracker";
            this.CoverSelectionItemEdgeTracker.Selected = false;
            this.CoverSelectionItemEdgeTracker.Size = new System.Drawing.Size(150, 150);
            this.CoverSelectionItemEdgeTracker.TabIndex = 1;
            this.CoverSelectionItemEdgeTracker.OnSelectedStatusChanged += new System.EventHandler(this.CoverSelectionItem_OnSelectedStatusChanged);
            // 
            // CoverSelectionItemLucaTechGX
            // 
            this.CoverSelectionItemLucaTechGX.BackColor = System.Drawing.Color.Transparent;
            this.CoverSelectionItemLucaTechGX.Caption = "LucaTech GX";
            this.CoverSelectionItemLucaTechGX.CoverThumbnail = global::WordAddIn1.Properties.Resources.pattern3;
            this.CoverSelectionItemLucaTechGX.Location = new System.Drawing.Point(310, 5);
            this.CoverSelectionItemLucaTechGX.Margin = new System.Windows.Forms.Padding(0);
            this.CoverSelectionItemLucaTechGX.Name = "CoverSelectionItemLucaTechGX";
            this.CoverSelectionItemLucaTechGX.Selected = false;
            this.CoverSelectionItemLucaTechGX.Size = new System.Drawing.Size(150, 150);
            this.CoverSelectionItemLucaTechGX.TabIndex = 2;
            this.CoverSelectionItemLucaTechGX.OnSelectedStatusChanged += new System.EventHandler(this.CoverSelectionItem_OnSelectedStatusChanged);
            // 
            // FlowLayoutPanel2
            // 
            this.FlowLayoutPanel2.AutoSize = true;
            this.FlowLayoutPanel2.Controls.Add(this.CoverSelectionItemGeneralPattern1);
            this.FlowLayoutPanel2.Controls.Add(this.CoverSelectionItemGeneralPattern2);
            this.FlowLayoutPanel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FlowLayoutPanel2.Location = new System.Drawing.Point(0, 155);
            this.FlowLayoutPanel2.Margin = new System.Windows.Forms.Padding(0);
            this.FlowLayoutPanel2.Name = "FlowLayoutPanel2";
            this.FlowLayoutPanel2.Padding = new System.Windows.Forms.Padding(5, 5, 0, 0);
            this.FlowLayoutPanel2.Size = new System.Drawing.Size(465, 155);
            this.FlowLayoutPanel2.TabIndex = 1;
            // 
            // CoverSelectionItemGeneralPattern1
            // 
            this.CoverSelectionItemGeneralPattern1.BackColor = System.Drawing.Color.Transparent;
            this.CoverSelectionItemGeneralPattern1.Caption = "汎用パターン1";
            this.CoverSelectionItemGeneralPattern1.CoverThumbnail = global::WordAddIn1.Properties.Resources.pattern1;
            this.CoverSelectionItemGeneralPattern1.Location = new System.Drawing.Point(5, 5);
            this.CoverSelectionItemGeneralPattern1.Margin = new System.Windows.Forms.Padding(0);
            this.CoverSelectionItemGeneralPattern1.Name = "CoverSelectionItemGeneralPattern1";
            this.CoverSelectionItemGeneralPattern1.Selected = false;
            this.CoverSelectionItemGeneralPattern1.Size = new System.Drawing.Size(150, 150);
            this.CoverSelectionItemGeneralPattern1.TabIndex = 0;
            this.CoverSelectionItemGeneralPattern1.OnSelectedStatusChanged += new System.EventHandler(this.CoverSelectionItem_OnSelectedStatusChanged);
            // 
            // CoverSelectionItemGeneralPattern2
            // 
            this.CoverSelectionItemGeneralPattern2.BackColor = System.Drawing.Color.Transparent;
            this.CoverSelectionItemGeneralPattern2.Caption = "汎用パターン2";
            this.CoverSelectionItemGeneralPattern2.CoverThumbnail = global::WordAddIn1.Properties.Resources.pattern2;
            this.CoverSelectionItemGeneralPattern2.Location = new System.Drawing.Point(155, 5);
            this.CoverSelectionItemGeneralPattern2.Margin = new System.Windows.Forms.Padding(0);
            this.CoverSelectionItemGeneralPattern2.Name = "CoverSelectionItemGeneralPattern2";
            this.CoverSelectionItemGeneralPattern2.Selected = false;
            this.CoverSelectionItemGeneralPattern2.Size = new System.Drawing.Size(150, 150);
            this.CoverSelectionItemGeneralPattern2.TabIndex = 1;
            this.CoverSelectionItemGeneralPattern2.OnSelectedStatusChanged += new System.EventHandler(this.CoverSelectionItem_OnSelectedStatusChanged);
            // 
            // FlowLayoutPanel3
            // 
            this.FlowLayoutPanel3.AutoSize = true;
            this.FlowLayoutPanel3.Controls.Add(this.BtnCancel);
            this.FlowLayoutPanel3.Controls.Add(this.BtnOK);
            this.FlowLayoutPanel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.FlowLayoutPanel3.FlowDirection = System.Windows.Forms.FlowDirection.RightToLeft;
            this.FlowLayoutPanel3.Location = new System.Drawing.Point(3, 313);
            this.FlowLayoutPanel3.Name = "FlowLayoutPanel3";
            this.FlowLayoutPanel3.Padding = new System.Windows.Forms.Padding(0, 0, 3, 0);
            this.FlowLayoutPanel3.Size = new System.Drawing.Size(459, 30);
            this.FlowLayoutPanel3.TabIndex = 2;
            // 
            // BtnCancel
            // 
            this.BtnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.BtnCancel.Location = new System.Drawing.Point(378, 3);
            this.BtnCancel.Name = "BtnCancel";
            this.BtnCancel.Size = new System.Drawing.Size(75, 23);
            this.BtnCancel.TabIndex = 0;
            this.BtnCancel.Text = "キャンセル";
            this.BtnCancel.UseVisualStyleBackColor = true;
            this.BtnCancel.Click += new System.EventHandler(this.BtnCancel_Click);
            // 
            // BtnOK
            // 
            this.BtnOK.Location = new System.Drawing.Point(297, 3);
            this.BtnOK.Name = "BtnOK";
            this.BtnOK.Size = new System.Drawing.Size(75, 23);
            this.BtnOK.TabIndex = 1;
            this.BtnOK.Text = "OK";
            this.BtnOK.UseVisualStyleBackColor = true;
            this.BtnOK.Click += new System.EventHandler(this.BtnOK_Click);
            // 
            // CoverSelectionForm
            // 
            this.AcceptButton = this.BtnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(217)))), ((int)(((byte)(217)))), ((int)(((byte)(217)))));
            this.CancelButton = this.BtnCancel;
            this.ClientSize = new System.Drawing.Size(465, 346);
            this.Controls.Add(this.tableLayoutPanel1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.MaximizeBox = false;
            this.Name = "CoverSelectionForm";
            this.ShowIcon = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.Text = "ヘルプ表紙パターン選択";
            this.tableLayoutPanel1.ResumeLayout(false);
            this.tableLayoutPanel1.PerformLayout();
            this.FlowLayoutPanel1.ResumeLayout(false);
            this.FlowLayoutPanel2.ResumeLayout(false);
            this.FlowLayoutPanel3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TableLayoutPanel tableLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel1;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel2;
        private CoverSelectionItem CoverSelectionItemEasyCloud;
        private CoverSelectionItem CoverSelectionItemEdgeTracker;
        private CoverSelectionItem CoverSelectionItemGeneralPattern1;
        private CoverSelectionItem CoverSelectionItemGeneralPattern2;
        private CoverSelectionItem CoverSelectionItemLucaTechGX;
        private System.Windows.Forms.FlowLayoutPanel FlowLayoutPanel3;
        private System.Windows.Forms.Button BtnCancel;
        private System.Windows.Forms.Button BtnOK;
    }
}

