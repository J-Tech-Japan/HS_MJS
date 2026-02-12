// SettingsForm.Designer.cs

namespace WordAddIn1
{
    partial class SettingsForm
    {
        /// <summary>
        /// Required designer variable variable.
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
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.chkBetaMode = new System.Windows.Forms.CheckBox();
            this.chkExtractHighQualityImages = new System.Windows.Forms.CheckBox();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.numColumnImageScale = new System.Windows.Forms.NumericUpDown();
            this.label3 = new System.Windows.Forms.Label();
            this.numTableImageScale = new System.Windows.Forms.NumericUpDown();
            this.label2 = new System.Windows.Forms.Label();
            this.numOutputScale = new System.Windows.Forms.NumericUpDown();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.numMaxOutputHeight = new System.Windows.Forms.NumericUpDown();
            this.label5 = new System.Windows.Forms.Label();
            this.numMaxOutputWidth = new System.Windows.Forms.NumericUpDown();
            this.label4 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.btnResetDefaults = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColumnImageScale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTableImageScale)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numOutputScale)).BeginInit();
            this.groupBox3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxOutputHeight)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxOutputWidth)).BeginInit();
            this.SuspendLayout();
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.chkBetaMode);
            this.groupBox1.Controls.Add(this.chkExtractHighQualityImages);
            this.groupBox1.Location = new System.Drawing.Point(12, 12);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(460, 90);
            this.groupBox1.TabIndex = 0;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "機能設定";
            // 
            // chkBetaMode
            // 
            this.chkBetaMode.AutoSize = true;
            this.chkBetaMode.Location = new System.Drawing.Point(15, 55);
            this.chkBetaMode.Name = "chkBetaMode";
            this.chkBetaMode.Size = new System.Drawing.Size(310, 16);
            this.chkBetaMode.TabIndex = 1;
            this.chkBetaMode.Text = "詳細ログとCSV出力を有効にする";
            this.chkBetaMode.UseVisualStyleBackColor = true;
            // 
            // chkExtractHighQualityImages
            // 
            this.chkExtractHighQualityImages.AutoSize = true;
            this.chkExtractHighQualityImages.Location = new System.Drawing.Point(15, 25);
            this.chkExtractHighQualityImages.Name = "chkExtractHighQualityImages";
            this.chkExtractHighQualityImages.Size = new System.Drawing.Size(204, 16);
            this.chkExtractHighQualityImages.TabIndex = 0;
            this.chkExtractHighQualityImages.Text = "高画質画像抽出機能を有効にする";
            this.chkExtractHighQualityImages.UseVisualStyleBackColor = true;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.numColumnImageScale);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.numTableImageScale);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.numOutputScale);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Location = new System.Drawing.Point(12, 108);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(460, 130);
            this.groupBox2.TabIndex = 1;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "画像スケール設定（0.1?10.0）";
            // 
            // numColumnImageScale
            // 
            this.numColumnImageScale.DecimalPlaces = 2;
            this.numColumnImageScale.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numColumnImageScale.Location = new System.Drawing.Point(230, 90);
            this.numColumnImageScale.Maximum = new decimal(new int[] {
            100,
            0,
            0,
            65536});
            this.numColumnImageScale.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numColumnImageScale.Name = "numColumnImageScale";
            this.numColumnImageScale.Size = new System.Drawing.Size(100, 19);
            this.numColumnImageScale.TabIndex = 5;
            this.numColumnImageScale.Value = new decimal(new int[] {
            12,
            0,
            0,
            65536});
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(15, 92);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(185, 12);
            this.label3.TabIndex = 4;
            this.label3.Text = "コラム内画像スケール倍率:";
            // 
            // numTableImageScale
            // 
            this.numTableImageScale.DecimalPlaces = 2;
            this.numTableImageScale.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numTableImageScale.Location = new System.Drawing.Point(230, 58);
            this.numTableImageScale.Maximum = new decimal(new int[] {
            100,
            0,
            0,
            65536});
            this.numTableImageScale.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numTableImageScale.Name = "numTableImageScale";
            this.numTableImageScale.Size = new System.Drawing.Size(100, 19);
            this.numTableImageScale.TabIndex = 3;
            this.numTableImageScale.Value = new decimal(new int[] {
            12,
            0,
            0,
            65536});
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 60);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(166, 12);
            this.label2.TabIndex = 2;
            this.label2.Text = "表内画像スケール倍率:";
            // 
            // numOutputScale
            // 
            this.numOutputScale.DecimalPlaces = 2;
            this.numOutputScale.Increment = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numOutputScale.Location = new System.Drawing.Point(230, 26);
            this.numOutputScale.Maximum = new decimal(new int[] {
            100,
            0,
            0,
            65536});
            this.numOutputScale.Minimum = new decimal(new int[] {
            1,
            0,
            0,
            65536});
            this.numOutputScale.Name = "numOutputScale";
            this.numOutputScale.Size = new System.Drawing.Size(100, 19);
            this.numOutputScale.TabIndex = 1;
            this.numOutputScale.Value = new decimal(new int[] {
            14,
            0,
            0,
            65536});
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(15, 28);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(175, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "通常画像スケール倍率:";
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.numMaxOutputHeight);
            this.groupBox3.Controls.Add(this.label5);
            this.groupBox3.Controls.Add(this.numMaxOutputWidth);
            this.groupBox3.Controls.Add(this.label4);
            this.groupBox3.Location = new System.Drawing.Point(12, 244);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(460, 90);
            this.groupBox3.TabIndex = 5;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "出力画像サイズ設定（100?4096）";
            // 
            // numMaxOutputHeight
            // 
            this.numMaxOutputHeight.Location = new System.Drawing.Point(230, 55);
            this.numMaxOutputHeight.Maximum = new decimal(new int[] {
            4096,
            0,
            0,
            0});
            this.numMaxOutputHeight.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.numMaxOutputHeight.Name = "numMaxOutputHeight";
            this.numMaxOutputHeight.Size = new System.Drawing.Size(100, 19);
            this.numMaxOutputHeight.TabIndex = 3;
            this.numMaxOutputHeight.Value = new decimal(new int[] {
            1024,
            0,
            0,
            0});
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(15, 57);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(158, 12);
            this.label5.TabIndex = 2;
            this.label5.Text = "出力画像の最大高さ（ピクセル）:";
            // 
            // numMaxOutputWidth
            // 
            this.numMaxOutputWidth.Location = new System.Drawing.Point(230, 23);
            this.numMaxOutputWidth.Maximum = new decimal(new int[] {
            4096,
            0,
            0,
            0});
            this.numMaxOutputWidth.Minimum = new decimal(new int[] {
            100,
            0,
            0,
            0});
            this.numMaxOutputWidth.Name = "numMaxOutputWidth";
            this.numMaxOutputWidth.Size = new System.Drawing.Size(100, 19);
            this.numMaxOutputWidth.TabIndex = 1;
            this.numMaxOutputWidth.Value = new decimal(new int[] {
            1024,
            0,
            0,
            0});
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(15, 25);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(146, 12);
            this.label4.TabIndex = 0;
            this.label4.Text = "出力画像の最大幅（ピクセル）:";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(280, 345);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(90, 30);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(382, 345);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(90, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // btnResetDefaults
            // 
            this.btnResetDefaults.Location = new System.Drawing.Point(12, 345);
            this.btnResetDefaults.Name = "btnResetDefaults";
            this.btnResetDefaults.Size = new System.Drawing.Size(130, 30);
            this.btnResetDefaults.TabIndex = 4;
            this.btnResetDefaults.Text = "デフォルト値に戻す";
            this.btnResetDefaults.UseVisualStyleBackColor = true;
            this.btnResetDefaults.Click += new System.EventHandler(this.btnResetDefaults_Click);
            // 
            // SettingsForm
            // 
            this.AcceptButton = this.btnOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.btnCancel;
            this.ClientSize = new System.Drawing.Size(484, 386);
            this.Controls.Add(this.btnResetDefaults);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.groupBox3);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "SettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "MJS プラグイン設定";
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.numColumnImageScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numTableImageScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numOutputScale)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxOutputWidth)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numMaxOutputHeight)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.CheckBox chkBetaMode;
        private System.Windows.Forms.CheckBox chkExtractHighQualityImages;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.NumericUpDown numColumnImageScale;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.NumericUpDown numTableImageScale;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.NumericUpDown numOutputScale;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.NumericUpDown numMaxOutputHeight;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.NumericUpDown numMaxOutputWidth;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Button btnCancel;
        private System.Windows.Forms.Button btnResetDefaults;
    }
}
