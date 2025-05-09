using System;

namespace WordAddIn1
{
    partial class bookInfo
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
            this.label1 = new System.Windows.Forms.Label();
            this.tbxDefaultValue = new System.Windows.Forms.TextBox();
            this.btnEnter = new System.Windows.Forms.Button();
            this.btnCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 29);
            this.label1.Margin = new System.Windows.Forms.Padding(5, 0, 5, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(135, 18);
            this.label1.TabIndex = 0;
            this.label1.Text = "初期値(数字2桁)";
            // 
            // tbxDefaultValue
            // 
            this.tbxDefaultValue.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.tbxDefaultValue.Location = new System.Drawing.Point(200, 24);
            this.tbxDefaultValue.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.tbxDefaultValue.MaxLength = 2;
            this.tbxDefaultValue.Name = "tbxDefaultValue";
            this.tbxDefaultValue.Size = new System.Drawing.Size(164, 25);
            this.tbxDefaultValue.TabIndex = 1;
            this.tbxDefaultValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbxDefaultValue_KeyPress);
            // 
            // btnEnter
            // 
            this.btnEnter.Location = new System.Drawing.Point(376, 19);
            this.btnEnter.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.btnEnter.Name = "btnEnter";
            this.btnEnter.Size = new System.Drawing.Size(110, 36);
            this.btnEnter.TabIndex = 2;
            this.btnEnter.Text = "実行";
            this.btnEnter.UseVisualStyleBackColor = true;
            this.btnEnter.Click += new System.EventHandler(this.btnEnter_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(496, 19);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(110, 36);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // bookInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 18F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(660, 89);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnEnter);
            this.Controls.Add(this.tbxDefaultValue);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(5, 5, 5, 5);
            this.Name = "bookInfo";
            this.Text = "書誌情報出力";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnEnter;
        private System.Windows.Forms.Button btnCancel;
        public System.Windows.Forms.TextBox tbxDefaultValue;
    }
}