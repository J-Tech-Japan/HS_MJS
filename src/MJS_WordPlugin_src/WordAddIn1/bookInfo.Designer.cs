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
            this.label1.Location = new System.Drawing.Point(31, 24);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(115, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "初期値(数字2桁)";
            // 
            // tbxDefaultValue
            // 
            this.tbxDefaultValue.ImeMode = System.Windows.Forms.ImeMode.Disable;
            this.tbxDefaultValue.Location = new System.Drawing.Point(160, 20);
            this.tbxDefaultValue.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.tbxDefaultValue.MaxLength = 2;
            this.tbxDefaultValue.Name = "tbxDefaultValue";
            this.tbxDefaultValue.Size = new System.Drawing.Size(132, 22);
            this.tbxDefaultValue.TabIndex = 1;
            // this.tbxDefaultValue.TextChanged += new System.EventHandler(this.tbxDefaultValue_TextChanged);
            this.tbxDefaultValue.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.tbxDefaultValue_KeyPress);
            // 
            // btnEnter
            // 
            this.btnEnter.Location = new System.Drawing.Point(301, 16);
            this.btnEnter.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnEnter.Name = "btnEnter";
            this.btnEnter.Size = new System.Drawing.Size(88, 30);
            this.btnEnter.TabIndex = 2;
            this.btnEnter.Text = "実行";
            this.btnEnter.UseVisualStyleBackColor = true;
            this.btnEnter.Click += new System.EventHandler(this.btnEnter_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(397, 16);
            this.btnCancel.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(88, 30);
            this.btnCancel.TabIndex = 3;
            this.btnCancel.Text = "キャンセル";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // bookInfo
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 74);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.btnEnter);
            this.Controls.Add(this.tbxDefaultValue);
            this.Controls.Add(this.label1);
            this.Margin = new System.Windows.Forms.Padding(4, 4, 4, 4);
            this.Name = "bookInfo";
            this.Text = "書誌情報出力";
            //this.Load += new System.EventHandler(this.bookInfo_Load_1);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        //private void bookInfo_Load_1(object sender, EventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        //private void tbxDefaultValue_TextChanged(object sender, EventArgs e)
        //{
        //    throw new NotImplementedException();
        //}

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnEnter;
        private System.Windows.Forms.Button btnCancel;
        public System.Windows.Forms.TextBox tbxDefaultValue;
    }
}