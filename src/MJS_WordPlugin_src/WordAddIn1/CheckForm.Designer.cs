namespace WordAddIn1
{
    partial class CheckForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.old_num = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.old_title = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.old_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.new_num = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.new_title = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.new_id = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Column1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.diff = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.editnew = new System.Windows.Forms.DataGridViewLinkColumn();
            this.editshow = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.AllowUserToDeleteRows = false;
            this.dataGridView1.AllowUserToResizeRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.old_num,
            this.old_title,
            this.old_id,
            this.new_num,
            this.new_title,
            this.new_id,
            this.Column1,
            this.diff,
            this.editnew,
            this.editshow});
            this.dataGridView1.Location = new System.Drawing.Point(16, 16);
            this.dataGridView1.Margin = new System.Windows.Forms.Padding(4);
            this.dataGridView1.MultiSelect = false;
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.RowHeadersVisible = false;
            this.dataGridView1.RowTemplate.Height = 21;
            this.dataGridView1.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.dataGridView1.Size = new System.Drawing.Size(1058, 529);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.DataGridView1_CellClick);
            this.dataGridView1.CellValueChanged += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView1_CellValueChanged);
            this.dataGridView1.DataBindingComplete += new System.Windows.Forms.DataGridViewBindingCompleteEventHandler(this.dataGridView1_DataBindingComplete);
            // 
            // old_num
            // 
            this.old_num.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.old_num.DataPropertyName = "old_num";
            this.old_num.FillWeight = 60F;
            this.old_num.HeaderText = "項番(旧)";
            this.old_num.Name = "old_num";
            this.old_num.ReadOnly = true;
            this.old_num.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // old_title
            // 
            this.old_title.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.old_title.DataPropertyName = "old_title";
            this.old_title.FillWeight = 260F;
            this.old_title.HeaderText = "タイトル(旧)";
            this.old_title.Name = "old_title";
            this.old_title.ReadOnly = true;
            this.old_title.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // old_id
            // 
            this.old_id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.old_id.DataPropertyName = "old_id";
            this.old_id.FillWeight = 120F;
            this.old_id.HeaderText = "ID(旧)";
            this.old_id.Name = "old_id";
            this.old_id.ReadOnly = true;
            this.old_id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // new_num
            // 
            this.new_num.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.new_num.DataPropertyName = "new_num";
            this.new_num.FillWeight = 60F;
            this.new_num.HeaderText = "項番(新)";
            this.new_num.Name = "new_num";
            this.new_num.ReadOnly = true;
            this.new_num.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // new_title
            // 
            this.new_title.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.new_title.DataPropertyName = "new_title";
            this.new_title.FillWeight = 260F;
            this.new_title.HeaderText = "タイトル(新)";
            this.new_title.Name = "new_title";
            this.new_title.ReadOnly = true;
            this.new_title.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // new_id
            // 
            this.new_id.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.new_id.DataPropertyName = "new_id";
            this.new_id.FillWeight = 120F;
            this.new_id.HeaderText = "ID(新)";
            this.new_id.Name = "new_id";
            this.new_id.ReadOnly = true;
            this.new_id.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // Column1
            // 
            this.Column1.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.Column1.DataPropertyName = "new_id_show";
            this.Column1.FillWeight = 120F;
            this.Column1.HeaderText = "新.ID（修正候補）";
            this.Column1.Name = "Column1";
            this.Column1.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // diff
            // 
            this.diff.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.diff.DataPropertyName = "diff";
            this.diff.FillWeight = 120F;
            this.diff.HeaderText = "差異内容";
            this.diff.Name = "diff";
            this.diff.ReadOnly = true;
            this.diff.Resizable = System.Windows.Forms.DataGridViewTriState.True;
            this.diff.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
            // 
            // editnew
            // 
            this.editnew.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.editnew.DataPropertyName = "edit";
            dataGridViewCellStyle1.Font = new System.Drawing.Font("MS UI Gothic", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(128)));
            this.editnew.DefaultCellStyle = dataGridViewCellStyle1;
            this.editnew.FillWeight = 80F;
            this.editnew.HeaderText = "新規追加";
            this.editnew.Name = "editnew";
            this.editnew.ReadOnly = true;
            // 
            // editshow
            // 
            this.editshow.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.editshow.DataPropertyName = "editshow";
            this.editshow.HeaderText = "修正処理（候補）";
            this.editshow.Name = "editshow";
            this.editshow.ReadOnly = true;
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Location = new System.Drawing.Point(16, 553);
            this.button1.Margin = new System.Windows.Forms.Padding(4);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(133, 31);
            this.button1.TabIndex = 1;
            this.button1.Text = "CSV出力";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.ExportCsvButton_Click);
            // 
            // button2
            // 
            this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button2.Location = new System.Drawing.Point(800, 553);
            this.button2.Margin = new System.Windows.Forms.Padding(4);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(133, 31);
            this.button2.TabIndex = 2;
            this.button2.Text = "更新";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.UpdateButton_Click);
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Location = new System.Drawing.Point(941, 553);
            this.button3.Margin = new System.Windows.Forms.Padding(4);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(133, 31);
            this.button3.TabIndex = 3;
            this.button3.Text = "キャンセル";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.CancelButton_Click); // Rename event handler
            // 
            // CheckForm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(1087, 596);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.dataGridView1);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "CheckForm";
            this.Text = "書誌情報比較";
            this.TopMost = true;
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private System.Windows.Forms.DataGridViewTextBoxColumn old_num;
        private System.Windows.Forms.DataGridViewTextBoxColumn old_title;
        private System.Windows.Forms.DataGridViewTextBoxColumn old_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn new_num;
        private System.Windows.Forms.DataGridViewTextBoxColumn new_title;
        private System.Windows.Forms.DataGridViewTextBoxColumn new_id;
        private System.Windows.Forms.DataGridViewTextBoxColumn Column1;
        private System.Windows.Forms.DataGridViewTextBoxColumn diff;
        private System.Windows.Forms.DataGridViewLinkColumn editnew;
        private System.Windows.Forms.DataGridViewTextBoxColumn editshow;
    }
}