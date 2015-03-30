namespace SeniorProject
{
    partial class RoyalWordUC
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

        #region Component Designer generated code

        /// <summary> 
        /// Required method for Designer support - do not modify 
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.btn_check = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.btn_next = new System.Windows.Forms.Button();
            this.btn_clearHighlightALL = new System.Windows.Forms.Button();
            this.btn_highlightALL = new System.Windows.Forms.Button();
            this.btn_fullStopEdit = new System.Windows.Forms.Button();
            this.lbl_header = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label9 = new System.Windows.Forms.Label();
            this.label10 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.lbl_waitCheck = new System.Windows.Forms.Label();
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.SuspendLayout();
            // 
            // listBox1
            // 
            this.listBox1.CausesValidation = false;
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(3, 98);
            this.listBox1.Name = "listBox1";
            this.listBox1.ScrollAlwaysVisible = true;
            this.listBox1.Size = new System.Drawing.Size(145, 212);
            this.listBox1.TabIndex = 3;
            this.listBox1.SelectedValueChanged += new System.EventHandler(this.listBox1_SelectedValueChanged);
            // 
            // btn_check
            // 
            this.btn_check.Location = new System.Drawing.Point(131, 54);
            this.btn_check.Name = "btn_check";
            this.btn_check.Size = new System.Drawing.Size(73, 35);
            this.btn_check.TabIndex = 7;
            this.btn_check.Text = "ตรวจสอบ";
            this.btn_check.UseVisualStyleBackColor = true;
            this.btn_check.Click += new System.EventHandler(this.button6_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label8.Location = new System.Drawing.Point(244, 198);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(18, 29);
            this.label8.TabIndex = 33;
            this.label8.Text = "-";
            // 
            // btn_next
            // 
            this.btn_next.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_next.Location = new System.Drawing.Point(229, 230);
            this.btn_next.Name = "btn_next";
            this.btn_next.Size = new System.Drawing.Size(53, 34);
            this.btn_next.TabIndex = 34;
            this.btn_next.Text = "Next";
            this.btn_next.UseVisualStyleBackColor = true;
            this.btn_next.Click += new System.EventHandler(this.btn_next_Click);
            // 
            // btn_clearHighlightALL
            // 
            this.btn_clearHighlightALL.Location = new System.Drawing.Point(154, 172);
            this.btn_clearHighlightALL.Name = "btn_clearHighlightALL";
            this.btn_clearHighlightALL.Size = new System.Drawing.Size(131, 23);
            this.btn_clearHighlightALL.TabIndex = 32;
            this.btn_clearHighlightALL.Text = "ลบไฮไลท์ทั้งหมด";
            this.btn_clearHighlightALL.UseVisualStyleBackColor = true;
            this.btn_clearHighlightALL.Click += new System.EventHandler(this.btn_clearHighlightALL_Click);
            // 
            // btn_highlightALL
            // 
            this.btn_highlightALL.Location = new System.Drawing.Point(154, 129);
            this.btn_highlightALL.Name = "btn_highlightALL";
            this.btn_highlightALL.Size = new System.Drawing.Size(131, 37);
            this.btn_highlightALL.TabIndex = 31;
            this.btn_highlightALL.Text = "ไฮไลท์ทั้งหมด";
            this.btn_highlightALL.UseVisualStyleBackColor = true;
            this.btn_highlightALL.Click += new System.EventHandler(this.btn_highlightALL_Click);
            // 
            // btn_fullStopEdit
            // 
            this.btn_fullStopEdit.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.btn_fullStopEdit.Location = new System.Drawing.Point(210, 54);
            this.btn_fullStopEdit.Name = "btn_fullStopEdit";
            this.btn_fullStopEdit.Size = new System.Drawing.Size(80, 35);
            this.btn_fullStopEdit.TabIndex = 30;
            this.btn_fullStopEdit.Text = "แก้ไขทั้งหมด";
            this.btn_fullStopEdit.UseVisualStyleBackColor = true;
            this.btn_fullStopEdit.Click += new System.EventHandler(this.btn_fullStopEdit_Click);
            // 
            // lbl_header
            // 
            this.lbl_header.Font = new System.Drawing.Font("Angsana New", 21.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_header.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lbl_header.Location = new System.Drawing.Point(0, 0);
            this.lbl_header.Name = "lbl_header";
            this.lbl_header.Size = new System.Drawing.Size(300, 40);
            this.lbl_header.TabIndex = 28;
            this.lbl_header.Text = "ตรวจสอบศัพท์บัญญัติ";
            this.lbl_header.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(7, 72);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(18, 13);
            this.label1.TabIndex = 37;
            this.label1.Text = "All";
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(31, 72);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(10, 13);
            this.label9.TabIndex = 38;
            this.label9.Text = "-";
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label10.Location = new System.Drawing.Point(205, 198);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(18, 29);
            this.label10.TabIndex = 39;
            this.label10.Text = "-";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label11.Location = new System.Drawing.Point(154, 198);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(46, 29);
            this.label11.TabIndex = 40;
            this.label11.Text = "Select";
            // 
            // lbl_waitCheck
            // 
            this.lbl_waitCheck.Location = new System.Drawing.Point(7, 335);
            this.lbl_waitCheck.Name = "lbl_waitCheck";
            this.lbl_waitCheck.Size = new System.Drawing.Size(104, 13);
            this.lbl_waitCheck.TabIndex = 42;
            this.lbl_waitCheck.Text = "รอโหลด";
            // 
            // progressBar
            // 
            this.progressBar.Location = new System.Drawing.Point(7, 358);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(275, 32);
            this.progressBar.TabIndex = 41;
            this.progressBar.Tag = "";
            // 
            // RoyalWordUC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lbl_waitCheck);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btn_next);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.btn_clearHighlightALL);
            this.Controls.Add(this.btn_highlightALL);
            this.Controls.Add(this.btn_fullStopEdit);
            this.Controls.Add(this.lbl_header);
            this.Controls.Add(this.btn_check);
            this.Controls.Add(this.listBox1);
            this.MinimumSize = new System.Drawing.Size(300, 0);
            this.Name = "RoyalWordUC";
            this.Size = new System.Drawing.Size(300, 455);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button btn_check;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button btn_next;
        private System.Windows.Forms.Button btn_clearHighlightALL;
        private System.Windows.Forms.Button btn_highlightALL;
        private System.Windows.Forms.Label lbl_header;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label lbl_waitCheck;
        public System.Windows.Forms.ProgressBar progressBar;
        public System.Windows.Forms.Button btn_fullStopEdit;
    }
}
