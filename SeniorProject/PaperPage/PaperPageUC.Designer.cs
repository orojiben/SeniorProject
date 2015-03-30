namespace SeniorProject
{
    partial class PaperPageUC
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
            this.btn_Edit = new System.Windows.Forms.Button();
            this.lbl_header = new System.Windows.Forms.Label();
            this.lbl_paperType = new System.Windows.Forms.Label();
            this.lbl_paperTypeValue = new System.Windows.Forms.Label();
            this.lbl_paperTypeCheck = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btn_Edit
            // 
            this.btn_Edit.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btn_Edit.Location = new System.Drawing.Point(196, 165);
            this.btn_Edit.Name = "btn_Edit";
            this.btn_Edit.Size = new System.Drawing.Size(75, 33);
            this.btn_Edit.TabIndex = 0;
            this.btn_Edit.Text = "แก้ไข";
            this.btn_Edit.UseVisualStyleBackColor = true;
            // 
            // lbl_header
            // 
            this.lbl_header.Font = new System.Drawing.Font("Angsana New", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_header.Location = new System.Drawing.Point(3, 0);
            this.lbl_header.Name = "lbl_header";
            this.lbl_header.Size = new System.Drawing.Size(283, 38);
            this.lbl_header.TabIndex = 1;
            this.lbl_header.Text = "ตรวจสอบชนิดกระดาษ";
            this.lbl_header.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_paperType
            // 
            this.lbl_paperType.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lbl_paperType.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_paperType.Location = new System.Drawing.Point(9, 54);
            this.lbl_paperType.Name = "lbl_paperType";
            this.lbl_paperType.Size = new System.Drawing.Size(105, 144);
            this.lbl_paperType.TabIndex = 2;
            this.lbl_paperType.Text = "-";
            this.lbl_paperType.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // lbl_paperTypeValue
            // 
            this.lbl_paperTypeValue.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_paperTypeValue.Location = new System.Drawing.Point(140, 54);
            this.lbl_paperTypeValue.Name = "lbl_paperTypeValue";
            this.lbl_paperTypeValue.Size = new System.Drawing.Size(78, 29);
            this.lbl_paperTypeValue.TabIndex = 3;
            this.lbl_paperTypeValue.Text = "-";
            // 
            // lbl_paperTypeCheck
            // 
            this.lbl_paperTypeCheck.AutoSize = true;
            this.lbl_paperTypeCheck.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_paperTypeCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_paperTypeCheck.Location = new System.Drawing.Point(224, 54);
            this.lbl_paperTypeCheck.Name = "lbl_paperTypeCheck";
            this.lbl_paperTypeCheck.Size = new System.Drawing.Size(28, 29);
            this.lbl_paperTypeCheck.TabIndex = 10;
            this.lbl_paperTypeCheck.Text = "✘";
            // 
            // PaperPageUC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lbl_paperTypeCheck);
            this.Controls.Add(this.lbl_paperTypeValue);
            this.Controls.Add(this.lbl_paperType);
            this.Controls.Add(this.lbl_header);
            this.Controls.Add(this.btn_Edit);
            this.Name = "PaperPageUC";
            this.Size = new System.Drawing.Size(289, 212);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_header;
        private System.Windows.Forms.Label lbl_paperType;
        private System.Windows.Forms.Label lbl_paperTypeValue;
        private System.Windows.Forms.Label lbl_paperTypeCheck;
        public System.Windows.Forms.Button btn_Edit;
    }
}
