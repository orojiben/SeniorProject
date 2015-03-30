namespace SeniorProject
{
    partial class ReferenceModelUC
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
            this.lbl_header = new System.Windows.Forms.Label();
            this.lbl_reference = new System.Windows.Forms.Label();
            this.lbl_numberReferenceValue = new System.Windows.Forms.Label();
            this.lbl_numberReference = new System.Windows.Forms.Label();
            this.tbx_referenceAllError = new System.Windows.Forms.TextBox();
            this.btn_back = new System.Windows.Forms.Button();
            this.btn_next = new System.Windows.Forms.Button();
            this.lbl_referenceAllError = new System.Windows.Forms.Label();
            this.lbl_referenceNameAndYear = new System.Windows.Forms.Label();
            this.btn_edit = new System.Windows.Forms.Button();
            this.lbl_margin = new System.Windows.Forms.Label();
            this.lbl_referenceAll = new System.Windows.Forms.Label();
            this.lbl_referenceAllCheck = new System.Windows.Forms.Label();
            this.lbl_referenceCheck = new System.Windows.Forms.Label();
            this.lbl_referenceNameAndYearCheck = new System.Windows.Forms.Label();
            this.lbl_marginCheck = new System.Windows.Forms.Label();
            this.lbl_Error = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lbl_header
            // 
            this.lbl_header.Font = new System.Drawing.Font("Angsana New", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_header.Location = new System.Drawing.Point(3, 0);
            this.lbl_header.Name = "lbl_header";
            this.lbl_header.Size = new System.Drawing.Size(283, 38);
            this.lbl_header.TabIndex = 2;
            this.lbl_header.Text = "ผลการตรวจสอบรูปแบบอ้างอิง";
            this.lbl_header.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // lbl_reference
            // 
            this.lbl_reference.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_reference.Location = new System.Drawing.Point(5, 198);
            this.lbl_reference.Name = "lbl_reference";
            this.lbl_reference.Size = new System.Drawing.Size(106, 25);
            this.lbl_reference.TabIndex = 3;
            this.lbl_reference.Text = "รูปแบบผิด: ";
            // 
            // lbl_numberReferenceValue
            // 
            this.lbl_numberReferenceValue.AutoSize = true;
            this.lbl_numberReferenceValue.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_numberReferenceValue.Location = new System.Drawing.Point(142, 114);
            this.lbl_numberReferenceValue.Name = "lbl_numberReferenceValue";
            this.lbl_numberReferenceValue.Size = new System.Drawing.Size(20, 29);
            this.lbl_numberReferenceValue.TabIndex = 15;
            this.lbl_numberReferenceValue.Text = "0";
            // 
            // lbl_numberReference
            // 
            this.lbl_numberReference.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_numberReference.Location = new System.Drawing.Point(3, 117);
            this.lbl_numberReference.Name = "lbl_numberReference";
            this.lbl_numberReference.Size = new System.Drawing.Size(137, 26);
            this.lbl_numberReference.TabIndex = 14;
            this.lbl_numberReference.Text = "พารากราฟที่เลือก: ";
            // 
            // tbx_referenceAllError
            // 
            this.tbx_referenceAllError.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.tbx_referenceAllError.Font = new System.Drawing.Font("Angsana New", 14.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.tbx_referenceAllError.ForeColor = System.Drawing.Color.Red;
            this.tbx_referenceAllError.Location = new System.Drawing.Point(146, 75);
            this.tbx_referenceAllError.Multiline = true;
            this.tbx_referenceAllError.Name = "tbx_referenceAllError";
            this.tbx_referenceAllError.ReadOnly = true;
            this.tbx_referenceAllError.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.tbx_referenceAllError.Size = new System.Drawing.Size(140, 31);
            this.tbx_referenceAllError.TabIndex = 13;
            this.tbx_referenceAllError.Text = "-";
            // 
            // btn_back
            // 
            this.btn_back.Location = new System.Drawing.Point(141, 153);
            this.btn_back.Name = "btn_back";
            this.btn_back.Size = new System.Drawing.Size(61, 23);
            this.btn_back.TabIndex = 12;
            this.btn_back.Text = "ย้อนกลับ";
            this.btn_back.UseVisualStyleBackColor = true;
            this.btn_back.Click += new System.EventHandler(this.btn_back_Click);
            // 
            // btn_next
            // 
            this.btn_next.Location = new System.Drawing.Point(208, 153);
            this.btn_next.Name = "btn_next";
            this.btn_next.Size = new System.Drawing.Size(56, 23);
            this.btn_next.TabIndex = 11;
            this.btn_next.Text = "ต่อไป";
            this.btn_next.UseVisualStyleBackColor = true;
            this.btn_next.Click += new System.EventHandler(this.btn_next_Click);
            // 
            // lbl_referenceAllError
            // 
            this.lbl_referenceAllError.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceAllError.Location = new System.Drawing.Point(4, 76);
            this.lbl_referenceAllError.Name = "lbl_referenceAllError";
            this.lbl_referenceAllError.Size = new System.Drawing.Size(158, 34);
            this.lbl_referenceAllError.TabIndex = 10;
            this.lbl_referenceAllError.Text = "พารากราฟที่ผิดทั้งหมด: ";
            // 
            // lbl_referenceNameAndYear
            // 
            this.lbl_referenceNameAndYear.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceNameAndYear.Location = new System.Drawing.Point(5, 241);
            this.lbl_referenceNameAndYear.Name = "lbl_referenceNameAndYear";
            this.lbl_referenceNameAndYear.Size = new System.Drawing.Size(157, 31);
            this.lbl_referenceNameAndYear.TabIndex = 19;
            this.lbl_referenceNameAndYear.Text = "ระบบนาม-ปีผิด: ";
            // 
            // btn_edit
            // 
            this.btn_edit.Location = new System.Drawing.Point(10, 487);
            this.btn_edit.Name = "btn_edit";
            this.btn_edit.Size = new System.Drawing.Size(252, 23);
            this.btn_edit.TabIndex = 27;
            this.btn_edit.Text = "แก้ไขเรียงปีกับเรียงตัวษรที่ผิด";
            this.btn_edit.UseVisualStyleBackColor = true;
            // 
            // lbl_margin
            // 
            this.lbl_margin.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_margin.Location = new System.Drawing.Point(5, 280);
            this.lbl_margin.Name = "lbl_margin";
            this.lbl_margin.Size = new System.Drawing.Size(103, 31);
            this.lbl_margin.TabIndex = 28;
            this.lbl_margin.Text = "ขอบกระดาษผิด: ";
            // 
            // lbl_referenceAll
            // 
            this.lbl_referenceAll.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceAll.Location = new System.Drawing.Point(5, 36);
            this.lbl_referenceAll.Name = "lbl_referenceAll";
            this.lbl_referenceAll.Size = new System.Drawing.Size(137, 34);
            this.lbl_referenceAll.TabIndex = 29;
            this.lbl_referenceAll.Text = "พารากราฟทั้งหมด: ";
            // 
            // lbl_referenceAllCheck
            // 
            this.lbl_referenceAllCheck.AutoSize = true;
            this.lbl_referenceAllCheck.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceAllCheck.Location = new System.Drawing.Point(142, 36);
            this.lbl_referenceAllCheck.Name = "lbl_referenceAllCheck";
            this.lbl_referenceAllCheck.Size = new System.Drawing.Size(20, 29);
            this.lbl_referenceAllCheck.TabIndex = 30;
            this.lbl_referenceAllCheck.Text = "0";
            // 
            // lbl_referenceCheck
            // 
            this.lbl_referenceCheck.AutoSize = true;
            this.lbl_referenceCheck.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceCheck.ForeColor = System.Drawing.Color.Green;
            this.lbl_referenceCheck.Location = new System.Drawing.Point(152, 198);
            this.lbl_referenceCheck.Name = "lbl_referenceCheck";
            this.lbl_referenceCheck.Size = new System.Drawing.Size(27, 29);
            this.lbl_referenceCheck.TabIndex = 31;
            this.lbl_referenceCheck.Text = "✔";
            // 
            // lbl_referenceNameAndYearCheck
            // 
            this.lbl_referenceNameAndYearCheck.AutoSize = true;
            this.lbl_referenceNameAndYearCheck.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_referenceNameAndYearCheck.ForeColor = System.Drawing.Color.Green;
            this.lbl_referenceNameAndYearCheck.Location = new System.Drawing.Point(152, 243);
            this.lbl_referenceNameAndYearCheck.Name = "lbl_referenceNameAndYearCheck";
            this.lbl_referenceNameAndYearCheck.Size = new System.Drawing.Size(27, 29);
            this.lbl_referenceNameAndYearCheck.TabIndex = 32;
            this.lbl_referenceNameAndYearCheck.Text = "✔";
            // 
            // lbl_marginCheck
            // 
            this.lbl_marginCheck.AutoSize = true;
            this.lbl_marginCheck.Font = new System.Drawing.Font("Angsana New", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lbl_marginCheck.ForeColor = System.Drawing.Color.Green;
            this.lbl_marginCheck.Location = new System.Drawing.Point(152, 282);
            this.lbl_marginCheck.Name = "lbl_marginCheck";
            this.lbl_marginCheck.Size = new System.Drawing.Size(27, 29);
            this.lbl_marginCheck.TabIndex = 33;
            this.lbl_marginCheck.Text = "✔";
            // 
            // lbl_Error
            // 
            this.lbl_Error.Font = new System.Drawing.Font("Angsana New", 28F, System.Drawing.FontStyle.Bold);
            this.lbl_Error.ForeColor = System.Drawing.Color.Red;
            this.lbl_Error.Location = new System.Drawing.Point(0, 110);
            this.lbl_Error.Name = "lbl_Error";
            this.lbl_Error.Size = new System.Drawing.Size(286, 133);
            this.lbl_Error.TabIndex = 34;
            this.lbl_Error.Text = "ไม่พบข้อมูลอ้างอิงหรือไม่ตรงตามข้อกำหนด";
            this.lbl_Error.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // ReferenceModelUC
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lbl_marginCheck);
            this.Controls.Add(this.lbl_referenceNameAndYearCheck);
            this.Controls.Add(this.lbl_referenceCheck);
            this.Controls.Add(this.lbl_referenceAllCheck);
            this.Controls.Add(this.lbl_referenceAll);
            this.Controls.Add(this.lbl_margin);
            this.Controls.Add(this.btn_edit);
            this.Controls.Add(this.lbl_referenceNameAndYear);
            this.Controls.Add(this.lbl_numberReferenceValue);
            this.Controls.Add(this.lbl_numberReference);
            this.Controls.Add(this.tbx_referenceAllError);
            this.Controls.Add(this.btn_back);
            this.Controls.Add(this.btn_next);
            this.Controls.Add(this.lbl_referenceAllError);
            this.Controls.Add(this.lbl_reference);
            this.Controls.Add(this.lbl_header);
            this.Controls.Add(this.lbl_Error);
            this.Name = "ReferenceModelUC";
            this.Size = new System.Drawing.Size(289, 528);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label lbl_header;
        private System.Windows.Forms.Label lbl_reference;
        private System.Windows.Forms.Label lbl_numberReferenceValue;
        private System.Windows.Forms.Label lbl_numberReference;
        private System.Windows.Forms.TextBox tbx_referenceAllError;
        private System.Windows.Forms.Label lbl_referenceAllError;
        private System.Windows.Forms.Label lbl_referenceNameAndYear;
        public System.Windows.Forms.Button btn_back;
        public System.Windows.Forms.Button btn_next;
        public System.Windows.Forms.Button btn_edit;
        private System.Windows.Forms.Label lbl_margin;
        private System.Windows.Forms.Label lbl_referenceAll;
        private System.Windows.Forms.Label lbl_referenceCheck;
        private System.Windows.Forms.Label lbl_referenceNameAndYearCheck;
        private System.Windows.Forms.Label lbl_marginCheck;
        public System.Windows.Forms.Label lbl_referenceAllCheck;
        private System.Windows.Forms.Label lbl_Error;
    }
}
