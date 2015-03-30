namespace SeniorProject
{
    partial class Ribbon1 : Microsoft.Office.Tools.Ribbon.RibbonBase
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        public Ribbon1()
            : base(Globals.Factory.GetRibbonFactory())
        {
            InitializeComponent();
        }

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
            this.tab1 = this.Factory.CreateRibbonTab();
            this.grp_model = this.Factory.CreateRibbonGroup();
            this.ddn_Model = this.Factory.CreateRibbonDropDown();
            this.ddn_Department = this.Factory.CreateRibbonDropDown();
            this.btn_SaveNewFile = this.Factory.CreateRibbonButton();
            this.grp_checking = this.Factory.CreateRibbonGroup();
            this.btn_checkMargin = this.Factory.CreateRibbonButton();
            this.btn_checkPaper = this.Factory.CreateRibbonButton();
            this.btn_checkFontSize = this.Factory.CreateRibbonButton();
            this.btn_checkRoyalWord = this.Factory.CreateRibbonButton();
            this.btn_checkPunctuationMark = this.Factory.CreateRibbonButton();
            this.btn_checkReference = this.Factory.CreateRibbonButton();
            this.btn_checkAll = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.grp_model.SuspendLayout();
            this.grp_checking.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.grp_model);
            this.tab1.Groups.Add(this.grp_checking);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // grp_model
            // 
            this.grp_model.Items.Add(this.ddn_Model);
            this.grp_model.Items.Add(this.ddn_Department);
            this.grp_model.Items.Add(this.btn_SaveNewFile);
            this.grp_model.Label = "รูปแบบ";
            this.grp_model.Name = "grp_model";
            // 
            // ddn_Model
            // 
            this.ddn_Model.Label = "รูปแบบปริญญานิพนธ์";
            this.ddn_Model.Name = "ddn_Model";
            this.ddn_Model.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // ddn_Department
            // 
            this.ddn_Department.Label = "ภาควิชา";
            this.ddn_Department.Name = "ddn_Department";
            this.ddn_Department.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.ddn_Department_SelectionChanged);
            // 
            // btn_SaveNewFile
            // 
            this.btn_SaveNewFile.Label = "เซฟไฟล์ Backup";
            this.btn_SaveNewFile.Name = "btn_SaveNewFile";
            this.btn_SaveNewFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_SaveNewFile_Click);
            // 
            // grp_checking
            // 
            this.grp_checking.Items.Add(this.btn_checkMargin);
            this.grp_checking.Items.Add(this.btn_checkPaper);
            this.grp_checking.Items.Add(this.btn_checkFontSize);
            this.grp_checking.Items.Add(this.btn_checkRoyalWord);
            this.grp_checking.Items.Add(this.btn_checkPunctuationMark);
            this.grp_checking.Items.Add(this.btn_checkReference);
            this.grp_checking.Items.Add(this.btn_checkAll);
            this.grp_checking.Label = "การตรวจสอบ";
            this.grp_checking.Name = "grp_checking";
            // 
            // btn_checkMargin
            // 
            this.btn_checkMargin.Label = "ตรวจสอบระยะขอบกระดาษ";
            this.btn_checkMargin.Name = "btn_checkMargin";
            this.btn_checkMargin.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkMargin_Click);
            // 
            // btn_checkPaper
            // 
            this.btn_checkPaper.Label = "ตรวจสอบชนิดกระดาษ";
            this.btn_checkPaper.Name = "btn_checkPaper";
            this.btn_checkPaper.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkPaper_Click);
            // 
            // btn_checkFontSize
            // 
            this.btn_checkFontSize.Label = "ตรวจสอบขนาดกับชนิดตัวอักษร";
            this.btn_checkFontSize.Name = "btn_checkFontSize";
            this.btn_checkFontSize.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkFontSize_Click);
            // 
            // btn_checkRoyalWord
            // 
            this.btn_checkRoyalWord.Label = "ตรวจสอบคำตามศัพท์บรรญัติ";
            this.btn_checkRoyalWord.Name = "btn_checkRoyalWord";
            this.btn_checkRoyalWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkRoyalWord_Click);
            // 
            // btn_checkPunctuationMark
            // 
            this.btn_checkPunctuationMark.Label = "ตรวจสอบเครื่องหมายวรรคตอน";
            this.btn_checkPunctuationMark.Name = "btn_checkPunctuationMark";
            this.btn_checkPunctuationMark.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkPunctuationMark_Click);
            // 
            // btn_checkReference
            // 
            this.btn_checkReference.Label = "ตรวจสอบรูปแบบอ้างอิง";
            this.btn_checkReference.Name = "btn_checkReference";
            this.btn_checkReference.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkReference_Click);
            // 
            // btn_checkAll
            // 
            this.btn_checkAll.Label = "ตรวจทั้งหมด";
            this.btn_checkAll.Name = "btn_checkAll";
            this.btn_checkAll.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkAll_Click);
            // 
            // Ribbon1
            // 
            this.Name = "Ribbon1";
            this.RibbonType = "Microsoft.Word.Document";
            this.Tabs.Add(this.tab1);
            this.Load += new Microsoft.Office.Tools.Ribbon.RibbonUIEventHandler(this.Ribbon1_Load);
            this.tab1.ResumeLayout(false);
            this.tab1.PerformLayout();
            this.grp_model.ResumeLayout(false);
            this.grp_model.PerformLayout();
            this.grp_checking.ResumeLayout(false);
            this.grp_checking.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddn_Model;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SaveNewFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkReference;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup grp_checking;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkRoyalWord;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup grp_model;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkPaper;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkMargin;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkPunctuationMark;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkFontSize;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown ddn_Department;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

