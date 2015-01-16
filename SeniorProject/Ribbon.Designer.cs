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
            this.group1 = this.Factory.CreateRibbonGroup();
            this.dropDown1 = this.Factory.CreateRibbonDropDown();
            this.button4 = this.Factory.CreateRibbonButton();
            this.button5 = this.Factory.CreateRibbonButton();
            this.button6 = this.Factory.CreateRibbonButton();
            this.btn_SaveNewFile = this.Factory.CreateRibbonButton();
            this.button2 = this.Factory.CreateRibbonButton();
            this.button3 = this.Factory.CreateRibbonButton();
            this.group2 = this.Factory.CreateRibbonGroup();
            this.btn_checkRoyalWord = this.Factory.CreateRibbonButton();
            this.btn_checkSing = this.Factory.CreateRibbonButton();
            this.group3 = this.Factory.CreateRibbonGroup();
            this.btn_checkFont = this.Factory.CreateRibbonButton();
            this.button1 = this.Factory.CreateRibbonButton();
            this.btn_correctFont = this.Factory.CreateRibbonButton();
            this.group4 = this.Factory.CreateRibbonGroup();
            this.btn_checkAll = this.Factory.CreateRibbonButton();
            this.tab1.SuspendLayout();
            this.group1.SuspendLayout();
            this.group2.SuspendLayout();
            this.group3.SuspendLayout();
            this.group4.SuspendLayout();
            // 
            // tab1
            // 
            this.tab1.ControlId.ControlIdType = Microsoft.Office.Tools.Ribbon.RibbonControlIdType.Office;
            this.tab1.Groups.Add(this.group1);
            this.tab1.Groups.Add(this.group2);
            this.tab1.Groups.Add(this.group3);
            this.tab1.Groups.Add(this.group4);
            this.tab1.Label = "TabAddIns";
            this.tab1.Name = "tab1";
            // 
            // group1
            // 
            this.group1.Items.Add(this.dropDown1);
            this.group1.Items.Add(this.button4);
            this.group1.Items.Add(this.button5);
            this.group1.Items.Add(this.button6);
            this.group1.Items.Add(this.btn_SaveNewFile);
            this.group1.Items.Add(this.button2);
            this.group1.Items.Add(this.button3);
            this.group1.Label = "group1";
            this.group1.Name = "group1";
            // 
            // dropDown1
            // 
            this.dropDown1.Label = "รูปแบบปริญญานิพนธ์";
            this.dropDown1.Name = "dropDown1";
            this.dropDown1.SelectionChanged += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.dropDown1_SelectionChanged);
            // 
            // button4
            // 
            this.button4.Label = "ตรวจสอบกระดาษที่ใช้";
            this.button4.Name = "button4";
            // 
            // button5
            // 
            this.button5.Label = "ตรวจสอบระยะขอบกระดาษ";
            this.button5.Name = "button5";
            // 
            // button6
            // 
            this.button6.Label = "ตรวจสอบเครื่องหมายวรรคตอน";
            this.button6.Name = "button6";
            this.button6.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button6_Click);
            // 
            // btn_SaveNewFile
            // 
            this.btn_SaveNewFile.Label = "save File New";
            this.btn_SaveNewFile.Name = "btn_SaveNewFile";
            this.btn_SaveNewFile.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Label = "ตรวจสอบรูปแบบอ้างอิง";
            this.button2.Name = "button2";
            this.button2.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Label = "button3";
            this.button3.Name = "button3";
            this.button3.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button3_Click);
            // 
            // group2
            // 
            this.group2.Items.Add(this.btn_checkRoyalWord);
            this.group2.Items.Add(this.btn_checkSing);
            this.group2.Label = "group2";
            this.group2.Name = "group2";
            // 
            // btn_checkRoyalWord
            // 
            this.btn_checkRoyalWord.Label = "ตรวจสอบคำตามศัพท์บรรญัติ";
            this.btn_checkRoyalWord.Name = "btn_checkRoyalWord";
            this.btn_checkRoyalWord.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkRoyalWord_Click);
            // 
            // btn_checkSing
            // 
            this.btn_checkSing.Label = "Check Sing";
            this.btn_checkSing.Name = "btn_checkSing";
            this.btn_checkSing.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkSing_Click);
            // 
            // group3
            // 
            this.group3.Items.Add(this.btn_checkFont);
            this.group3.Items.Add(this.button1);
            this.group3.Items.Add(this.btn_correctFont);
            this.group3.Label = "group3";
            this.group3.Name = "group3";
            // 
            // btn_checkFont
            // 
            this.btn_checkFont.Label = "ตรวจสอบชนิด Font";
            this.btn_checkFont.Name = "btn_checkFont";
            this.btn_checkFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_checkFont_Click);
            // 
            // button1
            // 
            this.button1.Label = "ตรวจสอบขนาด Font";
            this.button1.Name = "button1";
            this.button1.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.button1_Click_1);
            // 
            // btn_correctFont
            // 
            this.btn_correctFont.Label = "แก้ไข Font";
            this.btn_correctFont.Name = "btn_correctFont";
            this.btn_correctFont.Click += new Microsoft.Office.Tools.Ribbon.RibbonControlEventHandler(this.btn_correctFont_Click);
            // 
            // group4
            // 
            this.group4.Items.Add(this.btn_checkAll);
            this.group4.Label = "group4";
            this.group4.Name = "group4";
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
            this.group1.ResumeLayout(false);
            this.group1.PerformLayout();
            this.group2.ResumeLayout(false);
            this.group2.PerformLayout();
            this.group3.ResumeLayout(false);
            this.group3.PerformLayout();
            this.group4.ResumeLayout(false);
            this.group4.PerformLayout();

        }

        #endregion

        internal Microsoft.Office.Tools.Ribbon.RibbonTab tab1;
        internal Microsoft.Office.Tools.Ribbon.RibbonDropDown dropDown1;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_SaveNewFile;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button2;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group2;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkRoyalWord;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkSing;
        private Microsoft.Office.Tools.Ribbon.RibbonGroup group1;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group3;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_correctFont;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button3;
        internal Microsoft.Office.Tools.Ribbon.RibbonGroup group4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton btn_checkAll;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button4;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button5;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button6;
        internal Microsoft.Office.Tools.Ribbon.RibbonButton button1;
    }

    partial class ThisRibbonCollection
    {
        internal Ribbon1 Ribbon1
        {
            get { return this.GetRibbon<Ribbon1>(); }
        }
    }
}

