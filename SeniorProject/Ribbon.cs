using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;

using System.IO;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    public partial class Ribbon1
    {
        //Microsoft.Office.Tools.Word;
        static public MarginPage marginPage;
        static public PaperPage paperPage;
        static public RoyalWord royalWord;
        static public Punctuation punctuation;
        static public ReferenceModel referenceModel;
        
        List<Styles> loadStyles;


        static public string namefileSaveAuto = "";

        public static Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;

        static public ShowCheckAllUC showCheckAllUC;
        static public MarginPageUC marginPageUC;
        static public PaperPageUC paperPageUC;
        static public RoyalWordUC royalWordUC;
        static public PunctuationUC punctuationUC;
        static public ReferenceModelUC referenceModelUC;
        static public FontUC fontUC;
        static public Styles styles;
        static public string nameFile = "";
        int i = 0;
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            nameFile =  Globals.ThisAddIn.Application.ActiveDocument.Name;
            styles = new Styles();
            //Word.Application a = this.r
            Ribbon1.marginPage = new MarginPage();
            Ribbon1.paperPage = new PaperPage();
            Ribbon1.royalWord = new RoyalWord();
            Ribbon1.punctuation = new Punctuation();
            Ribbon1.referenceModel = new ReferenceModel();
            loadStyles = StyleFile.LoadStyle();
           
            readFileStyleToList();
           
            
        }



        private void readFileStyleToList()
        {
            try
            {
                this.ddn_Model.Items.Clear();
                foreach (Styles style in this.loadStyles)
                {
                    Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItemImpl1.Label = style.Name;
                    this.ddn_Model.Items.Add(ribbonDropDownItemImpl1);
                }
                loadDataStyles(0);
            }
            catch { };

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.showCustomTaskPane();
            loadDataStyles(this.ddn_Model.SelectedItemIndex);
        }





        private void btn_checkRoyalWord_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            royalWord.show();
        }

        private void btn_correctFont_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void btn_checkAll_Click(object sender, RibbonControlEventArgs e)
        {

            saveFileAuto();

            addCustomTaskPaneALL();
            
            Ribbon1.showCustomTaskPane(0);
            Ribbon1.showCheckAllUC.visibleAllClose();
            Ribbon1.showCheckAllUC.resetAll();
            //System.Windows.Forms.MessageBox.Show(s, "");
            
            

        }

        private void loadDataStyles(int index)
        {
            Ribbon1.styles = this.loadStyles[index];
            this.ddn_Department.Items.Clear();
            this.ddn_Department.Visible = false;
            Ribbon1.referenceModel.faculty = "";
            Ribbon1.referenceModel.department = "";
            if (styles.Departments.Count > 0)
            {
                this.ddn_Department.Visible = true;
                foreach (string departments in styles.Departments)
                {
                    Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItemImpl1.Label = departments;
                    this.ddn_Department.Items.Add(ribbonDropDownItemImpl1);
                }
                Ribbon1.referenceModel.department = styles.Departments[0];
            }
            Ribbon1.referenceModel.faculty = styles.Name;
            Ribbon1.namefileSaveAuto = Ribbon1.referenceModel.faculty + "_" + Ribbon1.referenceModel.department;
        }

        private void ddn_Department_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.showCustomTaskPane();
            Ribbon1.namefileSaveAuto = Ribbon1.namefileSaveAuto.Substring(0, Ribbon1.referenceModel.faculty.Length);
            Ribbon1.referenceModel.department = this.ddn_Department.Items[this.ddn_Department.SelectedItemIndex].Label;
            Ribbon1.namefileSaveAuto += Ribbon1.referenceModel.department;
        }

        private void btn_editReference_Click(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.referenceModel.runEditReferenceAll();
        }

        private void btn_checkPunctuationMark_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            punctuation.show();
        }

        private void btn_checkMargin_Click(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.marginPage.cheking();//)
            //  {
            //    System.Windows.Forms.MessageBox.Show("ไม่ตรง");
            //}
        }

        private void btn_editMargin_Click(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.marginPage.changing();
        }

        private void btn_editPaper_Click(object sender, RibbonControlEventArgs e)
        {
            Ribbon1.paperPage.changing();
        }

        private void btn_checkPaper_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.paperPage.cheking();
            //{
            //System.Windows.Forms.MessageBox.Show("ไม่ตรง");
            //}
        }

        private void btn_checkFont_Click(object sender, RibbonControlEventArgs e)
        {
            saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.fontUC.enableAll();
            Ribbon1.showCustomTaskPane(4);
        }

        public void show()
        {

            if (Globals.ThisAddIn.CustomTaskPanes.Count > 0)
            {
                for (int i = 0; i < Globals.ThisAddIn.CustomTaskPanes.Count; ++i)
                {
                    Globals.ThisAddIn.CustomTaskPanes.RemoveAt(i);
                }
            }

            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(showCheckAllUC, "Check All");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;
        }

        static public void addCustomTaskPaneALL()
        {
            if (Globals.ThisAddIn.CustomTaskPanes.Count != 0)
            {
                return;
            }
            showCheckAllUC = new ShowCheckAllUC(marginPage, paperPage, referenceModel, punctuation, royalWord);
            Ribbon1.marginPageUC = new MarginPageUC();
            Ribbon1.paperPageUC = new PaperPageUC();
            Ribbon1.royalWordUC = new RoyalWordUC();
            Ribbon1.punctuationUC = new PunctuationUC();
            Ribbon1.referenceModelUC = new ReferenceModelUC();
            Ribbon1.fontUC = new FontUC();
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.showCheckAllUC, "ตรวจสอบทั้งหมด");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.marginPageUC, "ตรวจสอบระยะขอบกระดาษ");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.paperPageUC, "ตรวจสอบชนิดกระดาษ");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(new UserControl(), "ตรวจสอบชนิดตัวอักษร");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.fontUC, "ตรวจสอบขนาดกับชนิดตัวอักษร");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.royalWordUC, "ตรวจสอบคำตามศัพท์บรรญัติ");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.punctuationUC, "ตรวจสอบเครื่องหมายวรรคตอน");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
            Ribbon1.myCustomTaskPane = Globals.ThisAddIn.CustomTaskPanes.Add(Ribbon1.referenceModelUC, "ตรวจสอบรูปแบบอ้างอิง");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = false;
        }

        static public void showCustomTaskPane(int show = -1, bool showAll = false)
        {
            Globals.ThisAddIn.CustomTaskPanes[0].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[1].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[2].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[3].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[4].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[5].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[6].Visible = false;
            Globals.ThisAddIn.CustomTaskPanes[7].Visible = false;
            if (show == -1)
            {
                return;
            }
            Globals.ThisAddIn.CustomTaskPanes[show].Visible = true;
            if (showAll)
            {
                Ribbon1.showCheckAllUC.setButtonClickALL();
                Globals.ThisAddIn.CustomTaskPanes[0].Visible = true;
            }

        }

        private void btn_SaveNewFile_Click(object sender, RibbonControlEventArgs e)
        {
            Word._Application oWord = Globals.ThisAddIn.Application;
            oWord.Visible = true;

            //object fileName = "NewDocument"+i+".docx";
            object fileName = nameFile+"_ปริญญานิพนธ์ใหม่" + i + ".docx";
            i++;
            string pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string path = pathDoc + "\\" + fileName;
            while (File.Exists(path))
            {
                //System.Windows.Forms.MessageBox.Show(path);
                fileName = nameFile + "_ปริญญานิพนธ์ใหม่" + i + ".docx";
                path = pathDoc + "\\" + fileName;
                i++;
            }
            object missing = System.Reflection.Missing.Value;
            //oWord.ActiveDocument.SaveAs(fileName);
            //oWord.Documents.Add(@"C:\NewDocument.docx");
            //oWord.Options.CreateBackup = true;
            //oWord.ActiveDocument.Optio
            oWord.ActiveDocument.SaveAs2(fileName, ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        static public void saveFileAuto()
        {
            Word._Application oWord = Globals.ThisAddIn.Application;
            oWord.Visible = true;

            //object fileName = "NewDocument"+i+".docx";
            object fileName = nameFile + "_ปริญญานิพนธ์" + Ribbon1.namefileSaveAuto + "ใหม่.docx";
            string pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
            string path = pathDoc+ "\\" + fileName;
            int count = 1;
            while (File.Exists(path))
            {
                //System.Windows.Forms.MessageBox.Show(path);
                fileName = nameFile + "_ปริญญานิพนธ์" + Ribbon1.namefileSaveAuto + "(" + count + ")ใหม่.docx";
                path = pathDoc + "\\" + fileName;
                count++;
            }
            object missing = System.Reflection.Missing.Value;
            //oWord.ActiveDocument.SaveAs(fileName);
            //oWord.Documents.Add(@"C:\NewDocument.docx");
            
            oWord.ActiveDocument.SaveAs(fileName, ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing);
        }

        private void btn_checkReference_Click(object sender, RibbonControlEventArgs e)
        {
            //FindAndReplace("ben","orojiben");

            saveFileAuto();
            Ribbon1.addCustomTaskPaneALL();
            Ribbon1.referenceModel.showUC = true;
            Ribbon1.referenceModel.runCheckReferenceAll();
        }

    }
}
