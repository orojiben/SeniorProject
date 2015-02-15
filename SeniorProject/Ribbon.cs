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
        MarginPage marginPage;
        PaperPage paperPage;
        List<Styles> loadStyles;
        private List<string> font;
        ReferenceModel referenceModel;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            referenceModel = new ReferenceModel();
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
            loadDataStyles(this.ddn_Model.SelectedItemIndex);
        }

        int i = 0;
        private void button1_Click(object sender, RibbonControlEventArgs e)
        {
            Word._Application oWord = Globals.ThisAddIn.Application;
            string s = Globals.ThisAddIn.Application.ActiveDocument.FullName;
            oWord.Visible = true;

            //object fileName = "NewDocument"+i+".docx";
            object fileName = "NewDocument.docx";
            i++;
            //oWord.ActiveDocument.SaveAs(fileName);
            //oWord.Documents.Add(@"C:\NewDocument.docx");
            oWord.ActiveDocument.SaveAs(fileName);
        }

        private void button2_Click(object sender, RibbonControlEventArgs e)
        {
            //FindAndReplace("ben","orojiben");
            referenceModel.runCheckReferenceAll();
        }

        

        private void btn_checkRoyalWord_Click(object sender, RibbonControlEventArgs e)
        {
            Verify_Royal_Word_TH verify_th = new Verify_Royal_Word_TH();
        }

        private void btn_checkSing_Click(object sender, RibbonControlEventArgs e)
        {
            Verify_Space_Sign verify = new Verify_Space_Sign();
        }

        private void btn_checkFont_Click(object sender, RibbonControlEventArgs e)
        {
            this.ShowFont();
        }

        private void ShowFont()
        {
            //FontManager.CheckFontName("abc","def",Globals.ThisAddIn.Application.ActiveDocument);
            FontManager.CheckFontName("Angsana New");
        }

        private void btn_correctFont_Click(object sender, RibbonControlEventArgs e)
        {
            FontManager.CorrectFont(font[0]);
        }

        private void button3_Click(object sender, RibbonControlEventArgs e)
        {
            var wordApp = Globals.ThisAddIn.Application;


            foreach (Word.Range range in wordApp.ActiveDocument.StoryRanges)
            {
                foreach (Microsoft.Office.Interop.Word.Range rngWord in range.Words)
                {
                    System.Windows.Forms.MessageBox.Show(rngWord.Text + " " + rngWord.Font.Name);
                }
            }
        }

        private void btn_checkAll_Click(object sender, RibbonControlEventArgs e)
        {

            this.ShowFont();
            referenceModel.runCheckReferenceAll();
            Verify_Royal_Word_TH verify_th = new Verify_Royal_Word_TH();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            FontManager.CheckFontSize(16, 18, 20);
        }

        private void button5_Click(object sender, RibbonControlEventArgs e)
        {
            this.marginPage.changing();
        }

        private void button4_Click(object sender, RibbonControlEventArgs e)
        {
            //System.Windows.Forms.MessageBox.Show(this.paperPage.cheking()+"");
            this.paperPage.changing();
        }

        private void loadDataStyles(int index)
        {
            Styles styles = this.loadStyles[index];
            string[] words = styles.Margin.Split(',');

            float leftMargin = centimeterToPoint((float)(Convert.ToDouble(words[0])));
            float rightMargin = centimeterToPoint((float)(Convert.ToDouble(words[1])));
            float topMargin = centimeterToPoint((float)(Convert.ToDouble(words[2])));
            float bottomMargin = centimeterToPoint((float)(Convert.ToDouble(words[3])));
            marginPage = new MarginPage(leftMargin, rightMargin, topMargin, bottomMargin);
            paperPage = new PaperPage(styles.Paper);
            font = styles.Fonts;
            this.ddn_Department.Items.Clear();
            this.ddn_Department.Visible = false;
            this.referenceModel.faculty = "";
            this.referenceModel.department = "";
            if (styles.Departments.Count > 0)
            {
                this.ddn_Department.Visible = true;
                foreach (string departments in styles.Departments)
                {
                    Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItemImpl1.Label = departments;
                    this.ddn_Department.Items.Add(ribbonDropDownItemImpl1);
                }
                this.referenceModel.department = styles.Departments[0];
            }
            this.referenceModel.faculty = styles.Name;
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
        }

        private void ddn_Department_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            this.referenceModel.department = this.ddn_Department.Items[this.ddn_Department.SelectedItemIndex].Label;
        }
    }
}
