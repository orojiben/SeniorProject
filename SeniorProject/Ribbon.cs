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
        Word.Range rng;
        float leftDf;
        float rightDf;
        float topDf;
        float bottomDf;
        List<MarginPage> list;
        List<Styles> loadStyles;
        private List<string> font;
        ReferenceModel rm;

        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
            rm = new ReferenceModel();
            loadStyles = StyleFile.LoadStyle();
            this.rng = Globals.ThisAddIn.Application.ActiveDocument.Range(0, 0);
            this.leftDf = this.rng.PageSetup.LeftMargin;
            this.rightDf = this.rng.PageSetup.RightMargin;
            this.topDf = this.rng.PageSetup.TopMargin;
            this.bottomDf = this.rng.PageSetup.BottomMargin;
            list = new List<MarginPage>();


            readFileToList();
            // Word.Document save = Globals.ThisAddIn.Application.ActiveDocument;

            //  save.Application.ActiveDocument.SaveAs2()
        }

        private void comboBox1_TextChanged(object sender, RibbonControlEventArgs e)
        {

            // Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl = (RibbonDropDownItem)this.comboBox1.Items.GetEnumerator();
            // MessageBox.Show(ribbonDropDownItemImpl.Label);
            //string[] words = buff.Split(',');
        }

        private void readFileToList()
        {
            // list.Add(new MarginPage("Defalf", this.leftDf, this.rightDf, this.topDf, this.bottomDf));



            try
            {
                this.dropDown1.Items.Clear();
                Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl = this.Factory.CreateRibbonDropDownItem();
                ribbonDropDownItemImpl.Label = "Default";
                this.dropDown1.Items.Add(ribbonDropDownItemImpl);
                foreach (Styles s in this.loadStyles)
                {
                    Microsoft.Office.Tools.Ribbon.RibbonDropDownItem ribbonDropDownItemImpl1 = this.Factory.CreateRibbonDropDownItem();
                    ribbonDropDownItemImpl1.Label = s.Name;
                    this.dropDown1.Items.Add(ribbonDropDownItemImpl1);
                }

            }
            catch { };

        }

        private void dropDown1_SelectionChanged(object sender, RibbonControlEventArgs e)
        {
            //this.dropDown1.SelectedItem.Label;
            /*MarginPage mp = list[this.dropDown1.SelectedItemIndex];
            float s = mp.getLeft();
            this.rng.PageSetup.LeftMargin = mp.getLeft();
            this.rng.PageSetup.RightMargin = mp.getRight();
            this.rng.PageSetup.TopMargin = mp.getTop();
            this.rng.PageSetup.BottomMargin = mp.getBottom();*/
            if (this.dropDown1.SelectedItemIndex != 0)
            {

                Styles s = this.loadStyles[this.dropDown1.SelectedItemIndex - 1];
                string[] words = s.Margin.Split(',');

                this.rng.PageSetup.LeftMargin = centimeterToPoint((float)(Convert.ToDouble(words[0])));
                this.rng.PageSetup.RightMargin = centimeterToPoint((float)(Convert.ToDouble(words[1])));
                this.rng.PageSetup.TopMargin = centimeterToPoint((float)(Convert.ToDouble(words[2])));
                this.rng.PageSetup.BottomMargin = centimeterToPoint((float)(Convert.ToDouble(words[3])));
                font = s.Font;

            }
            else
            {
                this.rng.PageSetup.LeftMargin = this.leftDf;
                this.rng.PageSetup.RightMargin = this.rightDf;
                this.rng.PageSetup.TopMargin = this.topDf;
                this.rng.PageSetup.BottomMargin = this.bottomDf;
            }
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
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
            rm.runCheckReferenceAll();
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
            FontManager.CheckFontName(font[0]);
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
            rm.runCheckReferenceAll();
            Verify_Royal_Word_TH verify_th = new Verify_Royal_Word_TH();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        private void button1_Click_1(object sender, RibbonControlEventArgs e)
        {
            FontManager.CheckFontSize(16, 18, 20);
        }

        /*Match match = Regex.Match(this.code, @"^[\(\)\{};]");
            if (match.Success)
            {
                this.code = this.code.Remove(0, 1); 
                this.Position += 1;
                return new Token(match.Value);*/

    }
}
