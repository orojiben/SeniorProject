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
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {
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
            this.CeckBio();
        }

        private void FindAndReplace(
                            object findText,
                            object replaceWithText)
        {
            object matchCase = true;
            object matchWholeWord = true;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object nmatchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 2;
            object wrap = 1;
            var wordApp = Globals.ThisAddIn.Application;

            List<Word.Range> lsR = new List<Word.Range>();
            List<string> lsS = new List<string>();
            foreach (Word.Range range in wordApp.ActiveDocument.StoryRanges)
            {
                FindReferences(lsR, lsS, range, wordApp);
                foreach (Word.Range rS in lsR)
                {
                    //System.Windows.Forms.MessageBox.Show(rS.Text + "");
                }
                int c = 0;
                Match match;

                foreach (Microsoft.Office.Interop.Word.Range rngWord in range.Words)
                {
                    Word.Range range2 = rngWord;
                    //  System.Windows.Forms.MessageBox.Show(range2.Text + " " + range2.Text.Length);
                    if (range2.Text == "(" || range2.Text == "). " || range2.Text == ")" || range2.Text == ").")
                    {
                        c++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (c == 1)
                    {
                        match = Regex.Match(range2.Text, @"^([0-9]{4})");
                        if (match.Success)
                        {
                            c++;
                            //   System.Windows.Forms.MessageBox.Show(range2.Text);
                        }

                    }
                    else if (c == 3)
                    {
                        if (range2.Text == ".")
                        {
                            c = 0;
                        }
                        else
                        {
                            if (range2.Bold != 0)
                            {
                                //System.Windows.Forms.MessageBox.Show(range2.Text + " " + range2.Font.NameBi);
                            }
                        }

                    }
                    while (range2 != null)//&& range2.Bold != 0)
                    {
                        // System.Windows.Forms.MessageBox.Show(range2.ParagraphFormat.Alignment.ToString());
                        range2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        /*range2.Find.Execute(ref findText, ref matchCase, ref matchWholeWord, ref matchWildCards,
                            ref matchSoundsLike, ref nmatchAllWordForms, ref forward, ref wrap, ref format,
                             findText, ref replace, ref matchKashida, ref matchDiacritics, ref matchAlefHamza,
                            ref matchControl);*/
                        if (range2.Text[0] == 13)
                        {
                            // System.Windows.Forms.MessageBox.Show("1");
                        }
                        else if (range2.Text[0] == 11)
                        {
                            //  System.Windows.Forms.MessageBox.Show("2");
                        }

                        // int s = range2.Text[0];
                        // System.Windows.Forms.MessageBox.Show(s+"");

                        range2 = range2.NextStoryRange;




                    }

                }
            }
        }

        private void CeckBio()
        {
            var wordApp = Globals.ThisAddIn.Application;

            List<Word.Range> lsR = new List<Word.Range>();
            List<string> lsS = new List<string>();
            foreach (Word.Range range in wordApp.ActiveDocument.StoryRanges)
            {
                FindReferences(lsR, lsS, range, wordApp);
                foreach (Word.Range rS in lsR)
                {
                    System.Windows.Forms.MessageBox.Show(rS.Text + "");
                }
                int c = 0;
                Match match;

                foreach (Microsoft.Office.Interop.Word.Range rngWord in range.Words)
                {
                    Word.Range range2 = rngWord;
                    //  System.Windows.Forms.MessageBox.Show(range2.Text + " " + range2.Text.Length);
                    if (range2.Text == "(" || range2.Text == "). " || range2.Text == ")" || range2.Text == ").")
                    {
                        c++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (c == 1)
                    {
                        match = Regex.Match(range2.Text, @"^([0-9]{4})");
                        if (match.Success)
                        {
                            c++;
                            //   System.Windows.Forms.MessageBox.Show(range2.Text);
                        }

                    }
                    else if (c == 3)
                    {
                        if (range2.Text == ".")
                        {
                            c = 0;
                        }
                        else
                        {
                            if (range2.Bold != 0)
                            {
                                //System.Windows.Forms.MessageBox.Show(range2.Text + " " + range2.Font.NameBi);
                            }
                        }

                    }
                    while (range2 != null)//&& range2.Bold != 0)
                    {
                        range2.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        if (range2.Text[0] == 13)
                        {
                            // System.Windows.Forms.MessageBox.Show("1");
                        }
                        else if (range2.Text[0] == 11)
                        {
                            //  System.Windows.Forms.MessageBox.Show("2");
                        }

                        // int s = range2.Text[0];
                        // System.Windows.Forms.MessageBox.Show(s+"");

                        range2 = range2.NextStoryRange;
                    }

                }
            }
        }

        private void FindReferences(List<Word.Range> lsR, List<string> lsS, Word.Range r, Word.Application doc)
        {
            string str = r.Text;
            lsR.Clear();

            Match match;
            int cout = 0;
            // System.Windows.Forms.MessageBox.Show(str);
            while (true)
            {//((((([\(])|([0-9a-zA-Zก-ฮะ-์])|([\)])|(\.)|(\,)|(\:)|([ \f\t\v]))*)(\n|\r)))
                match = Regex.Match(str, @"^(([a-zA-Z])*((\,)(\s)(([a-zA-Z])+(\.)(\s)?)*)?((\s)[a][n][d](\s)([a-zA-Z])*((\,)((\s)([a-zA-Z])+(\.))*)?)?((\s)(\()([0-9]{4})(\)(\.)))((\s)([a-zA-Z])([a-zA-Z]|(\s))*(\.))((\s)(([a-zA-Z]|(\s))*(\,)(\s))*([a-zA-Z]|(\s))*(\:)(\s)([a-zA-Z]|(\s))*(\.))(\n|\r))");
                if (match.Success)
                {
                    str = str.Remove(0, match.Value.Length);
                    lsS.Add(match.Value);
                    lsR.Add(doc.ActiveDocument.Range(cout, cout + match.Value.Length));
                    //System.Windows.Forms.MessageBox.Show(match.Value);
                    //System.Windows.Forms.MessageBox.Show(str);
                    cout = +match.Value.Length;
                }
                else
                {
                    break;
                }
            }
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
            Globals.ThisAddIn.GetFont(font[0]);
            
            //Globals.ThisAddIn.GetFont("Tahoma");
        }

        private void btn_correctFont_Click(object sender, RibbonControlEventArgs e)
        {
            Globals.ThisAddIn.CorrectFont(font[0]);
            //Globals.ThisAddIn.CorrectFont("Tahoma");
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
            this.CeckBio();
            Verify_Royal_Word_TH verify_th = new Verify_Royal_Word_TH();
        }

        private void button6_Click(object sender, RibbonControlEventArgs e)
        {

        }

        /*Match match = Regex.Match(this.code, @"^[\(\)\{};]");
            if (match.Success)
            {
                this.code = this.code.Remove(0, 1); 
                this.Position += 1;
                return new Token(match.Value);*/

    }
}
