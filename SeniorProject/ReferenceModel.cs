using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    public class ReferenceModel
    {
        public void runCheckReferenceAll()
        {
            CeckBio();
        }

        

        private void CeckBio()
        {
            var wordApp = Globals.ThisAddIn.Application;

            List<Word.Range> lsR = new List<Word.Range>();
            List<string> lsS = new List<string>();
            foreach (Word.Range range in wordApp.ActiveDocument.StoryRanges)
            {
                FindReferences(range);
                
                //System.Windows.Forms.MessageBox.Show(range.Text + "");
                /*foreach (Word.Range rS in lsR)
                {
                    System.Windows.Forms.MessageBox.Show(rS.Text + "");
                }
                int c = 0;
                Match match;

                foreach (Microsoft.Office.Interop.Word.Range rngWord in range.Words)
                {
                    Word.Range range2 = rngWord;
                    //  System.Windows.Forms.MessageBox.Show(range2.Text + " " + range2.Text.Length);
                    if (range2.Text == "(" || range2.Text == "). " || range2.Text == ")")
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
                                System.Windows.Forms.MessageBox.Show(range2.Text);
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
                        // System.Windows.Forms.MessageBox.Show("");

                        range2 = range2.NextStoryRange;
                    }

                }*/
            }
        }


        private void FindReferencesTestV1(List<Word.Range> lsR, List<string> lsS, Word.Range r, Word.Application doc)
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
                    //Word.Range newr = r;
                    //r.
                    str = str.Remove(0, match.Value.Length);
                    lsS.Add(match.Value);
                    //lsR.Add(doc.ActiveDocument.Range(cout, cout + match.Value.Length));
                    //newr.SetRange(cout, cout + match.Value.Length);
                    //Word.Range buff = newr;
                    lsR.Add(r.Application.ActiveDocument.Range(cout, cout + match.Value.Length));
                    //System.Windows.Forms.MessageBox.Show(match.Value);
                    //System.Windows.Forms.MessageBox.Show(buff.Text + " _");
                    cout = +match.Value.Length;
                }
                else
                {
                    break;
                }
            }
        }

        private void FindReferences(Word.Range r)
        {
            string str = r.Text;

            int cout = 0;
            // System.Windows.Forms.MessageBox.Show(str);
            while (true)
            {//((((([\(])|([0-9a-zA-Zก-ฮะ-์])|([\)])|(\.)|(\,)|(\:)|([ \f\t\v]))*)(\n|\r)))
               int value = Model_1(r, cout);
               if (value == 0)
               {
                   value = Model_2(r, cout);
                   if (value == 0)
                   {
                       break;
                   }
               }
               cout += value;
               r = r.Application.ActiveDocument.Range(cout);
               //System.Windows.Forms.MessageBox.Show(rr.Text + " ^^");
               
            }
        }

        private int Model_1(Word.Range r,int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^(([a-zA-Z])*((\,)(\s)(([a-zA-Z])+(\.)(\s)?)*)?((\s)[a][n][d](\s)([a-zA-Z])*((\,)((\s)([a-zA-Z])+(\.))*)?)?((\s)(\()([0-9]{4})(\)(\.)))((\s)([a-zA-Z])([a-zA-Z]|(\s))*(\.))((\s)(([a-zA-Z]|(\s))*(\,)(\s))*([a-zA-Z]|(\s))*(\:)(\s)([a-zA-Z]|(\s))*(\.))(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                 Word.Range rCheck =  r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                 int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "(" || rngWord.Text == "). " || rngWord.Text == ")")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1)
                    {
                        Match matchBC = Regex.Match(rngWord.Text, @"^([0-9]{4})");
                        if (matchBC.Success)
                        {
                            countCheck++;
                            //   System.Windows.Forms.MessageBox.Show(range2.Text);
                        }

                    }
                    else if (countCheck == 3)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text+"_");
                            
                        }
                        else if (rngWord.Text == ". ")
                        {
                            cout = +match.Value.Length;
                            return cout;
                        }
                        else
                        {

                            break;
                        }

                    }
                }
               // lsS.Add(match.Value);
               // lsR.Add(doc.ActiveDocument.Range(cout, cout + match.Value.Length));
                //System.Windows.Forms.MessageBox.Show(match.Value);
                //System.Windows.Forms.MessageBox.Show(str);
                
            }
            return 0;
        }

        private int Model_2(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^(([a-zA-Z])*((\,)(\s)(([a-zA-Z])+(\.)(\s)?)*)?((\s)[a][n][d](\s)([a-zA-Z])*((\,)((\s)([a-zA-Z])+(\.))*)?)?((\s)(\()([0-9]{4})(\)(\.)))((\s)([a-zA-Z])([a-zA-Z]|(\s))*(\.))((\s)(([a-zA-Z]|(\s))*(\,)(\s))*([a-zA-Z]|(\s))*(\:)(\s)([a-zA-Z]|(\s))*(\.))(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "(" || rngWord.Text == "). " || rngWord.Text == ")")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1)
                    {
                        Match matchBC = Regex.Match(rngWord.Text, @"^([0-9]{4})");
                        if (matchBC.Success)
                        {
                            countCheck++;
                            //   System.Windows.Forms.MessageBox.Show(range2.Text);
                        }

                    }
                    else if (countCheck == 3)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == ". ")
                        {
                            cout = +match.Value.Length;
                            return cout;
                        }
                        else
                        {

                            break;
                        }

                    }
                }
                // lsS.Add(match.Value);
                // lsR.Add(doc.ActiveDocument.Range(cout, cout + match.Value.Length));
                //System.Windows.Forms.MessageBox.Show(match.Value);
                //System.Windows.Forms.MessageBox.Show(str);

            }
            return 0;
        }

    }
}
