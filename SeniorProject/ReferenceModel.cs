using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
     class ReferenceModel
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
                int value = ModelBookTypeBookTH(r, cout);
               if (value == 0)
               {
                   break;
               }
               cout += value;
               r = r.Application.ActiveDocument.Range(cout);
               //System.Windows.Forms.MessageBox.Show(rr.Text + " ^^");
               
            }
        }

        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookTH(Word.Range r, int cout)
        {
            string strFromRange = r.Text;
            int countBuff = cout;
            int coutNew = 0;
            if (CutStringMatch(ref strFromRange, @"(((([ก-ฮะ-์])*(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+\.\s)", ref coutNew))
            {
                if (CutStringMatch(ref strFromRange, @"(\([1-9][0-9]{3}\)\.\s)", ref coutNew))
                {
                    if (CutStringMatch(ref strFromRange, @"(((([ก-ฮะ-์])*(\:)?(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+\.\s)", ref coutNew))
                    {
                        if (CutStringMatch(ref strFromRange, @"((\(((([ก-ฮะ-์])*(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+[ผ][ู][้][แ][ป][ล]\)\.\s)?)", ref coutNew))
                        {
                            if (CutStringMatch(ref strFromRange, @"(((([ก-ฮะ-์])*(ฯ)?(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)\:\s)", ref coutNew))
                            {
                                if (CutStringMatch(ref strFromRange, @"(((([ก-ฮะ-์])*(ฯ)?(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)\.(\s)?)", ref coutNew))
                                {
                                    if (CutStringMatch(ref strFromRange, @"(\((((([0-9ก-ฮะ-์])*(\.)?(\:)?(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)*)?\)(\s)?)", ref coutNew))
                                    {
                                        if (CutStringMatch(ref strFromRange, @"((\n|\r))", ref coutNew))
                                        {
                                            Word.Range rCheck = r.Application.ActiveDocument.Range(countBuff, countBuff + coutNew);
                                            int countCheck = 0;
                                            foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                                            {
                                                if (rngWord.Text == "). ")
                                                {
                                                    countCheck++;
                                                    //   System.Windows.Forms.MessageBox.Show(range2.Text);
                                                }
                                                else if (countCheck == 1)
                                                {

                                                    if (rngWord.Bold != 0)
                                                    {
                                                        System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");
                                                        if (rngWord.Text[rngWord.Text.Length - 2] == '.')
                                                        {
                                                            countCheck++;
                                                        }

                                                    }
                                                    else if (countCheck == 2)
                                                    {
                                                        return countBuff + coutNew;
                                                    }
                                                    else
                                                    {

                                                        break;
                                                    }

                                                }
                                            }
                                           
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeArticleTH(r, cout);
        }


        private int ModelBookTypeArticleTH(Word.Range r, int cout)
        {
            string strFromRange = r.Text;
            int countBuff = cout;
            int coutNew = 0;
            if (CutStringMatch(ref strFromRange, @"(((([ก-ฮะ-์])*(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+\.\s)", ref coutNew))
            {
                if (CutStringMatch(ref strFromRange, @"(\([1-9][0-9]{3}\)\.\s)", ref coutNew))
                {

                }
            }
            return 0;
        }

        private bool CutStringMatch(ref string strFromRange, string model,ref int countCut)
        {
            Match match = Regex.Match(strFromRange, model);
            if (match.Success)
            {
                strFromRange = strFromRange.Remove(0, match.Value.Length);
                countCut += match.Value.Length;
                return true;
            }
            return false;
        }

        //หนังสือทั่วไป เอกสารประเภทหนังสือ
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
            return Model_2(r, cout);
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int Model_2(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^([A-Z]([a-zA-Z])*(\,\s([A-Z]\.(\s)?)*)?((([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)*[a][n][d]\s([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)?\([0-9]{4}\)\.((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\.\s[I][n](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*(\,((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)*\s[a][n][d](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)?\s\([E][d][s]\.\)\,\s([A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\s\((([0-9a-zA-Z]|(\s))*\.\,(\s))?(([p]\.\s[1-9]([0-9])*)|([p][p]\.\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*))\)\.(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))(\,(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)))*:(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))\.(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == ".), ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == " (" || rngWord.Text == "(")
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
            return Model_3(r, cout);
        }

        //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int Model_3(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^([A-Z]([a-zA-Z])*(\,\s([A-Z]\.(\s)?)*)?((([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)*[a][n][d]\s([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)?\([0-9]{4}\)\.((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\.\s[I][n](\s([A-Z]\.(\s)?)*)?((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*(\,((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*)*\s[a][n][d](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*)?\s\(([V][o][l]\.([1-9](0-9)*)\,(\s))?(([p]\.\s[1-9]([0-9])*)|([p][p]\.\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*))\)\.(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))(\,(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)))*:(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))\.(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "In ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == " (" || rngWord.Text == "(")
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
            return Model_4(r, cout);
        }

        //เอกสารประกอบการบรรยาย เอกสารประเภทหนังสือ v1
        private int Model_4(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*\s)+\([ผ][ู][้][บ][ร][ร][ย][า][ย]\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])*(\s)?)+\.\s[ใ][น]\s((([0-9])|([ก-ฮะ-์])|([:]))*(\s)?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])*(\s)?)+\:\s(([ก-ฮะ-์])*(\s)?)+\.(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "ใน ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == " (" || rngWord.Text == "(")
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
            return Model_4_2(r, cout); ;
        }

        //เอกสารประกอบการบรรยาย เอกสารประเภทหนังสือ v2
        private int Model_4_2(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*\s)+\((([ผ][ู][้][บ][ร][ร][ย][า][ย])|([ผ][ู][้][ป][า][ฐ][ก][ถ][า]))\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])*(\s)?)+\.\s(([ก-ฮะ-์])*(\s)?)+\:\s(([ก-ฮะ-์])*(\s)?)+\.(\n|\r))");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "). ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == "." || rngWord.Text == ". ")
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
            return Model_5( r,  cout);
        }

        //เอกสารประเภทวารสาร บทความทั่วไป
        private int Model_5(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;

            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*(\s)?)+\.\s\([1-9][0-9]{3}\)\.\s((((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s){2})");
            
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "). " | rngWord.Text == ". " | rngWord.Text == "?. ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == "." || rngWord.Text == ". ")
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
            return Model_6(r, cout);
        }

        //เอกสารประเภทวารสาร บทวิจารณ์และบทความปริทัศน์หนังสือ
        private int Model_6(Word.Range r, int cout)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*(\s)?)+\.\s\([1-9][0-9]{3}\)\.\s(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s\[(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\]\.\s(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s[1-9]([0-9])*\([1-9]([0-9])*\)\,\s[1-9]([0-9])*((\s)?\-[1-9]([0-9])*(\s)?)?\.)");
            if (match.Success)
            {
                str = str.Remove(0, match.Value.Length);
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + match.Value.Length);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "). " | rngWord.Text == ". " | rngWord.Text == "?. ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2)
                    {

                        if (rngWord.Bold != 0)
                        {
                            System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else if (rngWord.Text == "." || rngWord.Text == ". ")
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
