using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

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

        public void SampleRegexUsage(Word.Range r,ref bool check)
        {
            string regex = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([บ][ร][ร][ณ][า][ธ][ิ][ก][า][ร]\)\,\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\r";
            //string regex = @"(((([ก-ฮะ-์])*(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+\.\s)";
            RegexOptions options = RegexOptions.RightToLeft | RegexOptions.None;
            string input = r.Text;
            int a = 9;
            try {

                Match matche = Regex.Match(input, regex);
               
           if (matche.Success)
            {
                check = true;
                a=1;
            }
            int b = 9;
            }
            catch (ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("");
                // Do nothing: assume that exception represents no match.
            }
        }

        private void FindReferences(Word.Range r)
        {
            bool check = false;
           // SampleRegexUsage(r,ref check);
            string[] strSpliteRanges = Regex.Split(r.Text, "\r");
            int cout = 0;
            foreach (string strSpliteRange in strSpliteRanges)
            {
                if (strSpliteRange.Length == 0)
                {
                    break;
                }
                int value = ModelBookTypeBookTH(r, strSpliteRange, cout);// ModelBookTypeBookTH(r, cout);
                if (value == 0)
                {
                    break;
                }
                cout += value;
                r = r.Application.ActiveDocument.Range(cout);
            }

            /*int cout = 0;

            while (true)
            {//((((([\(])|([0-9a-zA-Zก-ฮะ-์])|([\)])|(\.)|(\,)|(\:)|([ \f\t\v]))*)(\n|\r)))
                int value = ModelBookTypeBookTH(r, cout);// ModelBookTypeBookTH(r, cout);
               if (value == 0)
               {
                   break;
               }
               cout += value;
               r = r.Application.ActiveDocument.Range(cout);
               //System.Windows.Forms.MessageBox.Show(rr.Text + " ^^");
               
            }*/
        }

          //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookTH(Word.Range r,string strCheck ,int cout)
        {
            string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+[ผ][ู][้][แ][ป][ล]\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\((([0-9ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\))";
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(strCheck, regex, ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
           
            
            if (checkValue!=-1)
            {
                System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "). ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 1 || countCheck == 2)
                    {

                        if (rngWord.Bold != 0)
                        {
                            if (countCheck == 2)
                            {
                                countCheck = 1;
                            }
                            //System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");
                            if (rngWord.Text[rngWord.Text.Length - 2] == '.')
                            {
                                countCheck++;
                            }

                        }
                        else if (countCheck == 2)
                        {
                            System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ จบ");
                            cout = +checkValue;
                            return cout;
                        }
                        else
                        {

                            break;
                        }

                    }
                }
            }

            return ModelBookTypeArticleTH( r, strCheck, cout);
            
        }
        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int ModelBookTypeArticleTH(Word.Range r,string strCheck, int cout)
        {
            
            int checkValue = -1;
            int checkValue2 = -1;
            string str = r.Text;
            string []strCheckSplites = Regex.Split(strCheck,". ใน ");
            if (strCheckSplites.Length == 2)
            {
                string strFirst = strCheckSplites[0]+". ใน ";
                string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s[ใ][น]\s)";
                var task = Task.Factory.StartNew(() => CheckStringMatch(strFirst, regex, ref checkValue) );
                var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

                string strSecond = strCheckSplites[1];
                string regex2 = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([บ][ร][ร][ณ][า][ธ][ิ][ก][า][ร]\)\,\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+";
                var task2 = Task.Factory.StartNew(() => CheckStringMatch(strSecond, regex2, ref checkValue2));
                var completedWithinAllotedTime2 = task2.Wait(TimeSpan.FromMilliseconds(1000));
            }
            else
            {
                return 0;
            }

            if (checkValue != -1 && checkValue2 != -1)
            {
                checkValue += checkValue2-5;
                System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    if (rngWord.Text == "). " || rngWord.Text == "), ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2)
                    {

                        if (rngWord.Bold != 0)
                        {
                            //System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");

                        }
                        else
                        {
                            if (rngWord.Text == "(" || rngWord.Text == "หน้า" || rngWord.Text == "หน้า ")
                            {
                                countCheck++;
                            }
                            else
                            {
                                break;
                            }
                        }

                    }
                    else if (countCheck == 4)
                    {
                        System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ จบ");
                        cout = +checkValue;
                        return cout;
                    }
                }
            }
            return 0;
        }


        private void CheckStringMatch(string strFromRange, string regex, ref int checkValue)
        {
            Match match = Regex.Match(strFromRange, regex);
            if (match.Success)
            {
                checkValue = match.Value.Length;
                return;
            }
                checkValue= -1;
        }


        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int Model_1(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^(([a-zA-Z])*((\,)(\s)(([a-zA-Z])+(\.)(\s)?)*)?((\s)[a][n][d](\s)([a-zA-Z])*((\,)((\s)([a-zA-Z])+(\.))*)?)?((\s)(\()([0-9]{4})(\)(\.)))((\s)([a-zA-Z])([a-zA-Z]|(\s))*(\.))((\s)(([a-zA-Z]|(\s))*(\,)(\s))*([a-zA-Z]|(\s))*(\:)(\s)([a-zA-Z]|(\s))*(\.))(\n|\r))", options);
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
            return Model_2(r, cout, options);
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int Model_2(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^([A-Z]([a-zA-Z])*(\,\s([A-Z]\.(\s)?)*)?((([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)*[a][n][d]\s([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)?\([0-9]{4}\)\.((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\.\s[I][n](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*(\,((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)*\s[a][n][d](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)?\s\([E][d][s]\.\)\,\s([A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\s\((([0-9a-zA-Z]|(\s))*\.\,(\s))?(([p]\.\s[1-9]([0-9])*)|([p][p]\.\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*))\)\.(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))(\,(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)))*:(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))\.(\n|\r))", options);
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
            return Model_3(r, cout, options);
        }

        //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int Model_3(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^([A-Z]([a-zA-Z])*(\,\s([A-Z]\.(\s)?)*)?((([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)*[a][n][d]\s([a-zA-Z])*(\,\s([A-Z]\.\s)*)?)?\([0-9]{4}\)\.((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*\.\s[I][n](\s([A-Z]\.(\s)?)*)?((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*(\,((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*)*\s[a][n][d](\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)([a-zA-Z])*)*)?\s\(([V][o][l]\.([1-9](0-9)*)\,(\s))?(([p]\.\s[1-9]([0-9])*)|([p][p]\.\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*))\)\.(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))(\,(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)))*:(((\s([A-Z]\.(\s)?)*)((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*)|(((\s)[A-Z]([a-zA-Z])*)((\s)[A-Z]([a-zA-Z])*)*))\.(\n|\r))", options);
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
            return Model_4(r, cout, options);
        }

        //เอกสารประกอบการบรรยาย เอกสารประเภทหนังสือ v1
        private int Model_4(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*\s)+\([ผ][ู][้][บ][ร][ร][ย][า][ย]\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])*(\s)?)+\.\s[ใ][น]\s((([0-9])|([ก-ฮะ-์])|([:]))*(\s)?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])*(\s)?)+\:\s(([ก-ฮะ-์])*(\s)?)+\.(\n|\r))", options);
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
            return Model_4_2(r, cout, options);
        }

        //เอกสารประกอบการบรรยาย เอกสารประเภทหนังสือ v2
        private int Model_4_2(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*\s)+\((([ผ][ู][้][บ][ร][ร][ย][า][ย])|([ผ][ู][้][ป][า][ฐ][ก][ถ][า]))\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])*(\s)?)+\.\s(([ก-ฮะ-์])*(\s)?)+\:\s(([ก-ฮะ-์])*(\s)?)+\.(\n|\r))", options);
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
            return Model_5(r, cout, options);
        }

        //เอกสารประเภทวารสาร บทความทั่วไป
        private int Model_5(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;

            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*(\s)?)+\.\s\([1-9][0-9]{3}\)\.\s((((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s){2})", options);
            
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
            return Model_6(r, cout, options);
        }

        //เอกสารประเภทวารสาร บทวิจารณ์และบทความปริทัศน์หนังสือ
        private int Model_6(Word.Range r, int cout, RegexOptions options)
        {
            Match match;
            string str = r.Text;
            match = Regex.Match(r.Text, @"^((([ก-ฮะ-์])*(\s)?)+\.\s\([1-9][0-9]{3}\)\.\s(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s\[(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\]\.\s(((\“([ก-ฮะ-์])*\”)|([ก-ฮะ-์])*(\?)?)(\s)?)+\.\s[1-9]([0-9])*\([1-9]([0-9])*\)\,\s[1-9]([0-9])*((\s)?\-[1-9]([0-9])*(\s)?)?\.)", options);
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
