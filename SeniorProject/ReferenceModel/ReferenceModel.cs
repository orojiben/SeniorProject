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
            //RegexOptions options = RegexOptions.RightToLeft | RegexOptions.None;
            string input = r.Text;
            //int a = 9;
            try {

                Match matche = Regex.Match(input, regex);
               
           if (matche.Success)
            {
                check = true;
              //  a=1;
            }
            //int b = 9;
            }
            catch (ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("");
                // Do nothing: assume that exception represents no match.
            }
        }

        private void FindReferences(Word.Range r)
        {
            //bool check = false;
           // SampleRegexUsage(r,ref check);
            string[] strSpliteRanges = Regex.Split(r.Text, "\r");
            int cout = 0;

            //System.Windows.Forms.MessageBox.Show(strSpliteRanges.Length + " ^^");
            foreach (string strSpliteRange in strSpliteRanges)
            {
                if (strSpliteRange.Length == 0)
                {
                    break;
                }
                int value = 0;
                if (CheckTypeLanguage(strSpliteRange))
                {
                    value = ModelBookTypeBookEN(r, strSpliteRange, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    value = ModelBookTypeBookTH(r, strSpliteRange, cout);
                }
                if (value == 0)
                {
                    System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }
                cout += value;
                r = r.Application.ActiveDocument.Range(cout);
               // l.sentence = strSpliteRange;
                //l.ForNames();
              //  bool v1 = l.ForNames();
              //  bool v2 = l.ForYear();
             //   bool v3 = l.ForBookName();
             //   bool v4 = l.ForPlaceEnd();
               // System.Windows.Forms.MessageBox.Show(v1 + " " + v2 + " " + (v3&&v4) + " ^^");
               /* if (strSpliteRange.Length == 0)
                {
                    break;
                }
                int value = ModelBookTypeBookTH(r, strSpliteRange, cout);// ModelBookTypeBookTH(r, cout);
                if (value == 0)
                {
                    break;
                }
                cout += value;
                r = r.Application.ActiveDocument.Range(cout);*/
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

        //ตรวจสอบภาษา
        private bool CheckTypeLanguage(string strCheck)
        {
            for (char chars = 'A'; chars <= 'Z'; chars++)
            {
                if (strCheck[0] == chars)
                {
                    return true;
                }
            }
            for (char chars = 'a'; chars <= 'b'; chars++)
            {
                if (strCheck[0] == chars)
                {
                    return true;
                }
            }
            return false;
        }
           //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (!l.ForBookTranslator())
                        {
                            return ModelBookTypeArticleTH(r, strCheck, cout);
                        }

                        if (l.ForPlaceEnd())
                        {
                            if (!l.ForBookAddEnd())
                            {
                                return ModelBookTypeArticleTH(r, strCheck, cout);
                            }
                            System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                            return l.countLength;
                        }
                        
                    }
                }
            }
            return ModelBookTypeArticleTH(r, strCheck, cout);
        }
          //หนังสือทั่วไป เอกสารประเภทหนังสือ
        /*private int ModelBookTypeBookTH(Word.Range r,string strCheck ,int cout)
        {
            string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\.\s\([1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s(\((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+[ผ][ู][้][แ][ป][ล]\)\.\s)?(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\.(\s)?(\((([0-9ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\))?)";
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
            
        }*/
        
        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int ModelBookTypeArticleTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameIn())
                        {
                            if (l.ForPage())
                            {
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeEncyclopediaTH(r, strCheck, cout);
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
        /*private int ModelBookTypeArticleTH(Word.Range r,string strCheck, int cout)
        {
            
            int checkValue = -1;
            int checkValue2 = -1;
            string []strCheckSplites = Regex.Split(strCheck,". ใน ");
            if (strCheckSplites.Length == 2)
            {
                string strFirst = strCheckSplites[0]+". ใน ";
                string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s[ใ][น]\s)";
                var task = Task.Factory.StartNew(() => CheckStringMatch(strFirst, regex, ref checkValue) );
                var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

                string strSecond = strCheckSplites[1];
                string regex2 = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([บ][ร][ร][ณ][า][ธ][ิ][ก][า][ร]\)\,\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+";
                var task2 = Task.Factory.StartNew(() => CheckStringMatch(strSecond, regex2, ref checkValue2));
                var completedWithinAllotedTime2 = task2.Wait(TimeSpan.FromMilliseconds(1000));
            }

            if (checkValue != -1 && checkValue2 != -1)
            {
                checkValue += checkValue2;
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
            return ModelBookTypeEncyclopediaTH(r, strCheck, cout);
        }*/

        //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        /*private int ModelBookTypeEncyclopediaTH(Word.Range r, string strCheck, int cout)
        {
            int checkValue = -1;
            int checkValue2 = -1;
            string []strCheckSplites = Regex.Split(strCheck,". ใน ");
            if (strCheckSplites.Length == 2)
            {
                string strFirst = strCheckSplites[0]+". ใน ";
                string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s[ใ][น]\s)";
                var task = Task.Factory.StartNew(() => CheckStringMatch(strFirst, regex, ref checkValue) );
                var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

                string strSecond = strCheckSplites[1];
                string regex2 = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+";
                var task2 = Task.Factory.StartNew(() => CheckStringMatch(strSecond, regex2, ref checkValue2));
                var completedWithinAllotedTime2 = task2.Wait(TimeSpan.FromMilliseconds(1000));
            }


            if (checkValue != -1 && checkValue2 != -1)
            {
                checkValue += checkValue2;
                System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {
                    
                    if (rngWord.Text == "). " || rngWord.Text == "ใน ")
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
                        System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ จบ");
                        cout = +checkValue;
                        return cout;
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH(r, strCheck,cout);

        }*/
         //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int ModelBookTypeEncyclopediaTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameInDotEditor())
                        {
                            if (l.ForPageAndBook())
                            {
                                if (l.ForPlaceEnd())
                                {

                                    System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                    
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH( r,  strCheck,  cout);
        }


        //เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ
        /*private int ModelBookTypeHandoutLibraryTH(Word.Range r, string strCheck, int cout)
        {

            int checkValue = -1;
            int checkValue2 = -1;
            string[] strCheckSplites = Regex.Split(strCheck, ". ใน ");
            if (strCheckSplites.Length == 2)
            {
                string strFirst = strCheckSplites[0] + ". ใน ";
                string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([ผ][ู][้][บ][ร][ร][ย][า][ย]\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s[ใ][น]\s)";
                var task = Task.Factory.StartNew(() => CheckStringMatch(strFirst, regex, ref checkValue));
                var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

                string strSecond = strCheckSplites[1];
                string regex2 = @"(([0-9ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+";
                var task2 = Task.Factory.StartNew(() => CheckStringMatch(strSecond, regex2, ref checkValue2));
                var completedWithinAllotedTime2 = task2.Wait(TimeSpan.FromMilliseconds(1000));
            }


            if (checkValue != -1 && checkValue2 != -1)
            {
                checkValue += checkValue2;
                System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {

                    if (rngWord.Text == "). " || rngWord.Text == "ใน ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 3)
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
                    else if (countCheck == 5)
                    {
                        System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ จบ");
                        cout = +checkValue;
                        return cout;
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH2(r, strCheck, cout);
        }*/

         //เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNamesNF())
            {
                if (l.ForNarrator())
                {
                    if (l.ForDate())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForBookNameInDotEditor())
                            {
                                if (l.ForPageAndBook())
                                {
                                    if (l.ForPlaceEnd())
                                    {
                                        System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH2(r, strCheck, cout);
        }

        //เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ
        /*private int ModelBookTypeHandoutLibraryTH2(Word.Range r, string strCheck, int cout)
        {

            int checkValue = -1;
            string regex = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\((([ผ][ู][้][บ][ร][ร][ย][า][ย])|([ผ][ู][้][ป][า][ฐ][ก][ถ][า]))\)\.\s\((([1-9])|([1-3][0-9]))((\s)?\-(\s)?([1-9])|([1-3][0-9]))?\s(([ม][ก][ร][า][ค][ม])|([ก][ุ][ภ][า][พ][ั][น][ธ][์])|([ม][ี][น][า][ค][ม])|([เ][ม][ษ][า][ย][น])|([พ][ฤ][ษ][พ][า][ค][ม])|([ม][ิ][ถ][ุ][น][า][ย][น])|([ก][ร][ก][ฎ][า][ค][ม])|([ส][ิ][ง][ห][า][ค][ม])|([ก][ั][น][ย][า][ย][น])|([ต][ุ][ล][า][ค][ม])|([พ][ฤ][ศ][จ][ิ][ก][า][ย][น])|([ธ][ั][น][ว][า][ค][ม]))\s[1-9][0-9]{3}\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s))?)+\.";
            var task = Task.Factory.StartNew(() => CheckStringMatch(strCheck, regex, ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {

                    if (rngWord.Text == "). ")
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2 || countCheck == 3)
                    {

                        if (rngWord.Bold != 0)
                        {
                            if (countCheck == 3)
                            {
                                countCheck = 2;
                            }
                            //System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");
                            if (rngWord.Text[rngWord.Text.Length - 2] == '.')
                            {
                                countCheck++;
                            }

                        }
                        else if (countCheck == 3)
                        {
                            System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ จบ");
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
            return ModelJournalTypeArticlesTH( r,  strCheck,  cout);
        }*/

         //เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryTH2(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNamesNF())
            {
                if (l.ForNarrator())
                {
                    if (l.ForDate())
                    {
                        if (l.ForBookName())
                        {

                            if (l.ForPlaceEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
                        }   
                    }
                }
            }
            return ModelJournalTypeArticlesTH(r,strCheck,cout);
        }

        //บทความทั่วไป เอกสารประเภทวารสาร
        /*private int ModelJournalTypeArticlesTH(Word.Range r, string strCheck, int cout)
        {

            int checkValue = -1;
            string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)|(\“([ก-ฮะ-์])+(ฯ)?\”\s))+((\?\s)|(\.\s))((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)|(\“([ก-ฮะ-์])+(ฯ)?\”\s))+\,\s[1-9]([0-9])*\([1-9]([0-9])*\)\,\s([1-9]([0-9])*)((\s)?\-(\s)?([1-9]([0-9])*))?\.)";
            var task = Task.Factory.StartNew(() => CheckStringMatch(strCheck, regex, ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                int countCheck2 = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {

                    if (countCheck2 == 0 && (rngWord.Text == "). " || rngWord.Text == ". " || rngWord.Text == "? "))
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2 || countCheck == 3)
                    {

                        if (rngWord.Bold != 0)
                        {
                            countCheck2 = 1;
                            if (countCheck == 3)
                            {
                                countCheck = 2;
                            }
                            //System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");
                            if (rngWord.Text == ", ")
                            {
                                countCheck++;
                            }

                        }
                        else if (countCheck == 3)
                        {
                            System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร จบ");
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
            return 0;
        }*/

        //บทความทั่วไป เอกสารประเภทวารสาร
        private int ModelJournalTypeArticlesTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeReviewTH(r, strCheck, cout);
        }

        //บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร
        /*private int ModelJournalTypeReviewTH(Word.Range r, string strCheck, int cout)
        {

            int checkValue = -1;
            string regex = @"^((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\.\s\([1-9][0-9]{3}\)\.\s((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)|(\“([ก-ฮะ-์])+(ฯ)?\”\s))+((\?\s)|(\.\s))((([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)|(\“([ก-ฮะ-์])+(ฯ)?\”\s))+\,\s[1-9]([0-9])*\([1-9]([0-9])*\)\,\s([1-9]([0-9])*)((\s)?\-(\s)?([1-9]([0-9])*))?\.)";
            var task = Task.Factory.StartNew(() => CheckStringMatch(strCheck, regex, ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                Word.Range rCheck = r.Application.ActiveDocument.Range(cout, cout + checkValue);
                int countCheck = 0;
                int countCheck2 = 0;
                foreach (Microsoft.Office.Interop.Word.Range rngWord in rCheck.Words)
                {

                    if (countCheck2 == 0 && (rngWord.Text == "). " || rngWord.Text == ". " || rngWord.Text == "? "))
                    {
                        countCheck++;
                        //   System.Windows.Forms.MessageBox.Show(range2.Text);
                    }
                    else if (countCheck == 2 || countCheck == 3)
                    {

                        if (rngWord.Bold != 0)
                        {
                            countCheck2 = 1;
                            if (countCheck == 3)
                            {
                                countCheck = 2;
                            }
                            //System.Windows.Forms.MessageBox.Show(rngWord.Text + "_");
                            if (rngWord.Text == ", ")
                            {
                                countCheck++;
                            }

                        }
                        else if (countCheck == 3)
                        {
                            System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร จบ");
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
            return 0;
        }*/

         //บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร
        private int ModelJournalTypeReviewTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameReview(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeInterviewTH(r, strCheck, cout);
        }

        //บทสัมภาษณ์ เอกสารประเภทวารสาร
        private int ModelJournalTypeInterviewTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForNamesInterviewer(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทสัมภาษณ์ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeBookTH(r, strCheck, cout);
        }

        //หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeBookTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeColumnTH(r, strCheck, cout);
        }

        //กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeColumnTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForColumnEnd())
                    {
                    }
                    if (l.ForBookNameEC())
                    {
                        if (l.ForPageEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                            return l.countLength;
                        }
                    }

                }
            }
            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForBookName())
            {
                if (l.ForDate())
                {
                    if (l.ForColumnEnd())
                    {
                    }
                    if (l.ForBookNameEC())
                    {
                        if (l.ForPageEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                            return l.countLength;
                        }
                    }

                }
            }
            return ModelNewspaperTypeInterviewTH(r, strCheck, cout);
        }

        //กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeInterviewTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForNamesInterviewer(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForPageEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForBookNameEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisAbstractBookTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ
        private int ModelThesisTypeThesisAbstractBookTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameToIn())
                    {
                        if (l.ForBookNameInDotEditor())
                        {
                            if (l.ForPage())
                            {
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ");
                                    return l.countLength;
                                }
                            }
                        }
                    }

                }
            }
            return ModelThesisTypeThesisAbstractJournalTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร
        private int ModelThesisTypeThesisAbstractJournalTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelThesisTypeThesisOnlineTH(r, strCheck, cout);
        }

        //ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisOnlineTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForBookName())
                                {
                                    if (l.ForSearch())
                                    {
                                        if (l.ForURL())
                                        {
                                            System.Windows.Forms.MessageBox.Show("ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                            return l.countLength;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeLetterTH(r, strCheck, cout);
        }


        //จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeLetterTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;

            if (l.ForNames())
            {
                if (l.ForMonthYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if(l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeBrochuresAndLeafletsTH(r, strCheck, cout);
        }

        //จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeBrochuresAndLeafletsTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;

            if (l.ForBookNameToBracket())
            {
                if (l.ForYear())
                {
                    if (l.ForBrochuresAndLeaflets())
                    {
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
            }
            return ModelOtherTypeArchivesTH(r, strCheck, cout);
        }

        //จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeArchivesTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;

            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    string sentenceCopy = l.sentence;
                    int countLengthCopy = l.countLength;
                    if (l.ForColumnEnd())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForPageEnd())
                            {
                                return l.countLength;
                            }
                        }
                        else
                        {
                            if (l.ForBookNameEnd())
                            {
                                return l.countLength;
                            }
                        }
                    }
                    else
                    {
                        l.sentence = sentenceCopy;
                        l.countLength = countLengthCopy;
                        if (l.ForBookNameEC())
                        {
                            if (l.ForPageEnd2())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                        else
                        {
                            l.sentence = sentenceCopy;
                            l.countLength = countLengthCopy;
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }

                }
            }
            return ModelOtherTypeGovernmentGazetteTH(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;

            if (l.ForBookNameToBracket())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForAt())
                        {
                            if (l.ForPageEnd2())
                            {
                                System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeAudioTH(r, strCheck, cout);
        }

        //สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeAudioTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            bool check = false;
            if (l.ForNameOnePrevious())
            {
                if (l.ForNameYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForBookNameReview(0))
                        {
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            l.sentence = strCheck;
            check = true;
            if (check)
            {
                if (l.ForBookNameES())
                {
                    if (l.ForBookNameReview(0))
                    {
                        if (l.ForNameYear())
                        {
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeImageTH(r, strCheck, cout);
        }

        //ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeImageTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForBookNameReview(0))
                        {

                        }
                    }
                }

            }
            return ModelOnlineTypeOnlineTH(r, strCheck, cout);
        }

        //บทความออนไลน์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOnlineTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {

                        if (l.ForSearch())
                        {
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            return ModelOnlineTypeOtherTH(r, strCheck, cout);
        }

        //บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOtherTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                bool check = false;
                string sentenceCopy = l.sentence;
                int countLengthCopy = l.countLength;
                if (l.ForDate())
                {
                    check = true;
                }
                else
                {
                    l.sentence = sentenceCopy;
                    l.countLength = countLengthCopy;
                    if (l.ForMonthYear())
                    {
                        check = true;
                    }
                }
                if (check)
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookName())
                        {
                        }
                        if (l.ForSearch())
                        {
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            else
            {
                if (l.ForBookName())
                {
                    bool check = false;
                    string sentenceCopy = l.sentence;
                    int countLengthCopy = l.countLength;
                    if (l.ForDate())
                    {
                        check = true;
                    }
                    else
                    {
                        l.sentence = sentenceCopy;
                        l.countLength = countLengthCopy;
                        if (l.ForMonthYear())
                        {
                            check = true;
                        }
                    }
                    if (check)
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForBookName())
                            {
                            }
                                if (l.ForSearch())
                                {
                                    if (l.ForURL())
                                    {
                                        System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                        return l.countLength;
                                    }
                                }
                            
                        }
                    }
                }
            }
            return ModelOnlineTypeElectronicTH(r, strCheck, cout);
        }

        //บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeElectronicTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForNameYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForSearch())
                            {
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            return 0;
        }

//===========================================================================================================================//
//===========================================================================================================================//
        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                            return l.countLength;
                        }

                    }
                }
            }
            return ModelBookTypeArticleEN(r, strCheck, cout);
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int ModelBookTypeArticleEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameIn())
                        {
                            if (l.ForPage())
                            {
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeEncyclopediaEN(r, strCheck, cout);
        }

        //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int ModelBookTypeEncyclopediaEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameInDotEditor())
                        {
                            if (l.ForPage())
                            {
                                if (l.ForPlaceEnd())
                                {

                                    System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                    return l.countLength;

                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryEN(r, strCheck, cout);
        }

        //เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNamesNF())
            {
                if (l.ForNarrator())
                {
                    if (l.ForDate())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForBookNameInDotEditor())
                            {
                                if (l.ForPage())
                                {
                                    if (l.ForPlaceEnd())
                                    {
                                        System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryEN2(r, strCheck, cout);
        }

        //เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryEN2(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNamesNF())
            {
                if (l.ForNarrator())
                {
                    if (l.ForDate())
                    {
                        if (l.ForBookName())
                        {

                            if (l.ForPlaceEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelJournalTypeArticlesEN(r, strCheck, cout);
        }

        //บทความทั่วไป เอกสารประเภทวารสาร
        private int ModelJournalTypeArticlesEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeReviewEN(r, strCheck, cout);
        }


        //บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร
        private int ModelJournalTypeReviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameReview(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeInterviewEN(r, strCheck, cout);
        }

        //บทสัมภาษณ์ เอกสารประเภทวารสาร
        private int ModelJournalTypeInterviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForNamesInterviewer(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทสัมภาษณ์ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeBookEN(r, strCheck, cout);
        }

        //หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeColumnEN(r, strCheck, cout);
        }

        //กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeColumnEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForColumnEnd())
                    {
                    }
                    if (l.ForBookNameEC())
                    {
                        if (l.ForPageEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                            return l.countLength;
                        }
                    }

                }
            }
            else
            {
                l.sentence = strCheck;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    if (l.ForDate())
                    {
                        if (l.ForColumnEnd())
                        {
                        }
                        if (l.ForBookNameEC())
                        {
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            return ModelNewspaperTypeInterviewEN(r, strCheck, cout);
        }

        //กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeInterviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForNamesInterviewer(0))
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForPageEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameInitials())
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForBookNameEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisAbstractBookEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ
        private int ModelThesisTypeThesisAbstractBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameToIn())
                    {
                        if (l.ForBookNameInDotEditor())
                        {
                            if (l.ForPage())
                            {
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ");
                                    return l.countLength;
                                }
                            }
                        }
                    }

                }
            }
            return ModelThesisTypeThesisAbstractJournalEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร
        private int ModelThesisTypeThesisAbstractJournalEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelThesisTypeThesisOnlineEN(r, strCheck, cout);
        }

        //ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisOnlineEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForBookNameEC())
                            {
                                if (l.ForBookName())
                                {
                                    if (l.ForSearch())
                                    {
                                        if (l.ForURL())
                                        {
                                            System.Windows.Forms.MessageBox.Show("ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                            return l.countLength;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeLetterEN(r, strCheck, cout);
        }


        //จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeLetterEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;

            if (l.ForNames())
            {
                if (l.ForMonthYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookNameEC())
                        {
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeBrochuresAndLeafletsEN(r, strCheck, cout);
        }

        //จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeBrochuresAndLeafletsEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;

            if (l.ForBookNameToBracket())
            {
                if (l.ForYear())
                {
                    if (l.ForBrochuresAndLeaflets())
                    {
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
            }
            return ModelOtherTypeArchivesEN(r, strCheck, cout);
        }

        //จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeArchivesEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;

            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    string sentenceCopy = l.sentence;
                    int countLengthCopy = l.countLength;
                    if (l.ForColumnEnd())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForPageEnd())
                            {
                                return l.countLength;
                            }
                        }
                        else
                        {
                            if (l.ForBookNameEnd())
                            {
                                return l.countLength;
                            }
                        }
                    }
                    else
                    {
                        l.sentence = sentenceCopy;
                        l.countLength = countLengthCopy;
                        if (l.ForBookNameEC())
                        {
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                        else
                        {
                            l.sentence = sentenceCopy;
                            l.countLength = countLengthCopy;
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }

                }
            }
            return ModelOtherTypeGovernmentGazetteEN(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;

            if (l.ForBookNameToBracket())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForAt())
                        {
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeAudioEN(r, strCheck, cout);
        }

        //สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeAudioEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            bool check = false;
            if (l.ForNameOnePrevious())
            {
                if (l.ForNameYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForBookNameReview(0))
                        {
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            l.sentence = strCheck;
            check = true;
            if (check)
            {
                if (l.ForBookNameES())
                {
                    if (l.ForBookNameReview(0))
                    {
                        if (l.ForNameYear())
                        {
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeImageEN(r, strCheck, cout);
        }

        //ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeImageEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForYear())
                {
                    if (l.ForBookNameES())
                    {
                        if (l.ForBookNameReview(0))
                        {

                        }
                    }
                }

            }
            return ModelOnlineTypeOnlineEN(r, strCheck, cout);
        }

        //บทความออนไลน์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOnlineEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForDate())
                {
                    if (l.ForBookName())
                    {

                        if (l.ForSearch())
                        {
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            else
            {
                if (l.ForBookName())
                {
                    if (l.ForDate())
                    {
                        if (l.ForBookName())
                        {

                            if (l.ForSearch())
                            {
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }

                        }
                    }
                }
            }
            return ModelOnlineTypeOtherEN(r, strCheck, cout);
        }

        //บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOtherEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                bool check = false;
                string sentenceCopy = l.sentence;
                int countLengthCopy = l.countLength;
                if (l.ForDate())
                {
                    check = true;
                }
                else
                {
                    l.sentence = sentenceCopy;
                    l.countLength = countLengthCopy;
                    if (l.ForMonthYear())
                    {
                        check = true;
                    }
                }
                if (check)
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookName())
                        {
                        }
                        if (l.ForSearch())
                        {
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            else
            {
                if (l.ForBookName())
                {
                    bool check = false;
                    string sentenceCopy = l.sentence;
                    int countLengthCopy = l.countLength;
                    if (l.ForDate())
                    {
                        check = true;
                    }
                    else
                    {
                        l.sentence = sentenceCopy;
                        l.countLength = countLengthCopy;
                        if (l.ForMonthYear())
                        {
                            check = true;
                        }
                    }
                    if (check)
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForBookName())
                            {
                            }
                            if (l.ForSearch())
                            {
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }

                        }
                    }
                }
            }
            return ModelOnlineTypeElectronicEN(r, strCheck, cout);
        }

        //บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeElectronicEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            if (l.ForNames())
            {
                if (l.ForNameYear())
                {
                    if (l.ForBookName())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForSearch())
                            {
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            else
            {
                l.sentence = strCheck;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    if (l.ForNameYear())
                    {
                        if (l.ForBookName())
                        {
                            if (l.ForBookName())
                            {
                                if (l.ForSearch())
                                {
                                    if (l.ForURL())
                                    {
                                        System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                        return l.countLength;
                                    }
                                }
                            }

                        }
                    }
                }
            }
            return 0;
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
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

//===========================================================================================================================//
//===========================================================================================================================//

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
