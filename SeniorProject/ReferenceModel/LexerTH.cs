using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
namespace SeniorProject
{
    class LexerTH
    {
        public string sentence;
        public int countLength;
        public Word.Range range;
        public int countCutBold = 0;
        public int countCutNotBold = 0;
        string []listInitialsTHs;
        string[] monthTHs;
        int checkForBookName;
        bool checkC;//วงเล็บห
        public LexerTH()
        {
            checkC = false;
            checkForBookName = 0;
            countLength = 0;
            listInitialsTHs = new string[26];
            listInitialsTHs[0] = @"^([เ][อ])";
            listInitialsTHs[1] = @"^([บ][ี])";
            listInitialsTHs[2] = @"^([ซ][ี])";
            listInitialsTHs[3] = @"^([ด][ี])";
            listInitialsTHs[4] = @"^([อ][ี])";
            listInitialsTHs[5] = @"^([เ][อ][ฟ])";
            listInitialsTHs[6] = @"^([จ][ี])";
            listInitialsTHs[7] = @"^([เ][อ][ช])";
            listInitialsTHs[8] = @"^([ไ][อ])";
            listInitialsTHs[9] = @"^([เ][จ])";
            listInitialsTHs[10] = @"^([เ][ค])";
            listInitialsTHs[11] = @"^([แ][อ][ล])";
            listInitialsTHs[12] = @"^([เ][อ][็][ม])";
            listInitialsTHs[13] = @"^([เ][อ][็][น])";
            listInitialsTHs[14] = @"^([โ][อ])";
            listInitialsTHs[15] = @"^([พ][ี])";
            listInitialsTHs[16] = @"^([ค][ิ][ว])";
            listInitialsTHs[17] = @"^([อ][า][ร][์])";
            listInitialsTHs[18] = @"^([เ][อ][ส])";
            listInitialsTHs[19] = @"^([ท][ี])";
            listInitialsTHs[20] = @"^([ย][ู])";
            listInitialsTHs[21] = @"^([ว][ี])";
            listInitialsTHs[22] = @"^([ด][ั][บ][เ][บ][ิ][้][ล][ย][ู])";
            listInitialsTHs[23] = @"^([เ][อ][็][ก][ซ][์])";
            listInitialsTHs[24] = @"^([ว][า][ย])";
            listInitialsTHs[25] = @"^([แ][ซ][ด])";

            monthTHs = new string[12];
            monthTHs[0] = @"^[ม][ก][ร][า][ค][ม]";
            monthTHs[1] = @"^[ก][ุ][ม][ภ][า][พ][ั][น][ธ][์]";
            monthTHs[2] = @"^[ม][ี][น][า][ค][ม]";
            monthTHs[3] = @"^[เ][ม][ษ][า][ย][น]";
            monthTHs[4] = @"^[พ][ฤ][ษ][ภ][า][ค][ม]";
            monthTHs[5] = @"^[ม][ิ][ถ][ุ][น][า][ย][น]";
            monthTHs[6] = @"^[ก][ร][ก][ฎ][า][ค][ม]";
            monthTHs[7] = @"^[ส][ิ][ง][ห][า][ค][ม]";
            monthTHs[8] = @"^[ก][ั][น][ย][า][ย][น]";
            monthTHs[9] = @"^[ต][ุ][ล][า][ค][ม]";
            monthTHs[10] = @"^[พ][ฤ][ศ][จ][ิ][ก][า][ย][น]";
            monthTHs[11] = @"^[ธ][ั][น][ว][า][ค][ม]";
        }

        private void CheckStringMatch(string strFromRange, string regex, ref int checkValue)
        {
            Match match = Regex.Match(strFromRange, regex);
            if (match.Success)
            {
                checkValue = match.Value.Length;
                return;
            }
            checkValue = -1;
        }

        public int ForName()
        {
            int valueName2 = ForName2();
            if (valueName2 != -1)
            {
                return valueName2;
            }
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+",ref checkValue);
            if (checkValue!=-1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 3;
                    }
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return 2;
                                }
                                return 1;
                            }
                            return 2;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2()
        {
            string sentenceCoppy = this.sentence;
             int checkValue = -1;

             int valueForNamelistInitialsTH = ForNamelistInitialsTH();
             if (valueForNamelistInitialsTH == 0)
             {
                 return -1;
             }
             else if (valueForNamelistInitialsTH >= 9999)
             {
                 this.sentence = sentenceCoppy;
                 return -1;
             }
             else
             { 
                 CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+",ref checkValue);
                 if (checkValue != -1)
                 {
                      CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            return 2;
                        }
                     return 1;
                 }
                 return -1;
             }

        }

        public int ForNameDontAnd()
        {
            int valueName2 = ForName2DontAnd();
            if (valueName2 != -1)
            {
                return valueName2;
            }

            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }


            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 8888;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return -1;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return -1;
                                }
                                return 1;
                            }
                            return -1;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2DontAnd()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -1;
                        }
                        return -1;
                    }
                    return 1;
                }
                return -1;
            }

        }
        
        public bool ForNames()
        {
            int checkValue = -1;
            int checkValueFormate = ForName();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAnd();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNames();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue!=-1)
            {
                CutString(checkValue);

                return ForNames();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }

           /* CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                return true;
            }*/

            CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNames();
            }

            if (checkValueFormate == 3)
            {
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }

            
            return false;
        }

        public bool ForNamesForCheck()
        {
            int checkValue = -1;
            int checkValueFormate = ForName();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAnd();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNamesForCheck();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return ForNamesForCheck();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return true;
            }

            /* CheckStringMatch(this.sentence, @"^\(", ref checkValue);
             if (checkValue != -1)
             {
                 return true;
             }*/

            CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNamesForCheck();
            }

            if (checkValueFormate == 3)
            {
                return true;
            }

            return false;
        }

        public int ForNamelistInitialsTH()
        {
            Match match;
            foreach (string listInitialsTH in listInitialsTHs)
            {
                match = Regex.Match(this.sentence, listInitialsTH);
                if (match.Success)
                {
                    string listInitialsTHNew = listInitialsTH + @"\.\s";
                    match = Regex.Match(this.sentence, listInitialsTHNew);
                    if (match.Success)
                    {
                        CutString(match.Length);
                        return ForNamelistInitialsTH()+1;
                    }

                    listInitialsTHNew = listInitialsTH + @"\.\,\s";
                    match = Regex.Match(this.sentence, listInitialsTHNew);
                    if (match.Success)
                    {
                        CutString(match.Length);
                        return 9999;
                    }
                }
            }


            return 0;
        }

        public bool ForYear()
        {
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9][0-9]{3})+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([ก-ฮ])*", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([ก-ฮ].)+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                    }
                }
            }
            return false;
        }

        public bool ForYearForCheck()
        {
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9][0-9]{3})+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([ก-ฮ])*", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([ก-ฮ].)+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public bool ForYearCreate()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }
            }
            this.sentence = sentenceCopy;
            this.countLength = countLengthCopy;
            return false;
        }

        public bool ForBookName()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookName();
                        }
                    }
                    return false;
                }
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookName();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([A-Za-z๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookName();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookName();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }
            
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookName();
            }
            return false;
        }

        public bool ForBookNameBold()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameBold();
                        }
                    }
                    return false;
                }
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength - 2));
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameBold();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([A-Za-z๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameBold();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameBold();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return  CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold,this.countLength-2));
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameBold();
            }
            return false;
        }

        public bool ForBookNameEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameEnd();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            if (!checkC)
            {
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkC = true;
                }
            }

            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEnd();
                }

                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEnd();
                }

                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameEnd();
            }

            
            if (checkC)
            {
                CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                else
                {
                    return false;
                }
            }


            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence == "")
                {
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
                return ForBookNameEnd();
            }

            CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence != "")
                {
                    return ForBookNameEnd();
                }
            }
            return false;
        }

        public bool ForBookNameEndBold()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameEndBold();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            if (!checkC)
            {
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkC = true;
                }
            }

            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEndBold();
                }

                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEndBold();
                }

                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameEndBold();
            }


            if (checkC)
            {
                CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                else
                {
                    return false;
                }
            }


            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence == "")
                {
                    return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength-1));
                }
                return ForBookNameEndBold();
            }

            CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence != "")
                {
                    return ForBookNameEndBold();
                }
            }
            return false;
        }

        public bool ForPlaceEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForPlaceEnd();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForPlaceEnd();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForPlaceEnd();
            }
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1 && this.checkForBookName > 0)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }
            return false;
        }

        public bool ForBookTranslator()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                while (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);

                        CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }

                        CheckStringMatch(this.sentence, @"^(\,\s)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public bool ForBookAddEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\s", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                            }

                            CheckStringMatch(this.sentence, @"^(\,\s)+", ref checkValue);
                            if (checkValue != -1)
                            {
                                continue;
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public bool ForBookNameIn()
        {
            int countCutNotBoldIn = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([ใ][น]\s)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                this.countCutNotBold = this.countLength;
                if (this.ForNamesNF())
                {
                    CheckStringMatch(this.sentence, @"^(\()+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([บ][ร][ร][ณ][า][ธ][ิ][ก][า][ร])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\)\,\s)+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);

                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }

                if (!CheckNotBold(this.range.Application.ActiveDocument.Range(countCutNotBoldIn, this.countLength)))
                {
                    return false;
                }
                this.countCutBold = this.countLength;
                if (ForBookNameNFBold())
                {
                        return true;
                }
                
            }

            return false;
        }

        public int ForNameNF()
        {
            int valueName2 = ForName2NF();
            if (valueName2 != -1)
            {
                return valueName2;
            }
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 3;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return 1;
                                }
                                return 1;
                            }
                            return 1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2NF()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        return 1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public int ForNameDontAndNF()
        {
            int valueName2 = ForName2DontAndNF();
            if (valueName2 != -1)
            {
                return valueName2;
            }

            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }


            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 8888;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return -1;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return 1;
                                }
                                return 1;
                            }
                            return 1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2DontAndNF()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public bool ForNamesNF()
        {
            int checkValue = -1;
            int checkValueFormate = ForNameNF();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAndNF();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue != -1)
                    {
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNamesNF();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAndNF();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                        if (checkValue != -1)
                        {
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return ForNamesNF();
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {

                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }

            CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNamesNF();
            }

            if (checkValueFormate == 3)
            {
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }


            return false;
        }

        public bool ForBookNameNF()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameNF();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameNF();
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameNF();
            }
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^\,\s\(", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue-1);
                this.checkForBookName = 0;
                return true;
            }

            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameNF();
            }

            return false;
        }

        public bool ForBookNameNFBold()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameNFBold();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameNFBold();
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameNFBold();
            }
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength - 1));
            }
            CheckStringMatch(this.sentence, @"^\,\s\(", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength-1));
            }

            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameNFBold();
            }

            return false;
        }

        public bool ForPage()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            
            if (checkValue != -1)
            {
                 CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([ห][น][้][า]\s)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                }
                else
                {
                    return false;
                }

                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }

            }
            return false;
        }
 
        public bool ForBookNameInDotEditor()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([ใ][น]\s)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (!CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength)))
                {
                    return false;
                }
                this.countCutBold = this.countLength;
                if (ForBookNameNFBold())
                {
                    return true;
                }

            }

            return false;
        }

        public bool ForPageAndBook()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                string sentenceCoppy = this.sentence;
                int countLengthCoppy = this.countLength;
                if (!ForBookNameEC())
                {
                    this.sentence = sentenceCoppy;
                    this.countLength = countLengthCoppy;
                }
                
                CheckStringMatch(this.sentence, @"^([ห][น][้][า]\s)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                }
                else
                {
                    return false;
                }

                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }

            }
            return false;
        }

        public bool ForBookNameEC()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameEC();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEC();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameEC();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameEC(); 
            }


            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }

            CheckStringMatch(this.sentence, @"^\.", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameEC();
            }

            return false;
        }

        public bool ForBookNameECBold()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameECBold();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameECBold();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameECBold();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameECBold();
            }


            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength - 2));
            }

            CheckStringMatch(this.sentence, @"^\.", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameECBold();
            }

            return false;
        }

        public bool ForNarrator()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(([ผ][ู][้][บ][ร][ร][ย][า][ย])|([ผ][ู][้][ป][า][ฐ][ก][ถ][า]))", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }

            }
            return false;
        }

        public bool ForDate()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValue = ForNameMonthTH();
                        if (checkValue > 0)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([ก-ฮ]\.)+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool ForDateForCheck()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValue = ForNameMonthTH();
                        if (checkValue > 0)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([ก-ฮ]\.)+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    return true;
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public int ForNameMonthTH()
        {
            Match match;
            foreach (string monthTH in monthTHs)
            {
                match = Regex.Match(this.sentence, monthTH);
                if (match.Success)
                {
                    return match.Length;
                }
            }


            return 0;
        }

        public bool ForYearAndNumber()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([ก-ฮ])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^\)\,\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                            if (checkValue != -1)
                            {

                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                                if (checkValue != -1)
                                {

                                    CutString(checkValue);
                                }
                                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                                if (checkValue != -1)
                                {

                                    CutString(checkValue);
                                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool ForAt()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([ก-ฮ])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));

                        }
                    }
                }
            }
            return false;
        }

        bool checkForBookNameReview = false;
        public bool ForBookNameReview(int checkSame)
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\[", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1 || checkSame > 0)
            {
                if (checkSame == 0)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);

                        CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }

                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        while (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);

                                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                }
                                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                }

                                CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return ForBookNameReview(1);
                            }
                        }
                        return false;
                    }
                }
                //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
                string sentenceCopyForSubject = this.sentence;
                int countForSubject = 0;
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์๐-๙0-9-/])+", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);
                    countForSubject += checkValue;
                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        countForSubject += checkValue;
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        countForSubject += checkValue;
                    }

                    CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return ForBookNameReview(1);
                    }
                    if (this.checkForBookName == 0)
                    {
                        CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                        if (checkValue != -1)
                        {
                            this.checkForBookName++;
                            CutString(checkValue);
                            countForSubject += checkValue;
                        }
                    }

                    
                    CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        if (!checkForBookNameReview)
                        {
                            countForSubject += checkValue;
                            string subject = sentenceCopyForSubject.Substring(0, countForSubject);
                            subject = subject.Substring(subject.Length - 7);
                            if (subject == "เรื่อง ")
                            {
                                if (!CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength)))
                                {
                                    return false;
                                }
                                this.countCutBold = this.countLength;
                                checkForBookNameReview = true;
                            }
                        }
                    }
                    return ForBookNameReview(1);
                }
                CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (!CheckNotBold(this.range.Application.ActiveDocument.Range(this.countLength-3, this.countLength)))
                    {
                        return false;
                    }
                    this.checkForBookName = 0;
                    return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength-3));
                }
            }
            return false;
        }

        public bool ForBookNameNotPublished(int checkSame)
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\[", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1 || checkSame > 0)
            {
                if (checkSame == 0)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);

                        CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }

                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        while (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);

                                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                }
                                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                }

                                CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return ForBookNameNotPublished(1);
                            }
                        }
                        return false;
                    }
                }

                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์๐-๙0-9-/])+", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);

                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);

                    }

                    CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return ForBookNameNotPublished(1);
                    }
                    if (this.checkForBookName == 0)
                    {
                        CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                        if (checkValue != -1)
                        {
                            this.checkForBookName++;
                            CutString(checkValue);

                        }
                    }


                    CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    return ForBookNameNotPublished(1);
                }
                CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);
                    this.checkForBookName = 0;
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
            }
            return false;
        }

        public int ForNameInterviewer()
        {
            int valueName2 = ForNameInterviewer2();
            if (valueName2 != -1)
            {
                return valueName2;
            }
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 3;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return 2;
                                }
                                return 1;
                            }
                            return 2;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForNameInterviewer2()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        return 2;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public int ForNameInterviewerDontAnd()
        {
            int valueName2 = ForNameInterviewer2DontAnd();
            if (valueName2 != -1)
            {
                return valueName2;
            }

            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }


            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 8888;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return -1;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return -1;
                                }
                                return 1;
                            }
                            return -1;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForNameInterviewer2DontAnd()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -1;
                        }
                        return -1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public bool ForNamesInterviewer(int checkSame)
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\[", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1 || checkSame > 0)
            {
                if (checkSame == 0)
                {
                    CutString(checkValue);
                }
                int checkValueFormate = ForNameInterviewer();
                if (checkValueFormate == -1)
                {
                    return false;
                }
                else if (checkValueFormate == 4)
                {
                    checkValueFormate = ForNameInterviewerDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^[ผ][ู][้][ส][ั][ม][ภ][า][ษ][ณ][์]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                }
                            }
                        }
                    }
                    else if (checkValueFormate == 8888)
                    {
                        CheckStringMatch(this.sentence, @"^[ผ][ู][้][ส][ั][ม][ภ][า][ษ][ณ][์]", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                            }
                        }
                    }
                    return false;
                }
                else if (checkValueFormate == 0 || checkValueFormate == 9999)
                {
                    return ForNamesInterviewer(1);
                }
                else if (checkValueFormate == 2)
                {
                    CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValueFormate = ForNameInterviewerDontAnd();
                        if (checkValueFormate == 1)
                        {
                            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                            if (checkValue != -1)
                            {

                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^[ผ][ู][้][ส][ั][ม][ภ][า][ษ][ณ][์]", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                    }
                                }
                            }
                        }
                        return false;
                    }
                }

                //Match match = Regex.Match(this.sentence, @"^\,\s");
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[ผ][ู][้][ส][ั][ม][ภ][า][ษ][ณ][์]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                    }
                    return ForNamesInterviewer(1);
                }

                CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForNamesInterviewer(1);
                }

            }
            return false;
        }

        public bool ForBookNameES()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameES();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameES();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameES();
            }
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }

            return false;
        }

        public bool ForBookNameESBold()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameESBold();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameESBold();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameESBold();
            }
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength));
            }

            return false;
        }

        public bool ForBookNameDB(int check)
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                check = 1;
            }
            else if (check == 1)
            {

            }
            else
            {
                return false;
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์A-Za-z๐-๙0-9-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                if (checkValue != -1)
                {
                    this.checkForBookName++;
                    CutString(checkValue);
                }


                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameDB(check);
            }
            CheckStringMatch(this.sentence, @"^\)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                if (this.checkForBookName > 0)
                {
                    this.checkForBookName = 0;
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
            }

            return false;
        }

        public bool ForPageEnd()
        {
            int checkValue = -1;

                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }

            
            return false;
        }

        public bool ForPageEnd2()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([ห][น][้][า]\s)", ref checkValue);
            if (checkValue != -1)
            {

                CutString(checkValue);
            }
            else
            {
                return false;
            }
            CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
            if (checkValue != -1)
            {

                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
            }


            return false;
        }

        public bool ForColumnEnd()
        {
            string sentenceCopy =  this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForColumnEnd();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForColumnEnd();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForColumnEnd();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForColumnEnd();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1 && this.checkForBookName > 0)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }
            this.sentence = sentenceCopy;
            this.countLength = countLengthCopy;
            return false;
        }

        public bool ForBookNameToIn()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameToIn();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToIn();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameToIn();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1)
            {
                
                CutString(checkValue);
                this.checkForBookName = 0;
                 CheckStringMatch(this.sentence, @"^[ใ][น]", ref checkValue);
                 if (checkValue != -1)
                 {
                     return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                 }
                 return ForBookNameToIn();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToIn();
            }
            return false;
        }

        public bool ForSearch(){

            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^[ส][ื][บ][ค][้][น][เ][ม][ื][่][อ]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            checkValue = ForNameMonthTH();
                            if (checkValue > 0)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool ForURL()
        {

            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^[จ][า][ก]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([0-9a-zA-z./:=])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                    }
                }
            }

            return false;
        }

        public bool ForMonthYear()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {

                CutString(checkValue);
                checkValue = ForNameMonthTH();
                if (checkValue > 0)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\-", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValue = ForNameMonthTH();
                        if (checkValue > 0)
                        {
                            CutString(checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                            }
                        }
                    }


                }
            }
            return false;
        }

        public bool ForBookNameToBracket()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameToBracket();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToBracket();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameToBracket();
            }
            CheckStringMatch(this.sentence, @"^(\.\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue-1);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
            }
            

            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracket();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracket();
            }

            return false;
        }

        public bool ForBookNameToBracketBold()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameToBracketBold();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToBracketBold();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameToBracketBold();
            }
            CheckStringMatch(this.sentence, @"^(\.\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return CheckBold(this.range.Application.ActiveDocument.Range(this.countCutBold, this.countLength-2));
            }


            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketBold();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketBold();
            }

            return false;
        }

        public bool ForBookNameToBracketForCheck()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameToBracketForCheck();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\?)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToBracketForCheck();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameToBracketForCheck();
            }
            CheckStringMatch(this.sentence, @"^(\.\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^(\?\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return true;
            }

            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketForCheck();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketForCheck();
            }

            return false;
        }

        public bool ForBrochuresAndLeaflets()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\[)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            bool checkPass = false;
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[แ][ผ][่][น][พ][ั][บ]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkPass = true;
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^[จ][ุ][ล][ส][า][ร]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkPass = true;
                    }
                }
            }
            if (checkPass)
            {
                CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
            }
            return false;
        }

        public bool ForNamePrevious()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    while (true)
                    {
                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^([,.])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        break;
                    }

                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                    CheckStringMatch(this.sentence, @"^\)\;\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                        if (checkValue != -1)
                        {
                            return false;
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        public bool ForNameYear()
        {
            string sentanceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^([ก-ฮ]\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            this.countLength = countLengthCopy;
            this.sentence = sentanceCopy;   
            return ForDate();
        }

        public bool ForNameYearForCheck()
        {
            string sentanceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^([ก-ฮ]\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            this.countLength = countLengthCopy;
            this.sentence = sentanceCopy;
            return ForDateForCheck();
        }

        public bool ForNameOne()
        {
            bool valueName2 = ForNameOne2();
            if (valueName2)
            {
                return valueName2;
            }
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^((([ก-ฮ])+\.)+)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
            }
            else
            {
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsTH = ForNamelistInitialsTH();
                    CheckStringMatch(this.sentence, @"^(\[[น][า][ม][แ][ฝ][ง]\]\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    if (valueForNamelistInitialsTH == 0)
                    {
                        return false;
                    }
                    else if (valueForNamelistInitialsTH >= 9999)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\[[น][า][ม][แ][ฝ][ง]\]\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                    if (checkValue != -1)
                    {
                        return false;
                    }
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\[[น][า][ม][แ][ฝ][ง]\]\s)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return true;
                            }
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return false;
                            }
                            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CheckStringMatch(this.sentence, @"^(\[[น][า][ม][แ][ฝ][ง]\]\s)", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return true;
                                    }
                                    return true;
                                }
                                return false;
                            }
                            return true;
                        }
                        return false;
                    }
                    return true;
                }
                return false;
            }
            return false;
        }

        public bool ForNameOne2()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsTH = ForNamelistInitialsTH();
            if (valueForNamelistInitialsTH == 0)
            {
                return false;
            }
            else if (valueForNamelistInitialsTH >= 9999)
            {
                this.sentence = sentenceCoppy;
                return false;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\[[น][า][ม][แ][ฝ][ง]\]\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                        return true;
                    }
                    return false;
                }
                return false;
            }

        }

        public bool ForNameOnePrevious()
        {
            if (ForNameOne())
            {
                if (ForNamePrevious())
                {
                    int checkValue = -1;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue == -1)
                    {

                        return ForNameOnePrevious();
                    }
                    return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                }
            }
            return false;
        }

        public bool ForNameOnePreviousForCheck()
        {
            if (ForNameOne())
            {
                if (ForNamePrevious())
                {
                    int checkValue = -1;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue == -1)
                    {

                        return ForNameOnePreviousForCheck();
                    }
                    return true;
                }
            }
            return false;
        }

        public bool ForBookNumber()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([เ][ล][ข][ท][ี][่])", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[1-9][0-9]*", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Application.ActiveDocument.Range(this.countCutNotBold, this.countLength));
                        }
                    }
                }
            }
            return false;
        }


        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
            //this.range = this.range.Application.ActiveDocument.Range(0, strLength);
        }

        bool CheckBold(Word.Range range)
        {
            if (range.Bold == -1)
            {
                return true;
            }
            return false;
        }
        bool CheckNotBold(Word.Range range)
        {
            if (range.Bold == 0)
            {
                return true;
            }
            return false;
        }
    }
}
