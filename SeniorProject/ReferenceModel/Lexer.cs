using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeniorProject
{
    class Lexer
    {
        public string sentence;
        int countLength;
        string []listInitialsTHs;
        int checkLastName;
        public Lexer()
        {
            checkLastName = 0;
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
            }
            else
            {
                string sentenceCopy = this.sentence;
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    this.sentence = sentenceCopy;
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+",ref checkValue);
            if (checkValue!=-1)
            {
                CutString(checkValue);
                //this.sentence.Remove(0,10);
                string sentenceCopy = this.sentence;
                //match = Regex.Match(sentenceCopy, @"^\,\s");
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
                        return 1;
                    }
                }

               // match = Regex.Match(sentenceCopy, @"^\,\s");
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])+", ref checkValue);
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
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])+", ref checkValue);
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

        public int ForNameDontAnd()
        {
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
            }
            else
            {
                CheckStringMatch(this.sentence, @"^(([ก-ฮ]))+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^ฯ", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                }
            }

            CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                //match = Regex.Match(sentenceCopy, @"^\,\s");
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

                // match = Regex.Match(sentenceCopy, @"^\,\s");
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([แ][ล][ะ])+", ref checkValue);
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
                            CheckStringMatch(this.sentence, @"^([แ][ล][ะ])+", ref checkValue);
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
                    CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                    if (checkValue != -1)
                    {
                        return true;
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return true;
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
                    checkValueFormate = ForNameDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                        if (checkValue != -1)
                        {
                            return true;
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

                return true;
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                return true;
            }

            CheckStringMatch(this.sentence, @"^[แ][ล][ะ]", ref checkValue);
            if (checkValue != -1)
            {
                return ForNames();
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

        public int ForYear()
        {

            return 0;
        }

        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
        }
    }
}
