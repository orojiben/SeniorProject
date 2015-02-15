using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    class LexerECPE
    {
        public string sentence;
        public int countLength;
        private int numberCheckSort;
        private int numberMem;

        public LexerECPE()
        {
            numberMem = 0;
            numberCheckSort = 0;
        }

        private void CheckStringMatch(string strFromRange, string regex, ref int checkValue,bool getNumber)
        {
            Match match = Regex.Match(strFromRange, regex);
            if (match.Success)
            {
                checkValue = match.Value.Length;
                if (getNumber)
                {
                    numberMem = Int32.Parse(match.Value);
                }
                return;
            }
            checkValue = -1;
        }

        public int checkNumber()
        {
            int checkValue = -1;
            this.countLength = 0;
            CheckStringMatch(this.sentence, @"^\[", ref checkValue,false);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[1-9][0-9]*", ref checkValue,true);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\]\s", ref checkValue, false);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        if (CheckSort())
                        {
                            return this.countLength;
                        }
                    }
                }
            }
            return 0;
        }

        public bool CheckSort()
        {
            if ((this.numberMem - this.numberCheckSort) == 1)
            {
                this.numberCheckSort = this.numberMem;
                return true;
            }
            return false;
        }

        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
        }
    }
}
