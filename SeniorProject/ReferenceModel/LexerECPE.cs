using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class LexerECPE
    {
        public string sentence;
        public int countLength;
        private int numberCheckSort;
        public int numberMem;
        public Word.Range range;

        public LexerECPE()
        {
            numberMem = 1;
            numberCheckSort = 1;
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
            this.numberCheckSort++;
            return 0;
        }

        public bool editNumber()
        {
            int checkValue = -1;
            this.countLength = 0;
            CheckStringMatch(this.sentence, @"^\[", ref checkValue, false);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[1-9][0-9]*", ref checkValue, true);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\]\s", ref checkValue, false);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        if (CheckSort())
                        {
                            return true;
                        }
                        FindAndReplace2("[" + numberMem + "]", "[" + numberCheckSort + "]");
                        numberCheckSort++;
                        return true;
                    }
                }
            }
            return false;
        }

        public bool CheckSort()
        {

            if ((this.numberMem - this.numberCheckSort) == 0)
            {
                this.numberCheckSort++;
                return true;
            }
            
            return false;
        }

        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
        }

        public bool FindAndReplace2(object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 1;
            object wrap = 1;

            //execute find and replace


            return range.Find.Execute(
                ref findText,
                ref matchCase,
                ref matchWholeWord,
                ref matchWildCards,
                ref matchSoundsLike,
                ref matchAllWordForms,
                ref forward,
                ref wrap,
                ref format,
                ref replaceWithText,
                ref replace,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl); 
        }
    }
}
