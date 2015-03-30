using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class ReferenceNameYear
    {
        public Word.Range range;
        private string name;
        private string year;
        private char indexCharacter = 'ก';
        private List<char> countCharacter;
        private bool forCheck;

        public string Name
        {
            get
            {

                return this.name;

            }
            set
            {

                this.name = value;

            }
        }

        public string Year
        {
            get
            {

                return this.year;

            }
            set
            {

                this.year = value;

            }
        }

        public List<char> CountCharacter
        {
            get
            {

                return this.countCharacter;

            }
            set
            {

                this.countCharacter = value;

            }
        }

        public bool ForCheck
        {
            get
            {

                return this.forCheck;

            }
            set
            {

                this.forCheck = value;

            }
        }

        public ReferenceNameYear(string name, string year, char character,bool forCheck)
        {
            this.name = name;
            this.year = year;
            this.countCharacter = new List<char>();
            this.forCheck = forCheck;
            if (character != ' ')
            {
                this.countCharacter.Add(character);
            }
        }

        public bool Check(ReferenceNameYear value)
        {
            if (!this.forCheck)
            {
                return true;
            }
            if (this.name == value.Name)
            {
                if (this.year == value.year)
                {
                    int numThis = this.countCharacter.Count;
                    int numValue = value.countCharacter.Count;
                    if (numThis > 0 && numValue > 0)
                    {
                        if ((value.countCharacter[numValue - 1] == 'ค' && this.countCharacter[numThis - 1] == 'ข')||
                            (value.countCharacter[numValue - 1] == 'ฆ' && this.countCharacter[numThis - 1] == 'ค')||
                            ((value.countCharacter[numValue - 1] - this.countCharacter[numThis - 1]) == 1))
                        {
                            return true;
                        }
                    }
                    return false;
                }
            }
            return true;
        }

        public bool Edit(ReferenceNameYear value)
        {
            if (!this.forCheck)
            {
                this.countCharacter.Clear();
                indexCharacter = 'ก';
                return true;
            }
            if (this.name == value.Name)
            {
                if (this.year == value.year)
                {
                    if (indexCharacter == 'ก')
                    {
                        string memChar = "";
                        if (this.countCharacter.Count > 0)
                        {
                            memChar = this.countCharacter[0] + "";
                        }
                        this.FindAndReplace2("(" + this.year + memChar + ")", "(" + this.year + indexCharacter + ")");
                        ++indexCharacter;
                        if (indexCharacter == 'ฃ' || indexCharacter == 'ฅ')
                        {
                            ++indexCharacter;
                        }
                        memChar = "";
                        if (value.countCharacter.Count > 0)
                        {
                            memChar = value.countCharacter[0] + "";
                        }
                        this.FindAndReplace2("(" + this.year + memChar + ")", "(" + this.year + indexCharacter + ")");
                        ++indexCharacter;
                        if (indexCharacter == 'ฃ' || indexCharacter == 'ฅ')
                        {
                            ++indexCharacter;
                        }
                    }
                    else
                    {
                        string memChar = "";
                        if (value.countCharacter.Count > 0)
                        {
                            memChar = value.countCharacter[0] + "";
                        }
                        this.FindAndReplace2("(" + this.year + memChar + ")", "(" + this.year + indexCharacter + ")");
                        ++indexCharacter;
                        if (indexCharacter == 'ฃ' || indexCharacter == 'ฅ')
                        {
                            ++indexCharacter;
                        }
                    }

                            return true;
                }
            }
            this.countCharacter.Clear();
            indexCharacter = 'ก';
            return true;
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
