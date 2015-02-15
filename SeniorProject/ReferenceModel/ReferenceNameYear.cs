using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    class ReferenceNameYear
    {
        private string name;
        private string year;
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

    }
}
