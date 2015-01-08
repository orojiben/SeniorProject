using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorProject
{
    class Styles
    {
        private string name;
        private string margin;
        private List<string> font;
        private List<string> dictionary;

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

        public string Margin
        {
            get
            {

                return this.margin;

            }
            set
            {

                this.margin = value;

            }
        }

        public List<string> Font
        {
            get
            {

                return this.font;

            }
            set { }
        }

        public List<string> Dictionary
        {
            get
            {

                return this.dictionary;

            }
            set { }
        }

        public Styles()
        {
            this.font = new List<string>();
            this.dictionary = new List<string>();
        }

        public void addFont(string font)
        {
            this.font.Add(font);
        }

        public void addDictionary(string dictionary)
        {
            this.dictionary.Add(dictionary);
        }


    }
}
