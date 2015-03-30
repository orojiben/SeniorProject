using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorProject
{
    public class Styles
    {
        private string name;
        private string margin;
        private string paper;
        private List<StyleFont> styleFont;
        //private List<string> dictionarys;
        private List<string> departments;
        private float indent;
        private float leftMargin;
        private float rightMargin;
        private float topMargin;
        private float bottomMargin;
 
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
                string[] words =  this.margin.Split(',');
                this.leftMargin = centimeterToPoint((float.Parse(words[0])));
                this.leftMargin = float.Parse(String.Format(CultureInfo.InvariantCulture,
                                      "{0:0.0}", this.leftMargin));

                this.rightMargin = centimeterToPoint((float.Parse(words[1])));
                this.rightMargin = float.Parse(String.Format(CultureInfo.InvariantCulture,
                                      "{0:0.0}", this.rightMargin));

                this.topMargin = centimeterToPoint((float.Parse(words[2])));
                this.topMargin = float.Parse(String.Format(CultureInfo.InvariantCulture,
                                      "{0:0.0}", this.topMargin));
                this.bottomMargin = centimeterToPoint((float.Parse(words[3])));
                this.bottomMargin = float.Parse(String.Format(CultureInfo.InvariantCulture,
                                      "{0:0.0}", this.bottomMargin));
            }
        }

        public string Paper
        {
            get
            {

                return this.paper;

            }
            set
            {

                this.paper = value;

            }
        }

        public List<StyleFont> StyleFont
        {
            get
            {

                return this.styleFont;

            }
            set { }
        }

        /*public List<string> Dictionarys
        {
            get
            {

                return this.dictionarys;

            }
            set { }
        }*/

        public List<string> Departments
        {
            get
            {

                return this.departments;

            }
            set { }
        }

        public float Indent
        {
            get
            {

                return this.indent;

            }
            set
            {

                this.indent = value;
                this.indent = float.Parse(String.Format(CultureInfo.InvariantCulture,
                                      "{0:0.0}", this.indent));

            }
        }

        public float LeftMargin
        {
            get
            {

                return this.leftMargin;

            }
            set
            {


            }
        }

        public float RightMargin
        {
            get
            {

                return this.rightMargin;

            }
            set
            {

            }
        }

        public float TopMargin
        {
            get
            {

                return this.topMargin;

            }
            set
            {
            }
        }

        public float BottomMargin
        {
            get
            {

                return this.bottomMargin;

            }
            set
            {

            }
        }

        public Styles()
        {
            this.styleFont = new List<StyleFont>();
            //this.dictionarys = new List<string>();
            this.departments = new List<string>();
        }

        public void addFont(string fontName, string fontNameLanguage, float coverTitle, float coverOperator, float chapter,
            float namechapter, float topics, float subheading, float substance)
        {
            this.styleFont.Add(new StyleFont( fontName,  fontNameLanguage,  coverTitle,  coverOperator,  chapter,
             namechapter,  topics,  subheading,  substance));
        }

       /* public void addDictionary(string dictionary)
        {
            this.dictionarys.Add(dictionary);
        }*/

        public void addDepartment(string department)
        {
            this.departments.Add(department);
        }

        public StyleFont getStyleFont(string nameFont, string Language)
        {
            foreach (StyleFont sf in this.styleFont)
            {
                if (sf.FontName == nameFont && sf.FontNameLanguage == Language)
                {
                    return sf;
                }
            }
            return null;
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
        }

    }
}
