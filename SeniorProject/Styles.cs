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
        private string paper;
        private List<string> fonts;
        private List<string> dictionarys;
        private List<string> departments;

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

        public List<string> Fonts
        {
            get
            {

                return this.fonts;

            }
            set { }
        }

        public List<string> Dictionarys
        {
            get
            {

                return this.dictionarys;

            }
            set { }
        }

        public List<string> Departments
        {
            get
            {

                return this.departments;

            }
            set { }
        }

        public Styles()
        {
            this.fonts = new List<string>();
            this.dictionarys = new List<string>();
            this.departments = new List<string>();
        }

        public void addFont(string font)
        {
            this.fonts.Add(font);
        }

        public void addDictionary(string dictionary)
        {
            this.dictionarys.Add(dictionary);
        }

        public void addDepartment(string department)
        {
            this.departments.Add(department);
        }

    }
}
