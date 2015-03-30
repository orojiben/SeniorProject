using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    public class StyleFont
    {
        private string fontName;
        private string fontNameLanguage;
        private float coverTitle;
        private float coverOperator;
        private float chapter;
        private float namechapter;
        private float topics;
        private float subheading;
        private float substance;
        public string FontName
        {
            get
            {

                return this.fontName;

            }
            set
            {

            }
        }

        public string FontNameLanguage
        {
            get
            {

                return this.fontNameLanguage;

            }
            set
            {

            }
        }

        public float CoverTitle
        {
            get
            {

                return this.coverTitle;

            }
            set
            {

            }
        }

        public float CoverOperator
        {
            get
            {

                return this.coverOperator;

            }
            set
            {

            }
        }

        public float Chapter
        {
            get
            {

                return this.chapter;

            }
            set
            {

            }
        }

        public float Namechapter
        {
            get
            {

                return this.namechapter;

            }
            set
            {

            }
        }
        public float Topics
        {
            get
            {

                return this.topics;

            }
            set
            {

            }
        }

        public float Subheading
        {
            get
            {

                return this.subheading;

            }
            set
            {

            }
        }

        public float Substance
        {
            get
            {

                return this.substance;

            }
            set
            {

            }
        }

        public StyleFont(string fontName, string fontNameLanguage, float coverTitle, float coverOperator, float chapter,
            float namechapter, float topics, float subheading, float substance)
        {
            this.fontName = fontName;
            this.fontNameLanguage = fontNameLanguage;
            this.coverTitle = coverTitle;
            this.coverOperator = coverOperator;
            this.chapter = chapter;
            this.namechapter = namechapter;
            this.topics = topics;
            this.subheading = subheading;
            this.substance = substance;
        }
    }
}
