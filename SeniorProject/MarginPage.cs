using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SeniorProject
{
    class MarginPage
    {
        private string title;
        private float left;
        private float right;
        private float top;
        private float bottom;
        public MarginPage(string title, string left, string right, string top, string bottom, string tyepe)
        {
            this.title = title;
            if (tyepe == "change centimeter")
            {
                this.left = inchToPoint((float)(Convert.ToDouble(left)));
                this.right = inchToPoint((float)(Convert.ToDouble(right)));
                this.top = inchToPoint((float)(Convert.ToDouble(top)));
                this.bottom = inchToPoint((float)(Convert.ToDouble(bottom)));
            }
            else
            {
                this.left = centimeterToPoint((float)(Convert.ToDouble(left)));
                this.right = centimeterToPoint((float)(Convert.ToDouble(right)));
                this.top = centimeterToPoint((float)(Convert.ToDouble(top)));
                this.bottom = centimeterToPoint((float)(Convert.ToDouble(bottom)));
            }

        }

        public MarginPage(string title, float left, float right, float top, float bottom)
        {
            this.title = title;
            this.left = left;
            this.right = right;
            this.top = top;
            this.bottom = bottom;
        }

        public float getLeft()
        {
            return this.left;

        }

        public float getRight()
        {
            return this.right;

        }

        public float getTop()
        {
            return this.top;

        }

        public float getBottom()
        {
            return this.bottom;

        }

        public string getTitle()
        {
            return this.title;

        }


        public float inchToPoint(float inch)
        {
            return 72f * inch;
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
        }
    }
}
