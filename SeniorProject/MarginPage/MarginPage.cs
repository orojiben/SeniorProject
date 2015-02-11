using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    class MarginPage
    {
        private string title;
        private float leftOld;
        private float rightOld;
        private float topOld;
        private float bottomOld;

        public MarginPage(float left, float right, float top, float bottom)
        {
            this.leftOld = left;
            this.rightOld = right;
            this.topOld = top;
            this.bottomOld = bottom;
        }

        public float getLeft()
        {
            return this.leftOld;

        }

        public float getRight()
        {
            return this.rightOld;

        }

        public float getTop()
        {
            return this.topOld;

        }

        public float getBottom()
        {
            return this.bottomOld;

        }

        public float inchToPoint(float inch)
        {
            return 72f * inch;
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
        }

        public bool cheking()
        {
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;

            if ((pageSetup.LeftMargin == this.leftOld) &&
                (pageSetup.RightMargin == this.rightOld) &&
                (pageSetup.TopMargin == this.topOld) &&
                (pageSetup.BottomMargin == this.bottomOld))
            {
                return true;
            }
            return false;
        }

        public void changing()
        {
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            pageSetup.LeftMargin = this.leftOld;
            pageSetup.RightMargin = this.rightOld;
            pageSetup.TopMargin = this.topOld;
            pageSetup.BottomMargin = this.bottomOld;
        }

    }
}
