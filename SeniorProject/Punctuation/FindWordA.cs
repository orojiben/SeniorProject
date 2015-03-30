using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    class FindWordA
    {



        public FindWordA()
        {


        }


        public bool Find(object findText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = true;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 0;
            object wrap = 0;
            object replaceWithText = Type.Missing;

            Microsoft.Office.Interop.Word.Find f = Globals.ThisAddIn.Application.Selection.Find;



            return f.Execute(
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





        public bool FindAndReplace(object findText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = true;
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
            object replace = 0;
            object wrap = 1;
            object replaceWithText = Type.Missing;
            //execute find and replace


            Microsoft.Office.Interop.Word.Find f = Globals.ThisAddIn.Application.Selection.Find;


            return f.Execute(
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


        public bool FindAndReplace2(object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = true;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = false;/********/
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 2;
            object wrap = 1;


            Microsoft.Office.Interop.Word.Find f = Globals.ThisAddIn.Application.Selection.Find;


            return f.Execute(
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

        public bool FindtoEnd(object findText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = true;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;/********/
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 0;
            object wrap = 2;

            Microsoft.Office.Interop.Word.Find f = Globals.ThisAddIn.Application.Selection.Find;

            return f.Execute(
                ref findText,
                ref matchCase,
                ref matchWholeWord,
                ref matchWildCards,
                ref matchSoundsLike,
                ref matchAllWordForms,
                ref forward,
                ref wrap,
                ref format,
                ref findText,
                ref replace,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl);
        }


        public int FindCount(object findText)
        {
            int count = 0;
            while (Find(findText) == true) ++count;
            //FindAndReplace(findText);
            return count;
        }


        public bool FindAndReplaceNext(object findText, object replaceWithText)
        {
            //options
            object matchCase = true;
            object matchWholeWord = false;
            object matchWildCards = true;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;/********/
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object replace = 1;
            object replace2 = 0;
            object wrap = 0;


            Microsoft.Office.Interop.Word.Find f = Globals.ThisAddIn.Application.Selection.Find;


            f.Execute(
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

            f.Execute(
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
                ref replace2,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl);

            f.Execute(
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
            ref replace2,
            ref matchKashida,
            ref matchDiacritics,
            ref matchAlefHamza,
            ref matchControl);

            return true;
        }




        public void highlightColor(object findText)
        {
            Color redColor = Color.FromArgb(255, 0, 255);
            Color blackColor = Color.FromArgb(0, 0, 0);


            object highlightColor = redColor;
            object textColor = blackColor;
            object matchCase = false;
            object matchWholeWord = false;
            object matchPrefix = false;
            object matchSuffix = false;
            object matchPhrase = false;
            object matchWildcard = true;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object matchByte = false;
            object matchFuzzy = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object ignoreSpace = false;
            object ignorePunct = false;
            object hanjaPhoneticHangul = false;

            try
            {

                Globals.ThisAddIn.Application.Selection.Find.HitHighlight(
                    ref findText,
                    ref highlightColor,
                    ref textColor,
                    ref matchCase,
                    ref matchWholeWord,
                    ref matchPrefix,
                    ref matchSuffix,
                    ref matchPhrase,
                    ref matchWildcard,
                    ref matchSoundsLike,
                    ref matchAllWordForms,
                    ref matchByte,
                    ref matchFuzzy,
                    ref matchKashida,
                    ref matchDiacritics,
                    ref matchAlefHamza,
                    ref matchControl,
                    ref ignoreSpace,
                    ref ignorePunct,
                    ref hanjaPhoneticHangul);
            }
            catch
                (Exception e)
            {
            }
        }


    }
}
