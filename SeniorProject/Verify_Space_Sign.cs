using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;

namespace SeniorProject
{
    class Verify_Space_Sign
    {

        public List<String> correct_word = new List<String>();
        public List<List<String>> incorrect_word = new List<List<String>>();
        public List<String> word_containner = new List<String>();
        public Dictionary<String, String> dictionary = new Dictionary<String, String>();

        public List<String> regEx = new List<String>();



        public Verify_Space_Sign()
        {
            this.initializing();
            // this.wordsVerify_Forms();
            for (int i = 0; i < regEx.Count(); ++i) this.highlightColor(regEx[i]);
        }


        private void initializing()
        {
            regEx.Add("([a-z0-9]>.)( {3,})[A-Zก-ๆ0-9]");
            regEx.Add("([a-z0-9]>.)( )[A-Zก-ๆ0-9]");


            /*
            regEx.Add("([a-zก-ๆ0-9]>;)( {2,})[A-Zก-ๆ0-9]");
            regEx.Add("([a-zก-ๆ0-9]>;)[A-Zก-ๆ0-9]");

            
            regEx.Add("([a-zก-ๆ0-9]>:)[A-Zก-ๆ0-9]");
            regEx.Add("([a-zก-ๆ0-9]>:)( {3,})[A-Zก-ๆ0-9]");
            */



            regEx.Add("?,[! ]");
            regEx.Add("?( {1,}),?");
            regEx.Add("?,( {2,})?");
            regEx.Add("?( {1,}),( {2,})?");



            /*
            regEx.Add("(\")[a-zA-Zก-ๆ0-9]");
            regEx.Add("(\")( {2,})[a-zA-Zก-ๆ0-9]");

            regEx.Add("([a-zA-Zก-ๆ0-9](\")");
            regEx.Add("([a-zA-Zก-ๆ0-9]( {2,})(\")");
            */

        }


        private bool FindAndReplace(object findText)
        {
            //options
            object matchCase = false;
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
            object replaceWithText = "";
            //execute find and replace


            Word.Find f = Globals.ThisAddIn.Application.Selection.Find;


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


        private void highlightColor(object findText)
        {
            Color redColor = Color.FromArgb(0, 0, 255);
            Color blackColor = Color.FromArgb(255, 255, 255);

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
            {
            }



        }

        private void wordsVerify_Forms()
        {
            Word_Containner display_Forms = new Word_Containner();
            display_Forms.word_base.Clear();
            for (int x = 0; x < (incorrect_word.Count); ++x)
            {
                for (int y = 0; y < (incorrect_word[x].Count); ++y)
                {
                    if (FindAndReplace(incorrect_word[x][y]))
                    {
                        word_containner.Add(incorrect_word[x][y]);
                        display_Forms.listBox1.Items.Add(correct_word[x]);
                    }
                }
            }
            display_Forms.word_base.AddRange(word_containner);
            word_containner.Clear();
            display_Forms.Show();
        }

    }
}






/*


        private bool FindAndReplace(object findText)
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
            object replace = 0;
            object wrap = 1;
            object replaceWithText = "";
            //execute find and replace


            Word.Find f = Globals.ThisAddIn.Application.Selection.Find;


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

        private void highlightColor(object findText)
        {
            Color redColor = Color.FromArgb(255, 0, 0);
            Color blackColor = Color.FromArgb(0, 0, 0);

            object highlightColor = redColor;
            object textColor = blackColor;
            object matchCase = false;
            object matchWholeWord = true;
            object matchPrefix = false;
            object matchSuffix = false;
            object matchPhrase = false;
            object matchWildcard = false;
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

        private void wordsVrify_Forms()
        {
            Form1 display_Forms = new Form1();
            display_Forms.word_base.Clear();
            for (int x = 0; x < (incorrect_word.Count); ++x)
            {
                for (int y = 0; y < (incorrect_word[x].Count); ++y)
                {
                    if (FindAndReplace(incorrect_word[x][y]))
                    {
                        word_containner.Add(incorrect_word[x][y]);
                        display_Forms.listBox1.Items.Add(correct_word[x]);
                    }
                }
            }
            display_Forms.word_base.AddRange(word_containner);
            word_containner.Clear();
            display_Forms.Show();
        }
        
    }
}
*/