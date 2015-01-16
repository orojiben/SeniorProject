using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using System.Drawing;

namespace SeniorProject
{
    class Verify_Royal_Word_TH
    {
        public List<String> correct_word = new List<String>();
        public List<List<String>> incorrect_word = new List<List<String>>();
        public List<String> word_containner = new List<String>();
        public Dictionary<String, String> dictionary = new Dictionary<String, String>();


        public Verify_Royal_Word_TH()
        {
            this.initializing();
            this.wordsVerify_Forms();

        }




        private void initializing()
        {

            correct_word.Add("ดิจิทัล");
            correct_word.Add("เรจิสเตอร์");
            correct_word.Add("แมโคร");
            correct_word.Add("เมท็อด");
            correct_word.Add("เบราว์เซอร์");


            incorrect_word.Add(new List<String> { "ดิจิตอล", "ดิจิตอน", "ดิจิทอล", "ดิจิทอน" });
            incorrect_word.Add(new List<String> { "รีจิสเตอร์", "รีจิดเตอร์" });
            incorrect_word.Add(new List<String> { "มาโคร", "แมโคล", "มาโคล" }); // เจอ 2 ครั้ง ใน 1 คำ มาโคร    จะได้ทั้ง มาโคร และ มาโค
            incorrect_word.Add(new List<String> { "เม็ตท๊อด", "เม็ดต๊อด", "เมธอท" });
            incorrect_word.Add(new List<String> { "บราวเซอร์", "บาวเซอร์", "บลาวเซอร์" });

        }


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



        private void customPanel()
        {
            // Tools.CustomTaskPane cp;
            //cp = new CustomTaskPane();

        }

    }
}