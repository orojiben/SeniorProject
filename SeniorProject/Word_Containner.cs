using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Tools.Ribbon;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Reflection;
namespace SeniorProject
{
    public partial class Word_Containner : Form
    {
        public Word_Containner()
        {
            InitializeComponent();
        }





        public bool FindAndReplace2(object findText, object replaceWithText)
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
            object replace = 2;
            object wrap = 1;

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






        public List<String> word_base = new List<String>();
        object word_ajust;
        object bse;

        private void button1_Click(object sender, EventArgs e)
        {

            word_ajust = listBox1.SelectedItem.ToString();
            bse = word_base[listBox1.SelectedIndex].ToString();
            FindAndReplace2(bse, word_ajust);

        }




        private void button3_Click(object sender, EventArgs e)
        {
            word_base.Clear();
            this.Close();
        }




        public void highlightColor(object findText)
        {
            Color redColor = Color.FromArgb(255, 0, 0);
            Color blackColor = Color.FromArgb(0, 0, 0);

            object highlightColor = redColor;
            object textColor = blackColor;
            object matchCase = true;
            object matchWholeWord = false;
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



        private void button2_Click(object sender, EventArgs e)
        {
            //highlightColor("ดิจิตอล");
            highlightColor(word_base[listBox1.SelectedIndex].ToString());


        }
    }
}
