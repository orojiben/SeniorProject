using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Tools = Microsoft.Office.Tools;
using System.IO;
using System.Drawing;
using Microsoft.Office.Interop.Word;

namespace WordAddIn1
{
    public partial class UserControl1 : UserControl
    {

        public List<String> correct_word = new List<String>();
        public List<List<String>> incorrect_word = new List<List<String>>();

      
        public List<String> word_base = new List<String>();
        public List<String> word_adjust = new List<String>();
       

        public UserControl1()
        {
            InitializeComponent();
            load_Word();
            check_incorrect_word();
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


        public void clear()
        {
            this.correct_word.Clear();
            this.incorrect_word.Clear();

        }


        public void load_Word()
        {
          List<String> buff;
            string word = "";
            string path = "Royal_01.txt";
            bool stLine = true;
            if (!File.Exists(path))
            {
                this.Visible = false;
            }
            else
            {
                foreach (string line in File.ReadLines(path))
                {
                    buff = new List<string>();
                    stLine = true;
                    
                    foreach (char ch in line)
                    {
                        if (ch == ',')
                        {
                            if (stLine)
                            {
                                this.correct_word.Add(word);
                                stLine = false;
                                word = "";
                            }
                            else
                            {
                                buff.Add(word);
                                word = "";
                            }
                           
                        }
                        else
                        {
                            word = word + ch;
                        }
                    }
                    buff.Add(word);
                    this.incorrect_word.Add(new List<string>(buff));
                    buff.Clear();
                    word = "";
                }
            }
        }

        private void check_incorrect_word()
        {

            for (int x = 0; x < (incorrect_word.Count); ++x)
            {
                for (int y = 0; y < (incorrect_word[x].Count); ++y)
                {
                    if (FindAndReplace(incorrect_word[x][y]))
                    {
                        listBox1.Items.Add(incorrect_word[x][y]);
                        word_adjust.Add(incorrect_word[x][y]);
                        word_base.Add(correct_word[x]);
                    }
                }
            }
        }



        private void button4_Click(object sender, EventArgs e)
        {
           
        }

        private void button2_Click(object sender, EventArgs e)
        {

            for (int i = 0; i < this.word_adjust.Count; ++i)
            {
                this.FindAndReplace2((object)word_adjust[i], (object)word_base[i]);
                listBox1.Items.Remove(word_adjust[i]);
          
            }
            this.word_adjust.Clear();
            this.word_base.Clear();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            foreach (int selecteditem in listBox1.SelectedIndices)
            {
                this.FindAndReplace2((object)word_adjust[selecteditem], (object)word_base[selecteditem]);
                listBox1.Items.Remove((object)word_adjust[selecteditem]);
                word_adjust.RemoveAt(selecteditem);
                word_base.RemoveAt(selecteditem);
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            foreach (int selecteditem in listBox1.SelectedIndices)
            {
                highlightColor(word_adjust[selecteditem]);
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            Globals.ThisAddIn.Application.Selection.Find.ClearHitHighlight();
        }

    }
}
