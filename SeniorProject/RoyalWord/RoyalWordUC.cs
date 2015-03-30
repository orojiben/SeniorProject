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
using Microsoft.Office.Interop.Word;



namespace SeniorProject
{
    public partial class RoyalWordUC : UserControl
    {

        public class WordSet
        {
            public String correct_word;
            public String target;
            public List<Word.Range> rang;
            public int count;
        }

        FindWordB FB = new FindWordB();
        public List<String> correct_word = new List<String>();
        public List<List<String>> incorrect_word = new List<List<String>>();
        public List<List<WordSet>> wordset = new List<List<WordSet>>();
        //public List<String> word_base = new List<String>();
        //public List<String> word_adjust = new List<String>();
        public string path = "";
        public string folder = "";
        public string pathDoc = "";
        public int indexI, indexII, indexIII, max, min, selectnumall;


        public RoyalWordUC()
        {
            InitializeComponent();
              folder = "CheckingThesis";
              pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
              path = pathDoc + "\\" + folder + "\\Royal.txt";
            indexI = 0;
            max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            min = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;
        }

        public void setpathFile(string path)
        {
            this.path = path;
        }

        public void enableAll()
        {

            btn_fullStopEdit.Enabled = false;
            btn_check.Enabled = true;
            listBox1.Enabled = false;
            btn_next.Enabled = false;
            btn_highlightALL.Enabled = false;

            btn_clearHighlightALL.Enabled = false;
            progressBar.Value = 0;
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;

        }

        public void visibleAllOpen()
        {
            btn_fullStopEdit.Enabled = true;
            //  btn_editAll.Enabled = true;
            listBox1.Enabled = true;
            btn_next.Enabled = true;
            btn_highlightALL.Enabled = true;
            btn_highlightALL.Enabled = true;
            btn_clearHighlightALL.Enabled = true;
            btn_clearHighlightALL.Enabled = true;
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;
        }

        public void visibleAllClose()
        {
            btn_fullStopEdit.Enabled = false;
            listBox1.Enabled = false;
            btn_next.Enabled = false;
            btn_highlightALL.Enabled = false;
            btn_clearHighlightALL.Enabled = false;
            progressBar.Enabled = true;
            lbl_waitCheck.Enabled = true;
        }

        public void checkFinish()
        {
            if (wordset.Count == 0)
            {
                visibleAllClose();
            }
            Ribbon1.showCheckAllUC.setButtonClickALL();
        }


        public void clear()
        {
            //word_base.Clear();
            //word_adjust.Clear();
            listBox1.Items.Clear();
            listBox1.ClearSelected();
            wordset.Clear();
            this.correct_word.Clear();
            this.incorrect_word.Clear();
        }

        public void load_Word()
        {
            List<String> buff;
            string word = "";
            string path = this.path;
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

        public void check_incorrect_word()
        {
            clear();
            load_Word();
            WordSet ws;
            List<WordSet> wsL;
            List<Word.Range> r;
            Word.Range rangeAll = Globals.ThisAddIn.Application.ActiveDocument.Content;
            this.progressBar.Maximum = incorrect_word.Count;
            for (int y = 0; y < (incorrect_word.Count); ++y)
            {
                this.progressBar.Increment(1);
                wsL = new List<WordSet>();
                for (int x = 0; x < (incorrect_word[y].Count); ++x)
                {
                    if (FB.find31(incorrect_word[y][x], Globals.ThisAddIn.Application.ActiveDocument.Content.Find))
                    {
                        ws = new WordSet();
                        ws.correct_word = correct_word[y].ToString();
                        ws.target = incorrect_word[y][x].ToString();
                        r = new List<Word.Range>();
                        rangeAll = Globals.ThisAddIn.Application.ActiveDocument.Content;

                        while (FB.find31(incorrect_word[y][x], rangeAll.Find))
                        {
                            r.Add(FB.makeRange(rangeAll));
                        }

                        ws.rang = r;
                        wsL.Add(ws);
                    }
                }
                wordset.Add(wsL);
            }
            exclude();
            cInRange();
            addbox();
        }

        public void exclude()
        {
            List<List<WordSet>> ws = new List<List<WordSet>>();
            for (int y = 0; y < (wordset.Count); ++y)
            {             // row
                if (wordset[y].Count > 0)
                {
                    ws.Add(wordset[y]);
                }
            }
            wordset = ws;
        }

        public void addbox()
        {
            for (int i = 0; i < wordset.Count; ++i)
            {
                listBox1.Items.Add(wordset[i][0].correct_word.ToString());
            }
        }




        public void cInRange()
        {
            String tmp;
            bool atone = true;
            List<Word.Range> lR;
            for (int y = 0; y < (wordset.Count); ++y)
            {             // row
                for (int x = 0; x < (wordset[y].Count - 1); ++x)
                {      // colum
                    for (int i = x + 1; i < (wordset[y].Count); ++i)
                    {   // colum +1
                        if (wordset[y][x].target.Length > wordset[y][i].target.Length)
                        {
                            tmp = wordset[y][x].target.Remove(wordset[y][i].target.Length);
                            if (wordset[y][i].target == tmp)
                            {
                                lR = new List<Word.Range>();
                                for (int r = 0; r < wordset[y][i].rang.Count; ++r)
                                {


                                    atone = true;
                                    for (int s = 0; s < wordset[y][x].rang.Count; ++s)
                                    {

                                        if (wordset[y][i].rang[r].InRange(wordset[y][x].rang[s])) atone = atone && false;
                                        else atone = atone && true;
                                    }
                                    if (atone) lR.Add(wordset[y][i].rang[r]);
                                }
                                wordset[y][i].rang = lR;
                            }
                        }
                        else if (wordset[y][x].target.Length < wordset[y][i].target.Length)
                        {
                            tmp = wordset[y][i].target.Remove(wordset[y][x].target.Length);
                            if (wordset[y][x].target == tmp)
                            {
                                lR = new List<Word.Range>();
                                for (int r = 0; r < wordset[y][x].rang.Count; ++r)
                                {
                                    atone = true;
                                    for (int s = 0; s < wordset[y][i].rang.Count; ++s)
                                    {
                                        if (wordset[y][x].rang[r].InRange(wordset[y][i].rang[s])) atone = atone && false;
                                        else atone = atone && true;
                                    }
                                    if (atone) lR.Add(wordset[y][x].rang[r]);
                                }
                                wordset[y][x].rang = lR;
                            }
                        }
                    }
                }
            }
        }


        public void hilightall(WdColorIndex color)
        {
            for (int y = 0; y < wordset.Count; ++y)
            {
                for (int x = 0; x < wordset[y].Count; ++x)
                {
                    for (int i = 0; i < wordset[y][x].rang.Count; ++i)
                    {
                        wordset[y][x].rang[i].HighlightColorIndex = color;
                    }
                }
            }
        }


        public void editAll()
        {
            Word.Range vw;
            for (int y = 0; y < wordset.Count; ++y)
            {
                for (int x = 0; x < wordset[y].Count; ++x)
                {
                    for (int i = 0; i < wordset[y][x].rang.Count; ++i)
                    {
                        vw = FB.makeRange(wordset[y][x].rang[i]);

                        if (vw.Start <= min) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start, vw.End + 1);
                        else if (vw.End >= max) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End);
                        else vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End + 1);

                        FB.findReplace22(wordset[y][x].target, wordset[y][x].correct_word, vw);
                    }
                }
            }
            listBox1.ForeColor = Color.Gray;
        }

        public void editSelect()
        {
            Word.Range vw;
            for (int y = 0; y < wordset.Count; ++y)
            {
                for (int x = 0; x < wordset[y].Count; ++x)
                {
                    for (int i = 0; i < wordset[y][x].rang.Count; ++i)
                    {
                        vw = FB.makeRange(wordset[y][x].rang[i]);

                        if (vw.Start <= min) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start, vw.End + 1);
                        else if (vw.End >= max) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End);
                        else vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End + 1);

                        FB.findReplace22(wordset[y][x].target, wordset[y][x].correct_word, vw);
                    }
                }
            }
            listBox1.ForeColor = Color.Gray;
        }


        public void contAll()
        {
            int cont = 0;

            for (int y = 0; y < wordset.Count; ++y)
            {
                for (int x = 0; x < wordset[y].Count; ++x)
                {
                    for (int i = 0; i < wordset[y][x].rang.Count; ++i)
                    {
                        cont++;
                    }
                }

            }
            label9.Text = cont.ToString();
        }

        public int contselect(int index)
        {
            int cont = 0;
            for (int x = 0; x < wordset[index].Count; ++x)
            {
                for (int i = 0; i < wordset[index][x].rang.Count; ++i)
                {
                    cont++;
                }
            }
            return cont;

        }

        public int contselect(int index, int num)
        {
            int cont = 0;
            for (int x = 0; x < wordset[index].Count; ++x)
            {
                for (int i = 0; i < wordset[index][x].rang.Count; ++i)
                {
                    ++cont;
                    if (cont == num)
                    {
                        wordset[index][x].rang[i].Select();
                    }
                }
            }
            return cont;

        }

        public int selectactive(int index, int num)
        {
            Word.Range vw;
            int cont = 0;
            for (int x = 0; x < wordset[index].Count; ++x)
            {
                for (int i = 0; i < wordset[index][x].rang.Count; ++i)
                {
                    ++cont;
                    if (cont == num)
                    {
                        vw = FB.makeRange(wordset[index][x].rang[i]);

                        if (vw.Start <= min) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start, vw.End + 1);
                        else if (vw.End >= max) vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End);
                        else vw = Globals.ThisAddIn.Application.ActiveDocument.Range(vw.Start - 1, vw.End + 1);
                        FB.findReplace22(wordset[index][x].target, wordset[index][x].correct_word, vw);
                    }
                }
            }
            return cont;

        }

        public void checkWordAll()
        {
            visibleAllClose();
            check_incorrect_word();
            contAll();
            visibleAllOpen();
            checkFinish();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            checkWordAll();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            wordset[indexI][indexII].rang[indexIII].Select();
        }





        private void btn_highlightALL_Click(object sender, EventArgs e)
        {
            hilightall(WdColorIndex.wdRed);
        }

        private void btn_clearHighlightALL_Click(object sender, EventArgs e)
        {
            hilightall(WdColorIndex.wdNoHighlight);
        }


        private void button1_Click(object sender, EventArgs e)
        {

        }
        private void button3_Click(object sender, EventArgs e)
        {

        }

        public void editWordAll()
        {
            visibleAllClose();
            editAll();
            check_incorrect_word();
            contAll();
            visibleAllOpen();
            checkFinish();
        }

        public void editWordAllForAll()
        {
            visibleAllClose();
            editAll();
        }

        private void btn_fullStopEdit_Click(object sender, EventArgs e)
        {
            editWordAll();
        }


        private void btn_back_Click(object sender, EventArgs e)
        {

        }

        private void btn_next_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex < 0) { }
            else
            {
                ++indexIII;
                label8.Text = indexIII.ToString();
                contselect(indexI, indexIII);
                indexIII = indexIII % selectnumall;
            }

        }

        private void button2_Click_1(object sender, EventArgs e)
        {
            selectactive(indexI, indexIII);
        }

        private void listBox1_SelectedValueChanged(object sender, EventArgs e)
        {
            indexI = listBox1.SelectedIndex;
            indexII = 0;
            indexIII = 0;
            selectnumall = contselect(listBox1.SelectedIndex);
            label10.Text = selectnumall.ToString();
        }





    }
}
