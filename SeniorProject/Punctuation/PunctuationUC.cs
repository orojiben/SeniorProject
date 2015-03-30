using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    public partial class PunctuationUC : UserControl
    {

        FindWordB FB = new FindWordB();

        public class RangeNum
        {
            public Word.Range rang;
            public bool pos;

        }

        Word.Range rangeAll;
        List<Word.Range> rtmp;
        List<Word.Range> rCL, rCL_Worng, rCL_Bank;
        List<Word.Range> rSCL, rSCL_Worng, rSCL_Bank;
        List<Word.Range> rDQ;

        List<Word.Range> rFS, rFS2, rFS3, rFS_Worng, rFS_Bank;
        List<Word.Range> rCM, rCM_Worng, rCM_Bank;
        List<RangeNum> rDQ_Worng2;

        int index_of_Selection;

        public PunctuationUC()
        {

            InitializeComponent();
            index_of_Selection = -1;
            rangeAll = Globals.ThisAddIn.Application.ActiveDocument.Content;
            rCL = new List<Word.Range>();
            rSCL = new List<Word.Range>();
            rDQ = new List<Word.Range>();
            rFS = new List<Word.Range>();
            rFS2 = new List<Word.Range>();
            rFS3 = new List<Word.Range>();
            rCM = new List<Word.Range>();
            rtmp = new List<Word.Range>();

            rCL_Worng = new List<Word.Range>();
            rSCL_Worng = new List<Word.Range>();

            rFS_Worng = new List<Word.Range>();
            rCM_Worng = new List<Word.Range>();
            rDQ_Worng2 = new List<RangeNum>();

            rCL_Bank = new List<Word.Range>();
            rSCL_Bank = new List<Word.Range>();

            rFS_Bank = new List<Word.Range>();
            rCM_Bank = new List<Word.Range>();
        }


        public RangeNum makeRN(Word.Range rang, bool pos)
        {
            RangeNum tem = new RangeNum();
            tem.pos = pos;
            tem.rang = rang;
            return tem;
        }


        //Word.WdColorIndex.wdNoHighlight
        //Word.WdColorIndex.wdYellow

        public void hightlight(List<Word.Range> item, Word.WdColorIndex color)
        {
            foreach (Word.Range r in item) r.HighlightColorIndex = color;
        }
        public void hightlight(List<RangeNum> item, Word.WdColorIndex color)
        {
            foreach (RangeNum r in item) r.rang.HighlightColorIndex = color;
        }


        public List<Word.Range> Exclude(List<Word.Range> rang, String sign)
        {
            int start = 0;
            int end = 0;
            List<Word.Range> rang3 = new List<Word.Range>();
            Word.Range rang2 = null;
            for (int i = 0; i < rang.Count; ++i)
            {
                start = rang[i].Start - 1;
                end = rang[i].End + 1;
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                if (rang2.Characters.First.Text == sign || rang2.Characters.Last.Text == sign)
                {
                }
                else
                {
                    rang3.Add(rang[i]);
                }
            }
            return rang3;
        }


        public List<Word.Range> Exclude2(List<Word.Range> rang)
        {
            int toc = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents.Count;
            int tof = Globals.ThisAddIn.Application.ActiveDocument.TablesOfFigures.Count;
            int tocStart = 0;
            int tocEnd = 0;
            int tofStart = 0;
            int tofEnd = 0;

            List<Word.Range> rang3 = new List<Word.Range>();

            if (toc > 0)
            {
                tocStart = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents[1].Range.Start;
                tocEnd = Globals.ThisAddIn.Application.ActiveDocument.TablesOfContents[toc].Range.End;
            }
            if (tof > 0)
            {
                tofStart = Globals.ThisAddIn.Application.ActiveDocument.TablesOfFigures[1].Range.Start;
                tofEnd = Globals.ThisAddIn.Application.ActiveDocument.TablesOfFigures[tof].Range.End;
            }

            for (int i = 0; i < rang.Count; ++i)
            {
                if ((rang[i].Start < tocStart && rang[i].End <= tocStart) || (rang[i].Start >= tocEnd && rang[i].End >= tocEnd) &&
                    (rang[i].Start < tofStart && rang[i].End <= tofStart) || (rang[i].Start >= tofEnd && rang[i].End >= tofEnd))
                {
                    rang3.Add(rang[i]);
                }

                else { }
            }
            return rang3;
        }

        public void enableAll()
        {

            btn_editAll.Enabled = false;
            btn_check.Enabled = true;
            groupBox1.Enabled = false;
            btn_highlightALL.Enabled = false;

            btn_clearHighlightALL.Enabled = false;
            progressBar.Value = 0;
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;

            lbl_fullStopValue.Text = "-";
            lbl_commaValue.Text = "-";
            lbl_colonValue.Text = "-";
            lbl_semiColonValue.Text = "-";
            lbl_doubleQuoteValue.Text = "-";
            lbl_fullStopCheck.Text = "✘";
            lbl_commaCheck.Text = "✘";
            lbl_colonCheck.Text = "✘";
            lbl_semiColonCheck.Text = "✘";
            lbl_doubleQuoteCheck.Text = "✘";
        }

        public void visibleAllOpen()
        {
            btn_editAll.Enabled = true;
            //  btn_editAll.Enabled = true;
            btn_check.Enabled = true;
            groupBox1.Enabled = true;
            btn_highlightALL.Enabled = true;
            btn_highlightALL.Enabled = true;
            btn_clearHighlightALL.Enabled = true;
            btn_clearHighlightALL.Enabled = true;
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;
        }

        public void visibleAllClose()
        {
            btn_check.Enabled = false;
            btn_editAll.Enabled = false;
            groupBox1.Enabled = false;
            btn_highlightALL.Enabled = false;
            btn_clearHighlightALL.Enabled = false;
            progressBar.Enabled = true;
            lbl_waitCheck.Enabled = true;
        }

        public void HarvestAll()
        {
            HarvestFS();
            HarvestCM();
            HarvestCL();
            HarvestSCL();
            HarvestDQ();
        }

        public void checkAll()
        {
            visibleAllClose();
            progressBar.Value = 0;
            // Thread newThread = new Thread(this.checkEndLoad);
            // newThread.Start();
            //visibleAllClose();
            progressBar.Increment(1);

            HarvestAll();
            progressBar.Increment(1);

            CheckFullStop();
            progressBar.Increment(1);

            CheckComma();
            progressBar.Increment(1);

            CheckColon();
            progressBar.Increment(1);

            CheckSemiColon();

            progressBar.Increment(1);

            CheckDQ();
            progressBar.Increment(1);

            visibleAllOpen();
            checkFinish();
            // progressBar.Value = 100;
            //Thread.Sleep(3000);
            //visibleAllOpen();
        }

        public void checkAllForAll()
        {
            visibleAllClose();
            HarvestAll();
            CheckFullStop();
            CheckComma();
            CheckColon();
            CheckSemiColon();
            CheckDQ();
            visibleAllOpen();
            checkFinish();
        }

        public void checkFinish()
        {
            // System.Windows.Forms.MessageBox.Show("" + value + " " + this.radioButton1.Checked + " " + this.btn_fullStopEdit.Enabled);
            this.btn_fullStopEdit.Enabled = !(this.rFS_Worng.Count == 0);
            this.radioButton1.Checked = this.btn_fullStopEdit.Enabled;
            // System.Windows.Forms.MessageBox.Show("" + value + " " + this.radioButton1.Checked + " " + this.btn_fullStopEdit.Enabled);
            this.radioButton1.Enabled = true;
            if (!this.radioButton1.Checked)
            {
                this.radioButton1.Enabled = false;

            }
            this.btn_commaEdit.Enabled = !(this.rCM_Worng.Count == 0);
            this.radioButton2.Checked = this.btn_commaEdit.Enabled;
            this.radioButton2.Enabled = true;
            if (!this.radioButton2.Checked)
            {
                this.radioButton2.Enabled = false;

            }
            this.btn_colonEdit.Enabled = !(this.rCL_Worng.Count == 0);
            this.radioButton3.Checked = this.btn_colonEdit.Enabled;
            this.radioButton3.Enabled = true;
            if (!this.radioButton3.Checked)
            {
                this.radioButton3.Enabled = false;

            }
            this.btn_semiColonEdit.Enabled = !(this.rSCL_Worng.Count == 0);
            this.radioButton4.Checked = this.btn_semiColonEdit.Enabled;
            this.radioButton4.Enabled = true;
            if (!this.radioButton4.Checked)
            {
                this.radioButton4.Enabled = false;

            }
            this.btn_doubleQuoteEdit.Enabled = !(this.rDQ_Worng2.Count == 0);
            this.radioButton5.Checked = this.btn_doubleQuoteEdit.Enabled;
            this.radioButton5.Enabled = true;
            if (!this.radioButton5.Checked)
            {
                this.radioButton5.Enabled = false;
            }
            this.btn_editAll.Enabled = (this.btn_colonEdit.Enabled ||
                this.btn_semiColonEdit.Enabled ||
                this.btn_fullStopEdit.Enabled ||
                this.btn_commaEdit.Enabled ||
                this.btn_doubleQuoteEdit.Enabled);
            this.btn_highlightALL.Enabled = this.btn_editAll.Enabled;
            this.btn_clearHighlightALL.Enabled = this.btn_editAll.Enabled;
            this.btn_next.Enabled = this.btn_editAll.Enabled;
            this.btn_back.Enabled = this.btn_editAll.Enabled;
            if (this.btn_fullStopEdit.Enabled)
            {

                this.lbl_fullStopCheck.Text = "✘";
                this.lbl_fullStopCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_fullStopCheck.Text = "✔";
                this.lbl_fullStopCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (this.btn_colonEdit.Enabled)
            {
                this.lbl_colonCheck.Text = "✘";
                this.lbl_colonCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_colonCheck.Text = "✔";
                this.lbl_colonCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (this.btn_semiColonEdit.Enabled)
            {
                this.lbl_semiColonCheck.Text = "✘";
                this.lbl_semiColonCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_semiColonCheck.Text = "✔";
                this.lbl_semiColonCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (this.btn_doubleQuoteEdit.Enabled)
            {
                this.lbl_doubleQuoteCheck.Text = "✘";
                this.lbl_doubleQuoteCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_doubleQuoteCheck.Text = "✔";
                this.lbl_doubleQuoteCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (this.btn_commaEdit.Enabled)
            {
                this.lbl_commaCheck.Text = "✘";
                this.lbl_commaCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_commaCheck.Text = "✔";
                this.lbl_commaCheck.ForeColor = System.Drawing.Color.Green;
            }
            Ribbon1.showCheckAllUC.setButtonClickALL();
        }

        public void editAll()
        {
            visibleAllClose();
            progressBar.Value = 0;
            progressBar.Increment(1);
            editFullStop();
            progressBar.Increment(1);
            editComma();
            progressBar.Increment(1);
            editColon();
            progressBar.Increment(1);
            editSemiColon();
            progressBar.Increment(1);
            editDQ();
            progressBar.Increment(1);
            visibleAllOpen();
            progressBar.Increment(1);
            checkFinish();


        }


        /*
         *   =================================================================================================================
         *   =================================================================================================================
         *   =================================================================================================================

 
                        ##    ##     ######     ## ####      ##     ##    ## #####    #######    ##########
                        ##    ##    ##    ##    ##     ##    ##     ##    ##         ###     #       ##
                        ##    ##    ##    ##    ##     ##    ##     ##    ##         ###             ##
                        ## #####    ## #####    ## ###       ##     ##    ## ###       #####         ##
                        ##    ##    ##    ##    ##   ##      ##     ##    ##               ###       ##
                        ##    ##    ##    ##    ##    ##      ##   ##     ##         #     ###       ##
                        ##    ##    ##    ##    ##     ##       ###       ## #####    #######        ##
 
         * 
         *   =================================================================================================================
         *   =================================================================================================================
         *   =================================================================================================================
        */

        public void HarvestSCL()
        {

            rSCL.Clear();
            rSCL_Worng.Clear();
            rSCL = FB.findRange(";", Globals.ThisAddIn.Application.ActiveDocument.Content);
            rSCL = Exclude(rSCL, ";");
        }

        public void HarvestCL()
        {

            rCL.Clear();
            rCL_Worng.Clear();
            rCL = FB.findRange(":", Globals.ThisAddIn.Application.ActiveDocument.Content);
            rCL = Exclude(rCL, ":");
        }

        public void HarvestDQ()
        {

            rDQ.Clear();
            rDQ_Worng2.Clear();
            rDQ = FB.findRange2("\"", Globals.ThisAddIn.Application.ActiveDocument.Content);
            rDQ = Exclude(rDQ, "\"");
        }

        public void HarvestFS()
        {

            rFS.Clear();
            rFS_Worng.Clear();
            rFS2.Clear();

            rFS = FB.findRange2("^p", Globals.ThisAddIn.Application.ActiveDocument.Content);
            rFS.AddRange(FB.findRange2("^l", Globals.ThisAddIn.Application.ActiveDocument.Content));

            rFS = Exclude2(rFS);

            String r = "";
            Word.Range rang = null;
            Word.Range rang2 = null;
            int start = 0;
            int end = 0;
            bool loop = false;
            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;

            for (int i = 0; i < rFS.Count; ++i)
            {

                loop = true;
                start = rFS[i].Start;
                end = rFS[i].End;

                if (end + 1 < max)
                {
                    start = end;
                    end = end + 1;
                }
                else
                {
                }

                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);

                r = rang.Characters.Last.Text;

                while (loop)
                {
                    if (r == null)
                    {
                        loop = false;
                    }
                    else
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(r, "\\d"))
                        {
                            rFS2.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end));
                            loop = false;
                        }
                        else
                        {
                            if (FB.find3("^w", rang2) || r == " ")
                            {
                                if (end + 1 < max)
                                {
                                    end = end + 1;
                                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                                    r = rang.Characters.Last.Text;
                                }
                                else
                                {
                                    loop = false;
                                }
                            }
                            else
                            {
                                loop = false;
                            }
                        }
                    }
                }
            }
        }

        public void HarvestCM()
        {

            rCM.Clear();
            rCM_Worng.Clear();
            rCM = FB.findRange(",", Globals.ThisAddIn.Application.ActiveDocument.Content);
            rCM = Exclude(rCM, ",");
        }


        /*
         *   =================================================================================================================
         *   =================================================================================================================
         *   =================================================================================================================
         *   
         * 
                                              ######    ##    ##    ## #####     ######    ##     ##
                                             ##     #   ##    ##    ##          ##     #   ##   ##
                                            ##          ##    ##    ##         ##          ##  ##
                                            ##          ## #####    ## ####    ##          ## #
                                            ##          ##    ##    ##         ##          ##  ##
                                             ##     #   ##    ##    ##          ##     #   ##   ##  
                                              ######    ##    ##    ## #####     ######    ##     ##
                  
         * 
         *   =================================================================================================================
         *   =================================================================================================================
         *   =================================================================================================================
         */


        public void CheckColon()
        {
            HarvestCL();
            //rCL_Worng.Clear();
            // rCL_Bank.Clear();

            String l = "";
            String r = "";

            Word.Range rang = null;
            Word.Range rang2 = null;
            Word.Range rang3 = null;
            int start = 0;
            int end = 0;
            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            int begin = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;

            for (int i = 0; i < rCL.Count; ++i)
            {
                //rCL[i].Select();
                start = rCL[i].Start - 1;
                end = rCL[i].End + 1;
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);
                rang3 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;

                if (FB.find3("^w", rang2) || rang2.Text == " ")
                {
                    rCL_Worng.Add(FB.makeRange(rCL[i]));
                }
                else
                {
                    if (!FB.find3("^w", rang3) || rang3.Text != " ")
                    {

                        if (System.Text.RegularExpressions.Regex.IsMatch(r, "\\d") || r == "-" || r == "\\" || r == "/" || r == "-" || r == ":" || !FB.find3("^w", rang3) || !FB.find3("^p", rang3)) { }
                        else
                        {
                            rCL_Worng.Add(FB.makeRange(rCL[i]));
                        }
                    }
                    else
                    {
                        end = end + 1;
                        rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                        rang3 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                        if (FB.find3("^w", rang2) || rang2.Text == " ")
                        {
                            rCL_Worng.Add(FB.makeRange(rCL[i]));
                        }
                        else
                        {
                        }
                    }
                }
            }
            lbl_colonValue.Text = rCL_Worng.Count.ToString() + " -";
        }


        public void CheckSemiColon()
        {
            HarvestSCL();

            String l = "";
            String r = "";
            Word.Range rang = null;
            Word.Range rang2 = null;
            Word.Range rang3 = null;
            int start = 0;
            int end = 0;
            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            int begin = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;

            for (int i = 0; i < rSCL.Count; ++i)
            {

                start = rSCL[i].Start - 1;
                end = rSCL[i].End + 1;

                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);
                rang3 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;

                if (FB.find3("^w", rang2) || rang2.Text == " ")
                {
                    rSCL_Worng.Add(FB.makeRange(rSCL[i]));
                }
                else
                {
                    if ((!FB.find3("^w", rang3)) || rang3.Text != " ")
                    {
                        if (FB.find3("^p", rang3) || FB.find3("^w", rang3) || rang3.Text == " ") { } //reserve for last white space
                        else
                        {
                            rSCL_Worng.Add(FB.makeRange(rSCL[i]));
                        }
                    }
                    else
                    {
                        end = end + 1;
                        rang3 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                        r = rang.Characters.Last.Text;
                        if (FB.find3("^w", rang3) || rang3.Text == " ")
                        {
                            rSCL_Worng.Add(FB.makeRange(rSCL[i]));
                        }
                        else { }
                    }
                }
            }
            lbl_semiColonValue.Text = rSCL_Worng.Count.ToString() + " -";
        }


        public void CheckComma()
        {
            HarvestCM();

            String l = "";
            String r = "";
            Word.Range rang = null;
            Word.Range rang2 = null;
            Word.Range rang3 = null;
            int start = 0;
            int end = 0;
            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            int begin = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;

            for (int i = 0; i < rCM.Count; ++i)
            {
                start = rCM[i].Start - 1;
                end = rCM[i].End + 1;

                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);
                rang3 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;


                if (FB.find3("^w", (Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1))))
                {
                    rCM_Worng.Add(FB.makeRange(rCM[i]));
                }

                else
                {
                    if (r != " ")
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(r, "\\d")) { }
                        else
                        {
                            rCM_Worng.Add(FB.makeRange(rCM[i]));
                        }
                    }
                    else
                    {
                        end = end + 1;
                        rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                        r = rang.Characters.Last.Text;
                        if (r == " ")
                        {
                            rCM_Worng.Add(FB.makeRange(rCM[i]));
                        }
                        else
                        {
                        }
                    }
                }
            }
            lbl_commaValue.Text = rCM_Worng.Count.ToString() + " -";
        }


        public void CheckDQ()
        {
            HarvestDQ();

            Word.Range rang2 = null;
            int start = 0;
            int end = 0;
            int contDQ = 0;

            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            int begin = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;

            for (int i = 0; i < rDQ.Count; ++i)
            {

                if (contDQ == 0)
                {
                    // Open    ..._"XXX
                    start = rDQ[i].Start - 1;
                    end = rDQ[i].End + 1;

                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                    //    v  = _
                    if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                    { //..._"XXX
                        rDQ_Worng2.Add(makeRN(rDQ[i], false));
                    }
                    else
                    {
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);
                        //    v  = new paragraph
                        if (FB.find3("^p", rang2))
                        {  ////..._"XXX

                        }                                                                        //    v !=  _ 
                        else if (!(FB.find3("^w", rang2) || rang2.Characters.First.Text == " "))
                        {   ////..._"XXX
                            rDQ_Worng2.Add(makeRN(rDQ[i], false));
                        }
                        else
                        {                           //  v----  
                            start = start - 1;          //..._"XXX
                            rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);

                            if (FB.find3("^w", rang2) || rang2.Characters.First.Text == " ")
                            {
                                rDQ_Worng2.Add(makeRN(rDQ[i], false));
                            }
                            else
                            {
                            }
                        }
                    }
                    contDQ = 1;
                }
                else
                {
                    //End   XXX"_...

                    start = rDQ[i].Start - 1;
                    end = rDQ[i].End + 1;

                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start + 1);

                    if (FB.find3("^w", rang2) || rang2.Characters.First.Text == " ")
                    {
                        rDQ_Worng2.Add(makeRN(rDQ[i], true));
                    }
                    else
                    {
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                        if (FB.find3("^p", rang2))
                        {

                        }
                        else if (!(FB.find3("^w", rang2) || rang2.Characters.Last.Text == " "))
                        {
                            rDQ_Worng2.Add(makeRN(rDQ[i], true));
                        }
                        else
                        {
                            end = end + 1;
                            rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                            if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                            {
                                rDQ_Worng2.Add(makeRN(rDQ[i], true));
                            }
                            else
                            {
                            }
                        }
                    }
                    contDQ = 0;
                }
            }
            lbl_doubleQuoteValue.Text = rDQ_Worng2.Count.ToString() + " -";
        }



        public void CheckFullStop()
        {
            HarvestFS();
            String l = "";
            String r = "";
            Word.Range rang = null;
            Word.Range rang2 = null;
            int start = 0;
            int end = 0;
            bool loop = false;
            bool head = false;
            bool space = false;

            for (int i = 0; i < rFS2.Count; ++i)
            {


                start = rFS2[i].Start;
                end = rFS2[i].End + 1;


                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, end + 1);
                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);

                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;
                head = false;



                while (loop)
                {
                    if (FB.find2("[0-9]{1,}.", rang))
                    {
                        head = true;
                        start = end;
                        end = end + 1;
                        rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                        l = rang.Characters.First.Text;
                        r = rang.Characters.Last.Text;
                    }
                    else
                    {
                        if (head)
                        {

                            if (FB.find3("^w", rang2) || rang2.Text == " ")
                            {
                                end = end + 1;
                                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                                r = rang.Characters.Last.Text;

                                if (FB.find3("^w", rang2) || rang2.Text == " ")
                                {
                                    end = end + 1;
                                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);

                                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                                    r = rang.Characters.Last.Text;

                                    if ((!FB.find3("^w", rang2)) || rang2.Text != " ")
                                    {
                                        loop = false;
                                    }
                                    else
                                    {
                                        space = true;
                                        while (space)
                                        {
                                            if (FB.find3("^w", rang2) || rang2.Text == " ")
                                            {
                                                end = end + 1;
                                                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                                            }
                                            else
                                            {
                                                space = false;
                                            }
                                        }
                                        rFS_Worng.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rFS2[i].Start, end));
                                        loop = false;
                                    }
                                }
                                else
                                {
                                    rFS_Worng.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rFS2[i].Start, end));
                                    loop = false;
                                }
                            }
                            else
                            {
                                if (FB.find2("[0-9]{1,} ", rang))
                                {
                                    if (FB.find3("^w", rang2) || rang2.Text == " ")
                                    {
                                        end = end + 1;
                                        rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);

                                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);
                                        r = rang.Characters.Last.Text;

                                        if ((!FB.find3("^w", rang2)) || rang2.Text != " ")
                                        {
                                            loop = false;
                                        }
                                        else
                                        {
                                            space = true;
                                            while (space)
                                            {

                                                if (FB.find3("^w", rang2) || rang2.Text == " ")
                                                {
                                                    end = end + 1;
                                                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                                                }
                                                else
                                                {
                                                    space = false;
                                                }
                                            }
                                            rFS_Worng.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rFS2[i].Start, end));
                                            loop = false;
                                        }
                                    }
                                    else
                                    {
                                        rFS_Worng.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rFS2[i].Start, end));
                                        loop = false;
                                    }
                                }
                                else
                                {
                                    end = end + 1;
                                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);

                                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                                    l = rang.Characters.First.Text;
                                    r = rang.Characters.Last.Text;
                                    if (r == "." || System.Text.RegularExpressions.Regex.IsMatch(r, "\\d") || FB.find3("^w", rang2) || rang2.Text == " ") { }
                                    else
                                    {
                                        rFS_Worng.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rFS2[i].Start, end));
                                        loop = false;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (head) { }
                            else
                            {
                                end = end + 1;
                                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                                rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end - 1, end);

                                l = rang.Characters.First.Text;
                                r = rang.Characters.Last.Text;
                                if (r == "." || rang2.Text == " " || System.Text.RegularExpressions.Regex.IsMatch(r, "\\d") || FB.find3("^w", rang2)) { }
                                else { loop = false; }
                            }
                        }
                    }//else 
                }//end while
            } // end for
            lbl_fullStopValue.Text = rFS_Worng.Count.ToString() + " -";
        }


        /*
         *   =================================================================================================================
         *   =================================================================================================================
         *   =================================================================================================================
         *   
         * 
                                            ## #####  ## ####     ######   ##########
                                            ##        ##     ##     ##         ##
                                            ##        ##      ##    ##         ##
                                            ## ###    ##      ##    ##         ##
                                            ##        ##      ##    ##         ##
                                            ##        ##     ##     ##         ##
                                            ## #####  ## ####     ######       ##
         * 
         * 
         *   =================================================================================================================
         *   =================================================================================================================
         */


        public void editComma()
        {
            rCM_Bank.Clear();
            String l = "";
            String r = "";

            Word.Range rang = null;

            int start = 0;
            int end = 0;

            bool loop = false;

            for (int i = 0; i < rCM_Worng.Count; ++i)
            {

                start = rCM_Worng[i].Start;
                end = rCM_Worng[i].Start;

                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;

                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, start);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        start = start - 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = "";

                start = rCM_Worng[i].End;
                end = rCM_Worng[i].End;

                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;

                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end, end + 1);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        end = end + 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = " ";
                rCM_Bank.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rang.Start - 1, rang.End));
            }
            CheckComma();
        }




        public void editColon()
        {
            rCL_Bank.Clear();
            String l = "";
            String r = "";

            Word.Range rang = null;

            int start = 0;
            int end = 0;

            bool loop = false;

            for (int i = 0; i < rCL_Worng.Count; ++i)
            {
                start = rCL_Worng[i].Start;
                end = rCL_Worng[i].Start;
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;

                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, start);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        start = start - 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = "";
                start = rCL_Worng[i].End;
                end = rCL_Worng[i].End;
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;
                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end, end + 1);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        end = end + 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = " ";
                rCL_Bank.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rang.Start - 1, rang.End));
            }
            CheckColon();
        }

        public void editSemiColon()
        {
            rSCL_Bank.Clear();
            String l = "";
            String r = "";

            Word.Range rang = null;

            int start = 0;
            int end = 0;

            bool loop = false;

            for (int i = 0; i < rSCL_Worng.Count; ++i)
            {

                start = rSCL_Worng[i].Start;
                end = rSCL_Worng[i].Start;
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;

                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, start);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        start = start - 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = "";
                start = rSCL_Worng[i].End;
                end = rSCL_Worng[i].End;
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                l = rang.Characters.First.Text;
                r = rang.Characters.Last.Text;
                loop = true;

                while (loop)
                {
                    rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end, end + 1);
                    if (FB.find3("^w", rang) || rang.Text == " ")
                    {
                        end = end + 1;
                    }
                    else
                    {
                        loop = false;
                    }
                }
                rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
                rang.Text = " ";
                rSCL_Bank.Add(Globals.ThisAddIn.Application.ActiveDocument.Range(rang.Start - 1, rang.End));
            }
            CheckSemiColon();
        }



        public void editFullStop()
        {
            int start = 0;
            int end = 0;
            Word.Range rang = null;
            for (int i = 0; i < rFS_Worng.Count; ++i)
            {

                if (FB.find3("^w", rFS_Worng[i]) == true)
                {
                    // rFS_Worng[i].Text = "  ";
                    // FB.findReplace("( ){1,}","  ",rFS_Worng[i].Find);
                    FB.findReplace2("^w", "  ", rFS_Worng[i].Find);
                }

                else
                {/*  
                    for(int j = 0 ; j< rFS_Worng[i].Characters.Count ; ++j){
                     start = rFS_Worng[i].Start;
                       end = rFS_Worng[i].End-j;
                      rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start,end);
                      if(System.Text.RegularExpressions.Regex.IsMatch(rang.Characters.Last.Text, "\\d")){
  
                          rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end,end);
                          rang.Text = "  ";
                          break;
                      } */
                }
            }
            CheckFullStop();
        }






        public void editDQ()
        {
            Word.Range rang = null;
            List<Word.Range> rangl = new List<Word.Range>();
            Word.Range rang2 = null;
            int start = 0;
            int start2 = 0;
            int end = 0;
            int end2 = 0;

            int max = Globals.ThisAddIn.Application.ActiveDocument.Content.End;
            int begin = Globals.ThisAddIn.Application.ActiveDocument.Content.Start;
            bool loop = false;

            for (int i = 0; i < rDQ_Worng2.Count; ++i)
            {
                //rDQ_Worng2[i].rang.Select();
                start = rDQ_Worng2[i].rang.Start;
                start2 = rDQ_Worng2[i].rang.Start;
                end = rDQ_Worng2[i].rang.End;
                end2 = rDQ_Worng2[i].rang.End;
                loop = true;

                if (rDQ_Worng2[i].pos == false)
                {// Open Qute

                    while (loop)
                    {
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, start);
                        if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                        {
                            start = start - 1;
                        }
                        else
                        {
                            rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start2);
                            rang.Text = " ";
                            loop = false;
                        }
                    } loop = true;

                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(rDQ_Worng2[i].rang.Start - 1, rDQ_Worng2[i].rang.End + 1);

                    FB.find31("\"", rang2.Find);

                    end2 = rang2.End;
                    end = rang2.End;
                    // rang2.Select();

                    while (loop)
                    {

                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end, end + 1);
                        if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                        {
                            end = end + 1;
                        }
                        else
                        {
                            rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end2, end);
                            rang.Text = "";
                            loop = false;
                        }
                    }
                }
                else
                { // end Qute
                    loop = true;
                    while (loop)
                    {
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(start - 1, start);
                        if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                        {
                            start = start - 1;
                        }
                        else
                        {
                            rang = Globals.ThisAddIn.Application.ActiveDocument.Range(start, start2);
                            rang.Text = "";
                            loop = false;
                        }
                    }
                    loop = true;
                    rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(rDQ_Worng2[i].rang.Start - 1, rDQ_Worng2[i].rang.End + 1);

                    FB.find31("\"", rang2.Find);

                    end = rang2.End;
                    end2 = rang2.End;
                    while (loop)
                    {
                        //right
                        rang2 = Globals.ThisAddIn.Application.ActiveDocument.Range(end, end + 1);
                        if (FB.find3("^w", rang2) || rang2.Characters.Last.Text == " ")
                        {
                            end = end + 1;
                        }
                        else
                        {
                            rang = Globals.ThisAddIn.Application.ActiveDocument.Range(end2, end);
                            rang.Text = " ";
                            loop = false;
                        }
                    }
                }
            }
            CheckDQ();
        }

        private void btn_fullStopEdit_Click(object sender, EventArgs e)
        {
            visibleAllClose();
            editFullStop();
            visibleAllOpen();
            checkFinish();
        }


        private void btn_commaEdit_Click(object sender, EventArgs e)
        {
            visibleAllClose();
            editComma();
            visibleAllOpen();
            checkFinish();
        }


        private void btn_colonEdit_Click(object sender, EventArgs e)
        {
            visibleAllClose();
            editColon();
            visibleAllOpen();
            checkFinish();
        }

        private void btn_semiColonEdit_Click(object sender, EventArgs e)
        {
            visibleAllClose();
            editSemiColon();
            visibleAllOpen();
            checkFinish();
        }

        private void btn_doubleQuoteEdit_Click(object sender, EventArgs e)
        {
            visibleAllClose();
            editDQ();
            visibleAllOpen();
            checkFinish();
        }

        private void btn_editAll_Click(object sender, EventArgs e)
        {

            //lbl_check.Text = "Editing ";
            editAll();
            //lbl_check.Text = "Edit Finish";
        }




        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            index_of_Selection = -1;
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            try
            {


                if (index_of_Selection > 0)
                {
                    if (this.radioButton1.Checked)
                    {
                        //index_of_Selection = index_of_Selection - 1;
                        index_of_Selection = index_of_Selection % rFS_Worng.Count;
                        rFS_Worng[--index_of_Selection].Select();

                    }
                    else if (this.radioButton2.Checked)
                    {
                        //index_of_Selection = index_of_Selection - 1;
                        index_of_Selection = index_of_Selection % rCM_Worng.Count;
                        rCM_Worng[--index_of_Selection].Select();

                    }
                    else if (this.radioButton3.Checked)
                    {

                        //index_of_Selection = index_of_Selection - 1;
                        index_of_Selection = index_of_Selection % rCL_Worng.Count;
                        rCL_Worng[--index_of_Selection].Select();


                    }
                    else if (this.radioButton4.Checked)
                    {
                        // index_of_Selection = index_of_Selection - 1;
                        index_of_Selection = index_of_Selection % rSCL_Worng.Count;
                        rSCL_Worng[--index_of_Selection].Select();

                    }
                    else if (this.radioButton5.Checked)
                    {
                        //index_of_Selection = index_of_Selection - 1;
                        index_of_Selection = index_of_Selection % rDQ_Worng2.Count;
                        rDQ_Worng2[--index_of_Selection].rang.Select();

                    }
                    else { }
                }

                label3.Text = (index_of_Selection + 1).ToString();
            }
            catch (Exception) { }
        }


        private void btn_next_Click(object sender, EventArgs e)
        {
            try
            {

                if (this.radioButton1.Checked)
                {
                    rFS_Worng[++index_of_Selection % rFS_Worng.Count].Select();
                    index_of_Selection = index_of_Selection % rFS_Worng.Count;
                }

                else if (this.radioButton2.Checked)
                {
                    rCM_Worng[++index_of_Selection % rCM_Worng.Count].Select();
                    index_of_Selection = index_of_Selection % rCM_Worng.Count;
                }

                else if (this.radioButton3.Checked)
                {
                    rCL_Worng[++index_of_Selection % rCL_Worng.Count].Select();
                    index_of_Selection = index_of_Selection % rCL_Worng.Count;
                }
                else if (this.radioButton4.Checked)
                {
                    rSCL_Worng[++index_of_Selection % rSCL_Worng.Count].Select();
                    index_of_Selection = index_of_Selection % rSCL_Worng.Count;
                }

                else if (this.radioButton5.Checked)
                {
                    rDQ_Worng2[++index_of_Selection % rDQ_Worng2.Count].rang.Select();
                    index_of_Selection = index_of_Selection % rDQ_Worng2.Count;
                }

                else { }
            }
            catch (Exception) { }
            label3.Text = (index_of_Selection + 1).ToString();

        }

        private void btn_highlightALL_Click(object sender, EventArgs e)
        {
            hightlight(rCL_Bank, Word.WdColorIndex.wdRed);
            hightlight(rDQ_Worng2, Word.WdColorIndex.wdPink);
            hightlight(rSCL_Bank, Word.WdColorIndex.wdGreen);
            hightlight(rCM_Bank, Word.WdColorIndex.wdYellow);
            hightlight(rFS_Worng, Word.WdColorIndex.wdBrightGreen);


        }

        private void btn_clearHighlightALL_Click(object sender, EventArgs e)
        {
            hightlight(rCL_Bank, Word.WdColorIndex.wdNoHighlight);
            hightlight(rDQ_Worng2, Word.WdColorIndex.wdNoHighlight);
            hightlight(rSCL_Bank, Word.WdColorIndex.wdNoHighlight);
            hightlight(rCM_Bank, Word.WdColorIndex.wdNoHighlight);
            hightlight(rFS_Worng, Word.WdColorIndex.wdNoHighlight);
        }

        private void btn_check_Click(object sender, EventArgs e)
        {
            this.checkAll();
        }



    }
}
