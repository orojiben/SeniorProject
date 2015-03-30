using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;

namespace SeniorProject
{
    public partial class ShowCheckAllUC : UserControl
    {
        public MarginPage marginPage;
        public PaperPage paperPage;
        public ReferenceModel referenceModel;
        public Punctuation punctuation;
        public RoyalWord royalWord;
        public ShowCheckAllUC(MarginPage marginPage, PaperPage paperPage, ReferenceModel referenceModel, Punctuation punctuation, RoyalWord royalWord)
        {
            InitializeComponent();
            this.marginPage = marginPage;
            this.paperPage = paperPage;
            this.referenceModel = referenceModel;
            this.punctuation = punctuation;
            this.royalWord = royalWord;
        }

        public void clearAll()
        {
            this.marginPage = null;
            this.paperPage = null;
            this.referenceModel = null;
        }

        public void visibleAllOpen()
        {
            pnl_main.Enabled = true;
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;
        }

        public void visibleAllClose()
        {
            pnl_main.Enabled = false;
            progressBar.Enabled = true;
            lbl_waitCheck.Enabled = true;
        }

        public void resetAll()
        {
            this.lbl_marginCheck.Text = "✘";
            this.lbl_marginCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_paperPageCheck.Text = "✘";
            this.lbl_paperPageCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_fontTypeCheck.Text = "✘";
            this.lbl_fontTypeCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_fontSizeCheck.Text = "✘";
            this.lbl_fontSizeCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_checkRoyalWordCheck.Text = "✘";
            this.lbl_checkRoyalWordCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_checkPunctuationMarkCheck.Text = "✘";
            this.lbl_checkPunctuationMarkCheck.ForeColor = System.Drawing.Color.Red;
            this.lbl_checkReferenceCheck.Text = "✘";
            this.lbl_checkReferenceCheck.ForeColor = System.Drawing.Color.Red;
            this.progressBar.Value = 0;
        }

        public void setUC(MarginPage marginPage, PaperPage paperPage, ReferenceModel referenceModel, Punctuation punctuation,RoyalWord royalWord)
        {
            this.marginPage = marginPage;
            this.paperPage = paperPage;
            this.referenceModel = referenceModel;
            this.punctuation = punctuation;
            this.royalWord = royalWord;
            
        }

        public void setButtonClickALL()
        {
            this.btn_margin.Enabled = Ribbon1.marginPageUC.btn_Edit.Enabled;
            if (Ribbon1.marginPageUC.btn_Edit.Enabled)
            {
                this.lbl_marginCheck.Text = "✘";
                this.lbl_marginCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_marginCheck.Text = "✔";
                this.lbl_marginCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_paperPage.Enabled = Ribbon1.paperPageUC.btn_Edit.Enabled;
            if (Ribbon1.paperPageUC.btn_Edit.Enabled)
            {
                this.lbl_paperPageCheck.Text = "✘";
                this.lbl_paperPageCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_paperPageCheck.Text = "✔";
                this.lbl_paperPageCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_fontType.Enabled = Ribbon1.fontUC.btnEdit.Visible;
            if (Ribbon1.fontUC.btnEdit.Visible)
            {
                this.lbl_fontTypeCheck.Text = "✘";
                this.lbl_fontTypeCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_fontTypeCheck.Text = "✔";
                this.lbl_fontTypeCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_fontSize.Enabled = Ribbon1.fontUC.btn_lookError.Visible;
            if (Ribbon1.fontUC.btn_lookError.Visible)
            {
                this.lbl_fontSizeCheck.Text = "✘";
                this.lbl_fontSizeCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_fontSizeCheck.Text = "✔";
                this.lbl_fontSizeCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_checkRoyalWord.Enabled = Ribbon1.royalWordUC.btn_fullStopEdit.Enabled;
            if (Ribbon1.royalWordUC.btn_fullStopEdit.Enabled)
            {
                this.lbl_checkRoyalWordCheck.Text = "✘";
                this.lbl_checkRoyalWordCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_checkRoyalWordCheck.Text = "✔";
                this.lbl_checkRoyalWordCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_checkPunctuationMark.Enabled = Ribbon1.punctuationUC.btn_editAll.Enabled;
            if (Ribbon1.punctuationUC.btn_editAll.Enabled)
            {
                this.lbl_checkPunctuationMarkCheck.Text = "✘";
                this.lbl_checkPunctuationMarkCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_checkPunctuationMarkCheck.Text = "✔";
                this.lbl_checkPunctuationMarkCheck.ForeColor = System.Drawing.Color.Green;
            }
            this.btn_checkReference.Enabled = Ribbon1.referenceModelUC.btn_edit.Enabled;
            if (Ribbon1.referenceModelUC.btn_edit.Enabled)
            {
                this.lbl_checkReferenceCheck.Text = "✘";
                this.lbl_checkReferenceCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_checkReferenceCheck.Text = "✔";
                this.lbl_checkReferenceCheck.ForeColor = System.Drawing.Color.Green;
            }
        }

        private void btn_margin_Click(object sender, EventArgs e)
        {
            this.marginPage.showForAll();
        }

        private void btn_paperPage_Click(object sender, EventArgs e)
        {
            this.paperPage.showForAll();
        }

        private void btn_fontType_Click(object sender, EventArgs e)
        {
            Ribbon1.showCustomTaskPane(4, true);
        }

        private void btn_fontSize_Click(object sender, EventArgs e)
        {
            Ribbon1.showCustomTaskPane(4, true);
        }

        private void btn_checkRoyalWord_Click(object sender, EventArgs e)
        {
            this.royalWord.showForAll();
        }

        private void btn_checkPunctuationMark_Click(object sender, EventArgs e)
        {
            this.punctuation.showForAll();
        }

        private void btn_checkReference_Click(object sender, EventArgs e)
        {
            this.referenceModel.showForAll();
        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            this.progressBar.Maximum = 14;
            this.progressBar.Value = 0;
            visibleAllClose();
            this.editAll();
        }

        public void checkAll()
        {
          //  try
           // {
               
                
                //this.clearAll();
                marginPage.chekingNotShow();
                progressBar.Increment(1);
                Thread.Sleep(500);
               // Ribbon1.showCheckAllUC.progressBar.Value = 20;
               // Ribbon1.showCheckAllUC.lbl_ps.Text = "20 %";
                //bool marginPageCheck = Ribbon1.marginPageUC.btn_Edit.Enabled;
                paperPage.chekingNotShow();
                progressBar.Increment(1);
                Thread.Sleep(500);
                Ribbon1.fontUC.enableAll();
                Ribbon1.fontUC.checkFontName();
                progressBar.Increment(1);
                Thread.Sleep(500);
                Ribbon1.fontUC.FontSizeCheck();
                progressBar.Increment(1);
                Thread.Sleep(500);
                Ribbon1.royalWordUC.checkWordAll();
                progressBar.Increment(1);
                Thread.Sleep(500);

                Ribbon1.punctuationUC.checkAll();
                progressBar.Increment(1);
                Thread.Sleep(500);
                //bool paperPageCheck = paperPage.paperPageUC.btn_Edit.Enabled;
                //this.ShowFont();
                referenceModel.showUC = false;
                referenceModel.runCheckReferenceAll();
                progressBar.Increment(1);
                Thread.Sleep(500);
                setButtonClickALL();
                visibleAllOpen();
            //}
           // catch
           // {
            //    Ribbon1.showCustomTaskPane();
            //};
        }

        private void editAll()
        {
            Ribbon1.saveFileAuto();
            
            progressBar.Increment(1);
            Thread.Sleep(500);
            this.paperPage.changing();
            progressBar.Increment(1);
            Thread.Sleep(500);
            Ribbon1.fontUC.correctFont();
            progressBar.Increment(1);
            Thread.Sleep(500);
            Ribbon1.royalWordUC.editWordAllForAll();
            progressBar.Increment(1);
            Thread.Sleep(500);
            
           
          
           
            Ribbon1.punctuationUC.editAll();
            progressBar.Increment(1);
            Thread.Sleep(500);
            this.referenceModel.runEditReferenceAll();
            progressBar.Increment(1);
            Thread.Sleep(500);
            this.marginPage.changing();
            progressBar.Increment(1);
            Thread.Sleep(500);
            checkAll();
            
        }

        private void btn_check_Click(object sender, EventArgs e)
        {
            this.progressBar.Maximum = 7;
            this.progressBar.Value = 0;
            visibleAllClose();
            this.checkAll();
        }
    }
}
