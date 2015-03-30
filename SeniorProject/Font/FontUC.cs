using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SeniorProject
{
    public partial class FontUC : UserControl
    {
        FontManager fm;
        StyleFont sfMain;
        public FontUC()
        {
            fm = new FontManager();
            InitializeComponent();
            this.pnlSizeCheck.Visible = false;
        }

        public void enableAll()
        {
            btnEdit.Visible = false;
            btn_lookError.Visible = false;
            pgbFontName.Value = pgbFontName.Minimum;
            pgbFontSize.Value = pgbFontSize.Minimum;
            lblFontFault.Text = "---";
            lblSizeFault.Text = "---";
            cbx_fontName.Items.Clear();
            lbl_facultyValue.Text = Ribbon1.referenceModel.faculty;
            setCbxFontType();
            pnlSizeCheck.Visible = false;
            lblMainText.Text = "ข้อผิดพลาด";
            /*btn_check.Enabled = true;
            listBox1.Enabled = false;
            btn_next.Enabled = false;
            btn_highlightALL.Enabled = false;

            btn_clearHighlightALL.Enabled = false;
            progressBar.Value = 0;
            lbl_numberValueError.Text = "-";
            progressBar.Enabled = false;
            lbl_waitCheck.Enabled = false;*/

        }

        private void btnCheck_Click(object sender, EventArgs e)
        {
            //fm.checkFontName("Angsana New", this);
            checkFontName();
        }

        public void checkFontName()
        {
            fm.checkFontName(sfMain.FontName, this);
            Ribbon1.showCheckAllUC.setButtonClickALL();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            correctFont();
        }

        public void correctFont()
        {
            fm.CorrectFont(sfMain.FontName, this);
            checkFontName();
        }

        private void btnFontSizeCheck_Click(object sender, EventArgs e)
        {
            FontSizeCheck();
        }

        public void FontSizeCheck()
        {
            fm.CheckFontSize((int)sfMain.Substance,
                (int)sfMain.Subheading,
                (int)sfMain.Topics,
                (int)sfMain.Namechapter,
                (int)sfMain.Chapter, this);
            Ribbon1.showCheckAllUC.setButtonClickALL();
        }


        private void btnCorrectNext_Click(object sender, EventArgs e)
        {
            fm.IndexFaultFontSize(this);
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.pnlSizeCheck.Visible = true;
        }

        private void cbx_fontType_SelectedIndexChanged(object sender, EventArgs e)
        {
            sfMain = Ribbon1.styles.StyleFont[cbx_fontName.SelectedIndex];
            lbl_chapterValue.Text = "" + (int)sfMain.Chapter;
            lbl_nameChapterValue.Text = "" + (int)sfMain.Namechapter;
            lbl_topicValue.Text = "" + (int)sfMain.Topics;
            lbl_supheadingValue.Text = "" + (int)sfMain.Subheading;
            lbl_substanceValue.Text = "" + (int)sfMain.Substance;
        }

        public void setCbxFontType()
        {
            //int countFirst = 0;
            foreach (StyleFont sf in Ribbon1.styles.StyleFont)
            {
                cbx_fontName.Items.Add(sf.FontName+" ["+sf.FontNameLanguage+"]");
               /* if (countFirst == 0)
                {
                    sfMain = Ribbon1.styles.StyleFont[0];
                    countFirst++;
                }*/
            }
            cbx_fontName.SelectedIndex = 0;
        }

        private void btnClearHightLight_Click(object sender, EventArgs e)
        {
            fm.ClearHightLight();
        }
    }
}
