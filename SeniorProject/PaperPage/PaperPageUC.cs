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
    public partial class PaperPageUC : UserControl
    {
        public bool checkSetClick;
        public PaperPageUC()
        {
            checkSetClick = false;
            InitializeComponent();
        }

        public void setPaperPageUC(bool paperTypeCheck,
     string paperType)
        {
            paperType = paperType.Substring(7);
            this.lbl_paperType.Text = paperType;
            if (paperTypeCheck)
            {
                this.lbl_paperTypeCheck.Text = "✔";
                this.lbl_paperTypeCheck.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                this.lbl_paperTypeCheck.Text = "✘";
                this.lbl_paperTypeCheck.ForeColor = System.Drawing.Color.Red;
            }
            this.lbl_paperTypeValue.Text = "กระดาษ: " +paperType;
            //  if (cheked())
            // {
            this.btn_Edit.Enabled = !paperTypeCheck;
            // }
        }
    }
}
