using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Drawing;
using System.Data;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Globalization;

namespace SeniorProject
{
    public partial class MarginPageUC : UserControl
    {
        public bool checkSetClick;
        public MarginPageUC()
        {
            checkSetClick = false;
            InitializeComponent();
        }

       


        public void setMarginPageUC(bool value,bool leftCheck, bool rightCheck, bool topCheck, bool bottomCheck,
            float left, float right, float top, float bottom)
        {
            this.lbl_leftMarginValue.Text = (left * 0.0138888889).ToString("0.00", CultureInfo.InvariantCulture) + " นิ้ว";
            if (leftCheck)
            {
                this.lbl_leftMarginCheck.Text = "✔";
                this.lbl_leftMarginCheck.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                this.lbl_leftMarginCheck.Text = "✘";
                this.lbl_leftMarginCheck.ForeColor = System.Drawing.Color.Red;
            }
            this.lbl_rightMarginValue.Text = (right * 0.0138888889).ToString("0.00", CultureInfo.InvariantCulture) + " นิ้ว";
            if (rightCheck)
            {
                this.lbl_rightMarginCheck.Text = "✔";
                this.lbl_rightMarginCheck.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                this.lbl_rightMarginCheck.Text = "✘";
                this.lbl_rightMarginCheck.ForeColor = System.Drawing.Color.Red;
            }
            this.lbl_topMarginValue.Text = (top * 0.0138888889).ToString("0.00", CultureInfo.InvariantCulture) + " นิ้ว";
            if (topCheck)
            {
                this.lbl_topMarginCheck.Text = "✔";
                this.lbl_topMarginCheck.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                this.lbl_topMarginCheck.Text = "✘";
                this.lbl_topMarginCheck.ForeColor = System.Drawing.Color.Red;
            }
            this.lbl_bottomMarginValue.Text = (bottom * 0.0138888889).ToString("0.00", CultureInfo.InvariantCulture) + " นิ้ว";
            if (bottomCheck)
            {
                this.lbl_bottomMarginCheck.Text = "✔";
                this.lbl_bottomMarginCheck.ForeColor = System.Drawing.Color.Green;
            }
            else
            {
                this.lbl_bottomMarginCheck.Text = "✘";
                this.lbl_bottomMarginCheck.ForeColor = System.Drawing.Color.Red;
            }

          //  if (cheked())
           // {
            this.btn_Edit.Enabled = !value;
           // }
        }

    }
}
