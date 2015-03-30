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
    public partial class ReferenceModelUC : UserControl
    {
        public bool checkSetClick;
        List<RangeErrorForReference> rangeErrorForReferences;
        List<RangeErrorForReference> rangeErrorForReferencesSorts;
        public List<int> numberReferenceErrorSortYearName;
        public int index;
        public ReferenceModelUC()
        {
            checkSetClick = false;
            rangeErrorForReferences = new List<RangeErrorForReference>();
            numberReferenceErrorSortYearName = new List<int>();
            index = -1;
            InitializeComponent();
        }


        public void AddRangeErrorForReference(int numberReferenceError, Word.Range rangeReferenceError,int typeError)
        {
            foreach (RangeErrorForReference refr in this.rangeErrorForReferences)
            {
                if (numberReferenceError == refr.getNumberReferenceError())
                {
                    refr.setRangeError(typeError);
                    return;
                }
            }
            rangeErrorForReferences.Add(new RangeErrorForReference(numberReferenceError, rangeReferenceError, typeError));
            //List<RangeErrorForReference> rangeErrorForReferencesSort = new List<RangeErrorForReference>(rangeErrorForReferences.OrderBy(rangeErrorForReference => rangeErrorForReference.getNumberReferenceError()).ToArray());
        }

        public void Clear()
        {
            this.tbx_referenceAllError.ForeColor = System.Drawing.Color.Red;
            numberReferenceErrorSortYearName.Clear();
            rangeErrorForReferences.Clear();
            //rangeErrorForReferencesSorts.Clear();
            index = -1;
        }

        public void ShowErrorNumber()
        {
            btn_edit.Text = "แก้ไขระบบลำดับหมายเลขผิด";
            lbl_referenceNameAndYear.Text = "ระบบลำดับหมายเลขผิด: ";
            ShowError();
        }

        public void setErrorNull()
        {
            lbl_referenceAll.Visible = false;
            lbl_referenceAllCheck.Visible = false;
            lbl_referenceAllError.Visible = false;
            tbx_referenceAllError.Visible = false;
            lbl_numberReference.Visible = false;
            lbl_numberReferenceValue.Visible = false;
            btn_back.Visible = false;
            btn_next.Visible = false;
            lbl_reference.Visible = false;
            lbl_referenceNameAndYear.Visible = false;
            lbl_margin.Visible = false;
            lbl_referenceCheck.Visible = false;
            lbl_referenceNameAndYearCheck.Visible = false;
            lbl_marginCheck.Visible = false;
            btn_edit.Visible = false;
            lbl_Error.Visible = true;
        }

        public void visibledefault()
        {
            lbl_referenceAll.Visible = true;
            lbl_referenceAllCheck.Visible = true;
            lbl_referenceAllError.Visible = true;
            tbx_referenceAllError.Visible = true;
            lbl_numberReference.Visible = true;
            lbl_numberReferenceValue.Visible = true;
            btn_back.Visible = true;
            btn_next.Visible = true;
            lbl_reference.Visible = true;
            lbl_referenceNameAndYear.Visible = true;
            lbl_margin.Visible = true;
            lbl_referenceCheck.Visible = true;
            lbl_referenceNameAndYearCheck.Visible = true;
            lbl_marginCheck.Visible = true;
            btn_edit.Visible = true;
            lbl_Error.Visible = false;
        }

        public void ShowError()
        {

                this.lbl_marginCheck.Text = "✔";
                this.lbl_marginCheck.ForeColor = System.Drawing.Color.Green;


                this.lbl_referenceCheck.Text = "✔";
                this.lbl_referenceCheck.ForeColor = System.Drawing.Color.Green;

                this.lbl_referenceNameAndYearCheck.Text = "✔";
                this.lbl_referenceNameAndYearCheck.ForeColor = System.Drawing.Color.Green;

            this.rangeErrorForReferencesSorts = new List<RangeErrorForReference>(this.rangeErrorForReferences.OrderBy(rangeErrorForReference => rangeErrorForReference.getNumberReferenceError()).ToArray());
            if (this.rangeErrorForReferencesSorts.Count == 0)
            {
                tbx_referenceAllError.Text = "ผ่าน";
                tbx_referenceAllError.Font = new System.Drawing.Font("Angsana New", 16.0F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
                tbx_referenceAllError.ForeColor = System.Drawing.Color.Green;
                this.btn_back.Enabled = false;
                this.btn_next.Enabled = false;
                this.btn_edit.Enabled = false;
                return;
            }
            tbx_referenceAllError.Text = "";
            


            foreach (RangeErrorForReference rangeErrorForReferencesSort in this.rangeErrorForReferencesSorts)
            {
                tbx_referenceAllError.Text += rangeErrorForReferencesSort.getNumberReferenceError() + ", ";
            }

           // if (numberReferenceErrorSortYearName.Count == 0 &&
           //     numberReferenceErrorSortName.Count == 0)
           // {
                this.btn_edit.Enabled = checkBtnEditEnabled();
           // }
        }

        private bool checkBtnEditEnabled()
        {
            bool value = false;
            foreach (RangeErrorForReference rangeErrorForReferencesSort in this.rangeErrorForReferencesSorts)
            {
                value = value || rangeErrorForReferencesSort.checkBtnEditEnabled();
            }

            return value;
        }
        
        private void selectNext()
        {
            if (index == -1)
            {
                index = 0;
                //this.rangeErrorForReferencesSorts[index].getRangeReferenceError().Select();
                
            }
            else
            {
                index++;
                if (index == this.rangeErrorForReferencesSorts.Count)
                {
                    index = 0;
                }
                
            }
            this.rangeErrorForReferencesSorts[index].getRangeReferenceError().Select();
            this.setCheck(this.rangeErrorForReferencesSorts[index]);
            this.lbl_numberReferenceValue.Text = (index+1) + "";
        }

        private void setCheck(RangeErrorForReference rangeErrorForReferencesetCheck)
        {
            if (rangeErrorForReferencesetCheck.rangeErrorMargin)
            {
                this.lbl_marginCheck.Text = "✘";
                this.lbl_marginCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_marginCheck.Text = "✔";
                this.lbl_marginCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (rangeErrorForReferencesetCheck.rangeErrorModel)
            {
                this.lbl_referenceCheck.Text = "✘";
                this.lbl_referenceCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_referenceCheck.Text = "✔";
                this.lbl_referenceCheck.ForeColor = System.Drawing.Color.Green;
            }
            if (rangeErrorForReferencesetCheck.rangeErrorNameYear)
            {
                this.lbl_referenceNameAndYearCheck.Text = "✘";
                this.lbl_referenceNameAndYearCheck.ForeColor = System.Drawing.Color.Red;
            }
            else
            {
                this.lbl_referenceNameAndYearCheck.Text = "✔";
                this.lbl_referenceNameAndYearCheck.ForeColor = System.Drawing.Color.Green;
            }
        }

        private void selectBack()
        {
            if (index == -1)
            {
                index = 0;
            }
            else
            {
                index--;
                if (index == -1)
                {
                    index = this.rangeErrorForReferencesSorts.Count - 1;
                }
            }
            this.rangeErrorForReferencesSorts[index].getRangeReferenceError().Select();
            this.lbl_numberReferenceValue.Text = (index + 1) + "";
            this.setCheck(this.rangeErrorForReferencesSorts[index]);
        }

        public void setSortYear(List<Word.Range> listRange)
        {
            if (numberReferenceErrorSortYearName.Count > 0)
            {
                //int i = 1;
                foreach (int number in numberReferenceErrorSortYearName)
                {
                    //rangeReferenceErrorSortYear.Add(listRange[number - 1]);
                    this.AddRangeErrorForReference(number, listRange[number - 1], 2);
                }
            }
        }

        private void btn_next_Click(object sender, EventArgs e)
        {
            selectNext();
        }

        private void btn_back_Click(object sender, EventArgs e)
        {
            selectBack();
        }

        
        
    }
}
