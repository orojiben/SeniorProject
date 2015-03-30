using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeniorProject
{
    public class ReferenceModel
    {
        public string faculty;
        public string department;
        private Word.Range rangeReferece;
        private int paragraphEror = 0;
        private List<int> listParagraphEror;

        public bool showUC;
        private bool messageError = false;
        public ReferenceModel()
        {
            showUC = true;
            listParagraphEror = new List<int>();
        }

        private List<Word.Range> ListReferencesRange()
        {
            List<Word.Range> listReferences = new List<Word.Range>();
            for (int i = 1; i <= this.rangeReferece.Paragraphs.Count; i++)
            {
                if (this.rangeReferece.Paragraphs[i].Range.Text != "\r" ||
                    this.rangeReferece.Paragraphs[i].Range.Text != "")
                {
                    int start = this.rangeReferece.Paragraphs[i].Range.Start;
                    int end = this.rangeReferece.Paragraphs[i].Range.End - 1;
                    listReferences.Add(this.rangeReferece.Paragraphs[i].Range.Document.Range(start, end));
                }
            }

            return listReferences;
        }

        private Word.Range getRangeHaderReferece()
        {
            List<Word.Range> listReferences = new List<Word.Range>();
            int countPang = 0;
            Word.Range rangeForCheck = null;
            Word.Application wordApp = Globals.ThisAddIn.Application;
            //Word.Document d = wordApp.Documents[1];
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            object missing = System.Reflection.Missing.Value;
            Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int numberPang = document.ComputeStatistics(stat, missing);
            int start = 0;
            int end = 0;
            try
            {
                if (BackPangeStartAndEnd(ref start, ref end, ref countPang, ref rangeForCheck, numberPang, wordApp))
                {
                    NextPangeStartAndEnd(ref end, ref countPang, ref rangeForCheck, numberPang, wordApp);
                }
            }
            catch
            {
                return null;
            };
            if (start == 0 || end == 0)
            {
                return null;
            }
            Word.Range rangeReferences = Globals.ThisAddIn.Application.ActiveDocument.Range(start, end);
            return rangeReferences;

        }

        private bool BackPangeStartAndEnd(ref int strat, ref int end, ref int countPang, ref Word.Range rangeForCheck, int numberPang, Word.Application wordApp)
        {
            /* var wordApp = ThisAddIn.mainApplication.Application;
             //Word.Document d = wordApp.Documents[1];
             Word.Document document = ThisAddIn.mainApplication.Application.ActiveDocument;
             object missing = System.Reflection.Missing.Value;
             Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
             int numberPang = document.ComputeStatistics(stat, missing);*/
            // int checkStopPage = 0;
            for (int i = numberPang; i > 0; i--)
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";

                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;
                if (range.Paragraphs[1].Range.Text == "เอกสารอ้างอิง\r" && range.Paragraphs[1].Range.ParagraphFormat.Alignment == Word.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    bool checkBreak = false;
                    for (int ii = 3; ii <= range.Paragraphs.Count; ii++)
                    {
                        Word.Range r = range.Paragraphs[ii].Range;
                        if (r.Text == "\r" || r.Text[0] == ' ')
                        {
                            end = r.Start;
                            checkBreak = true;
                            break;
                        }
                        if (ii == 3)
                        {
                            strat = r.Start;
                        }
                        if (ii == range.Paragraphs.Count)
                        {
                            end = r.End;
                        }
                        rangeForCheck = r;
                        //listReferences.Add(r.Application.ActiveDocument.Range(r.Start, r.End - 1));
                        //string s = listReferences[0].Text;
                    }
                    if (checkBreak)
                    {
                        //countPang = i;
                        return false;
                    }
                    countPang = i;
                    return true;
                    //checkStopPage = 1;
                    //continue;
                }
            }
            return false;
        }

        private void NextPangeStartAndEnd(ref int end, ref int countPang, ref Word.Range rangeForCheck, int numberPang, Word.Application wordApp)
        {

            for (int i = countPang + 1; i <= numberPang; i++)
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";

                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;
                if (range.Paragraphs[1].Range.ParagraphFormat.Alignment != Word.WdParagraphAlignment.wdAlignParagraphCenter)
                {
                    bool checkBreak = false;
                    int ii = 1;
                    if (rangeForCheck != null)
                    {
                        if (rangeForCheck.Text == range.Paragraphs[1].Range.Text)
                        {
                            ii = 2;
                        }
                    }
                    for (; ii <= range.Paragraphs.Count; ii++)
                    {
                        Word.Range r = range.Paragraphs[ii].Range;
                        if (r.Text == "\r" || r.Text[0] == ' ')
                        {
                            checkBreak = true;
                            end = r.Start;
                            break;
                        }
                        if (ii == range.Paragraphs.Count)
                        {
                            end = r.End;
                        }

                        //listReferences.Add(r.Application.ActiveDocument.Range(r.Start, r.End - 1));
                    }
                    if (checkBreak)
                    {
                        return;
                    }
                }
                else
                {
                    break;

                }

            }
        }

        public void runCheckReferenceAll()
        {
            //System.Windows.Forms.MessageBox.Show("Paragraph ที่ผิด:");
            this.messageError = true;
            try
            {
                this.CheckReference();
            }
            catch
            {
                this.messageError = false;
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
            }
        }

        public void runEditReferenceAll()
        {
            if (!this.messageError)
            {
                return;
            }
            try
            {
                this.EditReference();
            }
            catch
            {
                
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
            }
        }

        private void CheckReference()
        {
            Ribbon1.referenceModelUC.Clear();
            Ribbon1.referenceModelUC.visibledefault();
            this.rangeReferece = null;
            this.rangeReferece = this.getRangeHaderReferece();
            if (rangeReferece != null)
            {
                if (this.faculty == "วิศวกรรมศาสตร์")
                {

                    //0.5 inches = 36 PostScript points
                    //this.rangeReferece.Paragraphs.LeftIndent = 36.0f;
                    //this.rangeReferece.Paragraphs.FirstLineIndent = -36.0f; // 0;
                    //System.Windows.Forms.MessageBox.Show(rangeReferece.Paragraphs.LeftIndent + "_" + rangeReferece.Paragraphs.FirstLineIndent);
                    //rangeReferece.Paragraphs.HangingPunctuation = -1;
                    //rangeReferece.Paragraphs[1].FirstLineIndent = 36.0f;
                    //rangeReferece.Paragraphs[1].Format.FirstLineIndent = -36.0f; // 0;
                    //rangeReferece.Paragraphs[1].Format.FirstLineIndent
                    //rangeReferece.ParagraphFormat.FirstLineIndent = -20.0f;

                    //System.Windows.Forms.MessageBox.Show(rangeReferece.Paragraphs[1].LeftIndent + "_" + rangeReferece.Paragraphs[1].FirstLineIndent);
                    this.FindReferencesEngineer();
                }
                else if (this.faculty == "บัณฑิตวิทยาลัย มน")
                {
                    //1.5 cm to point
                    //this.rangeReferece.Paragraphs.LeftIndent = 42.519685f;
                    //this.rangeReferece.Paragraphs.FirstLineIndent = -42.519685f;
                    this.FindReferencesGraduate();
                }
            }
            else
            {
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
            }
        }

        private List<string> SortWord()
        {
            NumberToThai numberToThai = new NumberToThai();
            object missing = System.Reflection.Missing.Value;
            List<string> listReferencesSort = new List<string>();
            //int start = 0;
            //int end = 0;
            //FindHaderRefereceStartAndEnd(ref start, ref end);
            //if (start==0)
            //{
            //    return null;
            //}
            //Word.Application aa = ThisAddIn.Application;
            //Word.Range rangeNew = //ThisAddIn.mainApplication.ActiveDocument.Range(start, end);
            //rangeNew.Select();
            
            this.rangeReferece.Copy();
            for (int i = 1; i <= this.rangeReferece.Paragraphs.Count; i++)
            {
                Word.Paragraph paragraph = this.rangeReferece.Paragraphs[i];
                string rangeNew = paragraph.Range.Text;
                string rangeOld = paragraph.Range.Text;
                int lengthCutNew   = 0;
                int lengthCutOld = numberToThai.NumberToName(ref rangeNew, ref lengthCutNew);
                if (lengthCutOld != 0)
                {
                    numberToThai.FindAndReplace(paragraph.Range, rangeOld.Substring(0, lengthCutOld), rangeNew.Substring(0, lengthCutNew));
                }
            }


            try
            {
                this.rangeReferece.Sort(false, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing,
                    Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, Word.WdLanguageID.wdThai);
            }
            catch
            {
                return null;
            }
            string s = this.rangeReferece.Text;
            listReferencesSort.AddRange(s.Split('\r'));
            //this.rangeReferece.Text = "";
            if (listReferencesSort[0] == "")
            {
                return null;
            }
            listReferencesSort.RemoveAt(listReferencesSort.Count - 1);
            while (true)
            {
                if (listReferencesSort[0] == "")
                {
                    listReferencesSort.RemoveAt(0);
                }
                else
                {
                    break;
                }
            }
            int countReferenceEng = 0;
            List<string> listReferencesSortNew = new List<string>();
            foreach (string listReference in listReferencesSort)
            {
                if (listReference == "")
                {
                    break;
                }
                if (!CheckTypeLanguage(listReference))
                {
                    break;
                }
                listReferencesSortNew.Add(listReference);
                countReferenceEng++;
            }
            bool checkNull = false;
            this.paragraphEror = 0;
            if (countReferenceEng > 0)
            {
                listReferencesSort.RemoveRange(0, countReferenceEng);
                if (!CheckNameYear(listReferencesSort, "TH"))
                {
                    checkNull = true;
                }
                numberToThai.NameToNumber(ref listReferencesSort);
                listReferencesSort.AddRange(listReferencesSortNew);
                if (!CheckNameYear(listReferencesSortNew, "EN"))
                {
                    checkNull = true;
                }
            }
            else
            {
                if (!CheckNameYear(listReferencesSort, "TH"))
                {
                    checkNull = true;
                }
                numberToThai.NameToNumber(ref listReferencesSort);
            }
            if (checkNull)
            {
                //System.Windows.Forms.MessageBox.Show("Paragraph ที่ผิด:" + this.paragraphEror);
                return null;
            }
            this.rangeReferece.Paste();
            return listReferencesSort;

        }

        private void FindReferencesGraduate()
        {
            this.FindReferencesEngineerNotForECPE();
        }

        private void FindReferencesEngineer()
        {
            if (this.department != "")
            {
                if (this.department == "วิศวกรรมไฟฟ้าและคอมพิวเตอร์")
                {
                    this.FindReferencesEngineerForECPE();
                }
                else if (this.department == "วิศวกรรมโยธา เครื่องกล และอุตสาหการ")
                {
                    this.FindReferencesEngineerNotForECPE();
                }
            }
        }

        private void FindReferencesEngineerNotForECPE()
        {
           
            ReferenceModelThai rmt = new ReferenceModelThai();
            ReferenceModelEN rme = new ReferenceModelEN();
            int cout = 0;
            
            List<Word.Range> listReferencesRange = this.ListReferencesRange();
            Ribbon1.referenceModelUC.setSortYear(listReferencesRange);
            if (listReferencesRange.Count == 0)
            {
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
                return;
            }
            Ribbon1.referenceModelUC.lbl_referenceAllCheck.Text = "" + listReferencesRange.Count;
            int countCheckSort = 0;
            int countRange = 0;
            this.listParagraphEror.Clear();
            foreach (Word.Range range in listReferencesRange)
            {
                int value = 0;
                if (range.Text == "" || range.Text == "\r" || range.Text[0] == '\t' || range.Text == null)
                {
                    break;
                }
                /* if (listReferencesSort[countCheckSort] != range.Text)
                 {
                     System.Windows.Forms.MessageBox.Show("Not Sort");
                     break;
                 }*/
                //countCheckSort++;
                string strCheck = range.Text;
                /*if (strCheck == "")
                {
                    break;
                }*/
                if (CheckTypeLanguage(strCheck))
                {
                    value = rme.ModelBookTypeBookEN(range, strCheck, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    //checkCharNumberTH(ref strOld, strCheck);
                    value = rmt.ModelBookTypeBookTH(range, strCheck, cout);
                }
                if (value == 0)
                {
                    Ribbon1.referenceModelUC.AddRangeErrorForReference(countRange + 1, range, 1);
                }
                if (range.Paragraphs.LeftIndent != Ribbon1.styles.Indent ||
                    range.Paragraphs.FirstLineIndent != -1 * Ribbon1.styles.Indent)
                {
                        Ribbon1.referenceModelUC.AddRangeErrorForReference(countRange + 1, range, 3);
                }
                countRange++;
            }
            List<string> listReferencesSort = this.SortWord();
            if (listReferencesSort == null)
            {
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
                return;
            }


            foreach (Word.Range range in listReferencesRange)
            {
                if (range.Text == "" || range.Text == "\r" || range.Text[0] == '\t' || range.Text == null)
                {
                    break;
                }

                if (listReferencesSort[countCheckSort] != range.Text)
                {
                    Ribbon1.referenceModelUC.AddRangeErrorForReference(countCheckSort + 1, range, 2);
                }
                countCheckSort++;
            }
            Ribbon1.referenceModelUC.ShowError();
            if (this.showUC)
            {
                this.show();
            }
        }

        private void FindReferencesEngineerForECPE()
        {
           
            
            ReferenceModelThai rmt = new ReferenceModelThai();
            ReferenceModelEN rme = new ReferenceModelEN();
            List<Word.Range> listReferencesRange = this.ListReferencesRange();
            if (listReferencesRange == null || listReferencesRange.Count == 0)
            {
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
                return;
            }
            Ribbon1.referenceModelUC.lbl_referenceAllCheck.Text = "" + listReferencesRange.Count;
            int cout = 0;
            LexerECPE lexerECPE = new LexerECPE();
            this.listParagraphEror.Clear();
            int countRangeError = 0;
            foreach (Word.Range range in listReferencesRange)
            {
                if (range.Text == "" || range.Text == "\r" || range.Text[0] == '\t' || range.Text == null)
                {
                    break;
                }
                string strCheck = range.Text;
                int value = 0;

                lexerECPE.sentence = strCheck;
                //lexerECPE.checkNumber();
                countRangeError++;
                if (lexerECPE.checkNumber() == 0)
                {
                    Ribbon1.referenceModelUC.AddRangeErrorForReference(countRangeError, range, 2);
                    //referenceModelUC.rangeReferenceErrorSortYear.Add(range);
                    //referenceModelUC.numberReferenceErrorSortYear.Add(countRangeError);
                    //System.Windows.Forms.MessageBox.Show(value + "");
                    //break;
                }

                if (CheckTypeLanguage(lexerECPE.sentence))
                {
                    value = rme.ModelBookTypeBookEN(range.Document.Range(range.Start + lexerECPE.numberMem.ToString().Length + 3, strCheck.Length + range.Start), lexerECPE.sentence, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    value = rmt.ModelBookTypeBookTH(range.Document.Range(range.Start + lexerECPE.numberMem.ToString().Length + 3, strCheck.Length + range.Start), lexerECPE.sentence, cout);
                }

                if (value == 0)
                {
                    Ribbon1.referenceModelUC.AddRangeErrorForReference(countRangeError, range, 1);
                }
                if (range.Paragraphs.LeftIndent != Ribbon1.styles.Indent ||
                    range.Paragraphs.FirstLineIndent != -1*Ribbon1.styles.Indent)
                {
                    Ribbon1.referenceModelUC.AddRangeErrorForReference(countRangeError, range, 3);
                }
            }

            Ribbon1.referenceModelUC.ShowErrorNumber();
            if (this.showUC)
            {
                this.show();
            }
        }

        private void EditReference()
        {
            this.rangeReferece = null;
            this.rangeReferece = this.getRangeHaderReferece();
            if (this.rangeReferece != null)
            {
                if (this.faculty == "วิศวกรรมศาสตร์")
                {
                    //0.5 inches = 36 PostScript points
                    //this.rangeReferece.Paragraphs.LeftIndent = 36.0f;
                    //this.rangeReferece.Paragraphs.FirstLineIndent = -36.0f;
                    this.FindEditReferencesEngineer();
                }
                else if (this.faculty == "บัณฑิตวิทยาลัย มน")
                {
                    //1.5 cm to point
                    // this.rangeReferece.Paragraphs.LeftIndent = 42.519685f;
                    // this.rangeReferece.Paragraphs.FirstLineIndent = -42.519685f;
                    this.FindEditReferencesEngineerNotForECPE();
                }
            }
            else
            {
                Ribbon1.referenceModelUC.setErrorNull();
                Ribbon1.referenceModelUC.btn_edit.Enabled = true;
                if (this.showUC)
                {
                    this.show();
                }
            }
        }

        private void FindEditReferencesEngineer()
        {
            if (this.department != "")
            {
                if (this.department == "วิศวกรรมไฟฟ้าและคอมพิวเตอร์")
                {
                    

                    this.FindEditReferencesEngineerForECPE();
                }
                else if (this.department == "วิศวกรรมโยธา เครื่องกล และอุตสาหการ")
                {
                    this.FindEditReferencesEngineerNotForECPE();
                }
            }
        }

        private void FindEditReferencesEngineerNotForECPE()
        {
            object missing = System.Reflection.Missing.Value;
            NumberToThai numberToThai = new NumberToThai();
            for (int i = 1; i <= this.rangeReferece.Paragraphs.Count; i++)
            {
                Word.Paragraph paragraph = this.rangeReferece.Paragraphs[i];
                string rangeNew = paragraph.Range.Text;
                string rangeOld = paragraph.Range.Text;
                int lengthCutNew = 0;
                int lengthCutOld = numberToThai.NumberToName(ref rangeNew, ref lengthCutNew);
                if (lengthCutOld != 0)
                {
                    numberToThai.FindAndReplace(paragraph.Range, rangeOld.Substring(0, lengthCutOld), rangeNew.Substring(0, lengthCutNew));
                }
            }
            this.rangeReferece.Sort(false, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing,
                        Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending,
                        ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, Word.WdLanguageID.wdThai);
            for (int i = 1; i <= this.rangeReferece.Paragraphs.Count; i++)
            {
                Word.Paragraph paragraph = this.rangeReferece.Paragraphs[i];
                /*string rangeNew = paragraph.Range.Text;
                string rangeOld = paragraph.Range.Text;
                int lengthCutNew = 0;
                int lengthCutOld = numberToThai.NumberToName(ref rangeNew, ref lengthCutNew);
                if (lengthCutOld != 0)
                {*/
                    numberToThai.NameToNumber(paragraph.Range);
               // }
            }
            int startCopy = 0;
            for (int i = 1; i <= this.rangeReferece.Paragraphs.Count; i++)
            {
                if (!CheckTypeLanguageFirst(this.rangeReferece.Paragraphs[i].Range.Text))
                {
                    startCopy = this.rangeReferece.Paragraphs[i].Range.Start;
                    Word.Range rangeCopy = this.rangeReferece.Paragraphs[i].Range.Document.Range(startCopy, this.rangeReferece.End);
                    rangeCopy.Copy();
                    rangeCopy.Text = "";
                    this.rangeReferece.Document.Range(this.rangeReferece.Start, this.rangeReferece.Start).Paste();

                    break;
                    //rangeCopy.Text = "";
                }
                /*else if(startCopy == 0)
                {
                    startCopy = -1;
                }*/
            }
            this.rangeReferece = this.getRangeHaderReferece();
            List<Word.Range> listReferencesRange = this.ListReferencesRange();
            if (listReferencesRange.Count == 0)
            {
                return;
            }
            ReferenceNameYear referenceNameYear = null;
            this.paragraphEror = 0;
            foreach (Word.Range r in listReferencesRange)
            {
                r.Paragraphs.LeftIndent = Ribbon1.styles.Indent;
                r.Paragraphs.FirstLineIndent = -Ribbon1.styles.Indent;
                string language = "";
                if (this.CheckTypeLanguage(r.Text))
                {
                    language = "EN";
                }
                else
                {
                    language = "TH";
                }

                if (language == "TH")
                {
                    if (referenceNameYear == null)
                    {
                        referenceNameYear = this.CheckNameYearTH(r.Text);
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                        referenceNameYear.range = this.rangeReferece;
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        ReferenceNameYear referenceNameYearNew = this.CheckNameYearTH(r.Text);
                        if (referenceNameYearNew == null)
                        {
                            continue;
                        }
                        referenceNameYearNew.range = this.rangeReferece;
                        referenceNameYear.Edit(referenceNameYearNew);
                        referenceNameYear = referenceNameYearNew;
                    }
                }
                else if (language == "EN")
                {
                    if (referenceNameYear == null)
                    {
                        referenceNameYear = this.CheckNameYearEN(r.Text);
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                        referenceNameYear.range = this.rangeReferece;
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        ReferenceNameYear referenceNameYearNew = this.CheckNameYearEN(r.Text);
                        if (referenceNameYearNew == null)
                        {
                            continue;
                        }
                        referenceNameYearNew.range = this.rangeReferece;
                        referenceNameYear.Edit(referenceNameYearNew);
                        referenceNameYear = referenceNameYearNew;
                    }
                }

            }
            //string rStr = range.Text;
        }

        private void FindEditReferencesEngineerForECPE()
        {
           
            ReferenceModelThai rmt = new ReferenceModelThai();
            ReferenceModelEN rme = new ReferenceModelEN();
            List<Word.Range> listReferencesRange = ListReferencesRange();
            if (listReferencesRange == null || listReferencesRange.Count == 0)
            {
                return;
            }
            //string[] listReferences = Regex.Split(r.Text, "\r");
            LexerECPE lexerECPE = new LexerECPE();
            foreach (Word.Range range in listReferencesRange)
            {
                if (range.Text == "" || range.Text == "\r" || range.Text[0] == '\t' || range.Text == null)
                {
                    break;
                }
                range.Paragraphs.LeftIndent = Ribbon1.styles.Indent;
                range.Paragraphs.FirstLineIndent = -Ribbon1.styles.Indent;
                string strCheck = range.Text;
                lexerECPE.sentence = strCheck;
                lexerECPE.range = range;
                if (!lexerECPE.editNumber())
                {
                    //System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }

                /*if (CheckTypeLanguage(lexerECPE.sentence))
                {
                    value = rme.ModelBookTypeBookEN(range.Document.Range(range.Start + lexerECPE.numberMem.ToString().Length + 3, strCheck.Length + range.Start), lexerECPE.sentence, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    value = rmt.ModelBookTypeBookTH(range.Document.Range(range.Start + lexerECPE.numberMem.ToString().Length + 3, strCheck.Length + range.Start), lexerECPE.sentence, cout);
                }
                value += lexerECPE.countLength;
                if (value == 0)
                {
                    System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }*/
                //cout += value;
                //r = r.Application.ActiveDocument.Range(cout);
            }
        }

        //ตรวจสอบภาษา
        private bool CheckTypeLanguage(string strCheck)
        {
            Match match = Regex.Match(strCheck, @"[ก-ฮะ-์]");
            if (match.Success)
            {
                return false;
            }
            return true;
        }

        //ตรวจสอบภาษา
        private bool CheckTypeLanguageFirst(string strCheck)
        {
            Match match = Regex.Match(strCheck, @"^[ก-ฮะ-์]");
            if (match.Success)
            {
                return false;
            }
            return true;
        }

        private bool CheckNameYear(List<string> listReferencesSort, string language)
        {
            ReferenceNameYear referenceNameYear = null;
            foreach (string listReference in listReferencesSort)
            {
                if (language == "TH")
                {
                    if (referenceNameYear == null)
                    {
                        referenceNameYear = this.CheckNameYearTH(listReference);
                        this.paragraphEror++;
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        this.paragraphEror++;
                        ReferenceNameYear referenceNameYearNew = this.CheckNameYearTH(listReference);
                        if (referenceNameYearNew == null)
                        {
                            continue;
                        }
                        if (!referenceNameYear.Check(referenceNameYearNew))
                        {
                            Ribbon1.referenceModelUC.numberReferenceErrorSortYearName.Add(this.paragraphEror);
                            referenceNameYear = null;
                            //return false;
                        }
                        referenceNameYear = referenceNameYearNew;
                    }
                }
                else if (language == "EN")
                {
                    if (referenceNameYear == null)
                    {
                        referenceNameYear = this.CheckNameYearEN(listReference);
                        this.paragraphEror++;
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        this.paragraphEror++;
                        ReferenceNameYear referenceNameYearNew = this.CheckNameYearEN(listReference);
                        if (referenceNameYearNew == null)
                        {
                            continue;
                        }
                        if (!referenceNameYear.Check(referenceNameYearNew))
                        {
                            Ribbon1.referenceModelUC.numberReferenceErrorSortYearName.Add(this.paragraphEror);
                            referenceNameYear = null;
                            //return false;
                        }
                        referenceNameYear = referenceNameYearNew;
                    }
                }
            }
            return true;
        }

        private bool CutNameYearTH(ref string year, ref char character)
        {

            Match match = Regex.Match(year[year.Length - 1] + "", "[0-9]");
            if (match.Success)
            {
                return true;
            }
            else
            {
                match = Regex.Match(year[year.Length - 1] + "", "[ก-ฮ]");
                if (match.Success)
                {
                    character = year.Substring(year.Length - 1)[0];
                    year = year.Substring(0, year.Length - 1);
                    return true;
                }
            }
            return false;
        }

        private bool CutNameYearEN(ref string year, ref char character)
        {

            Match match = Regex.Match(year[year.Length - 1] + "", "[0-9]");
            if (match.Success)
            {
                return true;
            }
            else
            {
                match = Regex.Match(year[year.Length - 1] + "", "[a-z]");
                if (match.Success)
                {
                    character = year.Substring(year.Length - 1)[0];
                    year = year.Substring(0, year.Length - 1);
                    return true;
                }
            }
            return false;
        }

        private ReferenceNameYear CheckNameYearTH(string strCheck)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.countLength = 0;
            int memNum = 0;
            if (l.ForNamesForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearTH(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }

            }
            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForBookNameToBracketForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearTH(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }
            }

            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForNameOnePreviousForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForNameYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearTH(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }

            }

            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForNamesForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForNameYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearTH(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }
            }

            return null;
        }

        private ReferenceNameYear CheckNameYearEN(string strCheck)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.countLength = 0;
            int memNum = 0;
            if (l.ForNamesForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearEN(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }

            }
            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForBookNameToBracketForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearEN(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }
            }

            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForNameOnePreviousForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForNameYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearEN(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }

            }

            l.sentence = strCheck;
            l.countLength = 0;
            if (l.ForNamesForCheck())
            {
                memNum = l.countLength;
                string name = strCheck.Substring(0, memNum);
                if (l.ForNameYearForCheck())
                {
                    string year = strCheck.Substring(memNum, l.countLength - memNum - 3).Substring(1);
                    char character = ' ';
                    bool forCheck = CutNameYearEN(ref year, ref character);
                    return new ReferenceNameYear(name, year, character, forCheck);
                }
            }
            return null;
        }

        public void show()
        {
            Ribbon1.showCustomTaskPane(7);
            if (!Ribbon1.referenceModelUC.checkSetClick)
            {
                Ribbon1.referenceModelUC.btn_edit.Click += new System.EventHandler(this.btn_edit_Click);
                Ribbon1.referenceModelUC.checkSetClick = true;
            }
            /*
            if (ThisAddIn.mainCustomTaskPane.Count > 0)
            {
                for (int i = 0; i < ThisAddIn.mainCustomTaskPane.Count; ++i)
                {
                    ThisAddIn.mainCustomTaskPane.RemoveAt(i);
                }
            }
            
            Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.referenceModelUC, "Reference");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/
            
        }

        public void showForAll()
        {
            Ribbon1.showCustomTaskPane(7, true);
            if (!Ribbon1.referenceModelUC.checkSetClick)
            {
                Ribbon1.referenceModelUC.btn_edit.Click += new System.EventHandler(this.btn_edit_Click);
                Ribbon1.referenceModelUC.checkSetClick = true;
            }
            /*if (ThisAddIn.mainCustomTaskPane.Count > 1)
            {
                for (int i = 1; i < ThisAddIn.mainCustomTaskPane.Count; ++i)
                {
                    //ThisAddIn.mainCustomTaskPane.RemoveAt(i);
                    ThisAddIn.mainCustomTaskPane[i].Visible = false;
                }
            }
            this.runCheckReferenceAll();
            Ribbon1.referenceModelUC.btn_edit.Click += new System.EventHandler(this.btn_edit_Click);
            Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.referenceModelUC, "Reference");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/

        }

        private void btn_edit_Click(object sender, EventArgs e)
        {
            this.EditReference();
            this.CheckReference();
            if (!this.showUC)
            {
                Ribbon1.showCheckAllUC.setButtonClickALL();
            }
            Ribbon1.saveFileAuto();
        }
    }
}
