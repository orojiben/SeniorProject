using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeniorProject
{
     class ReferenceModel
    {
        public string faculty;
        public string department;
        public void runCheckReferenceAll()
        {
            CheckReference();
          /*  try
            {
                System.IO.File.Delete(@"C:\Users\Nisit\Documents\bin.docx");
                
            }
            catch 
            {
                try
                {
                    System.IO.File.Delete(@"C:\Users\Nisit\Documents\bin.docx");
                }
                catch
                {
                }
            }*/
        }

        public void DeleteBin()
        {
            try
            {
                System.IO.File.Delete(@"C:\Users\Nisit\Documents\bin.docx");
                return;
            }
            catch
            {
                DeleteBin();
            }
        }

        private void CheckReference()
        {
            var wordApp = Globals.ThisAddIn.Application;

            //List<Word.Range> lsR = new List<Word.Range>();
           // List<string> lsS = new List<string>();
            // wordApp.ActiveDocument.Paragraphs.IndentCharWidth(7);
            //wordApp.ActiveDocument.Paragraphs.CharacterUnitFirstLineIndent= 0.5f;

            /*1.5 cm to point
            wordApp.ActiveDocument.Paragraphs.LeftIndent = 42.519685f;
            wordApp.ActiveDocument.Paragraphs.FirstLineIndent = -42.519685f;
             */
            //0.5 inches = 36 PostScript points
            wordApp.ActiveDocument.Paragraphs.LeftIndent = 36.0f;
            wordApp.ActiveDocument.Paragraphs.FirstLineIndent = -36.0f;
            //wordApp.ActiveDocument.Paragraphs.TabHangingIndent(0);
            //wordApp.ActiveDocument.Paragraphs.ca
            foreach (Word.Range range in wordApp.ActiveDocument.StoryRanges)
            {
                string[] newRanges = range.Text.Split('\r');

                List<string> listReferences = new List<string>(newRanges);
                
                while (listReferences[listReferences.Count - 1] == "")
                {
                    listReferences.RemoveAt(listReferences.Count - 1);
                }

                List<Word.Range> listReferencesRange = SubRange(range,listReferences);

                if (this.faculty == "Engineering")
                {
                    FindReferencesEngineer(listReferencesRange);
                }
                else if (this.faculty == "Graduate")
                {
                    FindReferencesGraduate(listReferencesRange);
                }
                //Word.Range rangeNew = range.InlineShapes.Application.ActiveDocument.StoryRanges[0];
            }
        }

        private List<string> SortWord(List<Word.Range> litsRange)
        {
            //ช้าเพราะเปิด App
            var winword = new Microsoft.Office.Interop.Word.Application();
            object missing = System.Reflection.Missing.Value;
            Microsoft.Office.Interop.Word.Document document = winword.Documents.Add(ref missing, ref missing, ref missing, ref missing);
            document.Content.SetRange(0, 0);
            document.Content.Text = "";
            foreach (Word.Range Range in litsRange)
            {
                document.Content.Text += Range.Text;

            }
            //document.Content.Text = "เอโมโตะ สิงห์น้อย. (2549). คำนามประสม: ศาสตร์และศิลป์ในการสร้างคำไทย. กรุงเทพฯ: สำนักพิมพ์แห่งจุฬาลงกรณ์มหาวิทยาลัย.";
            //document.Content.Text += "อัญชลี สิงห์น้อย. (2549). คำนามประสม: ศาสตร์และศิลป์ในการสร้างคำไทย. กรุงเทพฯ: สำนักพิมพ์แห่งจุฬาลงกรณ์มหาวิทยาลัย.";
            List<string> listReferencesSort = new List<string>();
           
            foreach (Word.Range range in document.StoryRanges)
            {

                Word.Range rangeNew = range;
                rangeNew.Sort(false, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing,
                    Word.WdSortOrder.wdSortOrderAscending, ref missing, ref missing, Word.WdSortOrder.wdSortOrderAscending,
                    ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, ref missing, Word.WdLanguageID.wdThai);
                string s = rangeNew.Text;
                listReferencesSort.AddRange(s.Split('\r'));
            }
            listReferencesSort.RemoveAt(listReferencesSort.Count-1);
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
            if (countReferenceEng > 0)
            {
                listReferencesSort.RemoveRange(0, countReferenceEng);
                if (!CheckNameYear(listReferencesSort, "TH"))
                {
                    checkNull = true;
                }
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
            }
            object fileName = "bin.docx";
            winword.ActiveDocument.SaveAs(ref fileName,
    ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing,
    ref missing, ref missing, ref missing, ref missing, ref missing);
            //winword.Selection.Delete();
            winword.Quit();
            
            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(winword);
            DeleteBin();
            if (checkNull)
            {
                return null;
            }
            return listReferencesSort;
        }

        public void SampleRegexUsage(Word.Range r,ref bool check)
        {
            string regex = @"(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\s\([บ][ร][ร][ณ][า][ธ][ิ][ก][า][ร]\)\,\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\([ห][น][้][า]\s[1-9]([0-9])*(\s)?\-(\s)?[1-9]([0-9])*\)\.\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\:\s(([ก-ฮะ-์])+(ฯ)?(((\.)?\,\s)|((\.)(\s)?)|(\s)|((\:)(\s)))?)+\r";
            //string regex = @"(((([ก-ฮะ-์])*(\s)?)(\,\s(([ก-ฮะ-์])*(\s)?)+)?)+\.\s)";
            //RegexOptions options = RegexOptions.RightToLeft | RegexOptions.None;
            string input = r.Text;
            //int a = 9;
            try {

                Match matche = Regex.Match(input, regex);
               
           if (matche.Success)
            {
                check = true;
              //  a=1;
            }
            //int b = 9;
            }
            catch (ArgumentException)
            {
                System.Windows.Forms.MessageBox.Show("");
                // Do nothing: assume that exception represents no match.
            }
        }

        private void FindReferencesGraduate(List<Word.Range> litsRange)
        {
            if (this.department == "")
            {
            }
        }

        private void FindReferencesEngineer(List<Word.Range> litsRange)
        {
            if (this.department != "")
            {
                if (this.department == "วิศวกรรมไฟฟ้าและคอมพิวเตอร์")
                {
                    FindReferencesEngineerForECPE(litsRange);
                }
                else if (this.department == "วิศวกรรมโยธา เครื่องกล และอุตสาหการ")
                {
                    FindReferencesEngineerNotForECPE(litsRange);
                }
            }
        }

        private void FindReferencesEngineerNotForECPE(List<Word.Range> litsRange)
        {
            //bool check = false;
           // SampleRegexUsage(r,ref check);
           // string[] listReferences = Regex.Split(r.Text, "\r");
            int cout = 0;
            List<string> listReferencesSort = SortWord(litsRange);
            if (listReferencesSort == null)
            {
                return;
            }
            //System.Windows.Forms.MessageBox.Show(listReferences.Length + " ^^");
            //string strOld = "";
            //checkCharNumberTH
            int countCheckSort = 0;
            int countRange = 0;
            foreach (Word.Range range in litsRange)
            {
                int value = 0;
                if (range.Text == "" || range.Text == "\r" || range.Text == null)
                {
                    break;
                }
                if (listReferencesSort[countCheckSort] != range.Text)
                {
                    System.Windows.Forms.MessageBox.Show("Not Sort");
                    break;
                }
                countCheckSort++;
                string strCheck = range.Text;
                if (strCheck == "")
                {
                    break;
                }
                if (CheckTypeLanguage(strCheck))
                {
                    value = ModelBookTypeBookEN(litsRange[countRange], strCheck, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    //checkCharNumberTH(ref strOld, strCheck);
                    value = ModelBookTypeBookTH(litsRange[countRange], strCheck, cout);
                }
                if (value == 0)
                {
                    System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }
                cout += value;
                //r = r.Application.ActiveDocument.Range(cout);
                countRange++;
            }
        }

        private void FindReferencesEngineerForECPE(List<Word.Range> litsRange)
        {
            //string[] listReferences = Regex.Split(r.Text, "\r");
            int cout = 0;
            LexerECPE lexerECPE = new LexerECPE();
            foreach (Word.Range range in litsRange)
            {
                if (range.Text == "" || range.Text == "\r" || range.Text == null)
                {
                    break;
                }
                string strCheck = range.Text;
                int value = 0;

                lexerECPE.sentence = strCheck;
                if (lexerECPE.checkNumber() == 0)
                {
                    System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }

                if (CheckTypeLanguage(lexerECPE.sentence))
                {
                    value = ModelBookTypeBookEN(range.Application.ActiveDocument.Range(lexerECPE.countLength, strCheck.Length), lexerECPE.sentence, cout);// ModelBookTypeBookTH(r, cout);
                }
                else
                {
                    value = ModelBookTypeBookTH(range.Application.ActiveDocument.Range(lexerECPE.countLength, strCheck.Length), lexerECPE.sentence, cout);
                }
                value += lexerECPE.countLength;
                if (value == 0)
                {
                    System.Windows.Forms.MessageBox.Show(value + "");
                    break;
                }
                //cout += value;
                //r = r.Application.ActiveDocument.Range(cout);
            }
        }
        //ตรวจสอบภาษา
        private bool CheckTypeLanguage(string strCheck)
        {
            Match match = Regex.Match(strCheck, @"^[0-9]*(\s)?[A-Za-z]");
            if (match.Success)
            {
                return true;
            }
            return false;
        }

        private bool CheckNameYear(List<string> listReferencesSort,string language)
        {
            ReferenceNameYear referenceNameYear = null;
            foreach (string listReference in listReferencesSort)
            {
                if (language == "TH")
                {
                    if (referenceNameYear == null)
                    {
                        referenceNameYear = CheckNameYearTH(listReference);
                        if (referenceNameYear == null)
                        {
                            continue;
                        }
                    }
                    else
                    {
                        ReferenceNameYear referenceNameYearNew = CheckNameYearTH(listReference);
                        if (referenceNameYearNew == null)
                        {
                            continue;
                        }
                        if (!referenceNameYear.Check(referenceNameYearNew))
                        {
                            return false;
                        }
                        referenceNameYear = referenceNameYearNew;
                    }
                }
                else if (language == "EN")
                {
                    referenceNameYear = CheckNameYearEN(listReference);
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

        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForYearCreate())
                        {

                        }
                        l.countCutNotBold = l.countLength;
                        if (!l.ForBookTranslator())
                        {
                            return ModelBookTypeArticleTH(r, strCheck, cout);
                        }

                        if (l.ForPlaceEnd())
                        {
                            l.countCutNotBold = l.countLength;
                            if (!l.ForBookAddEnd())
                            {
                                return ModelBookTypeArticleTH(r, strCheck, cout);
                            }
                            System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                            return l.countLength;
                        }
                        
                    }
                }
            }
            return ModelBookTypeArticleTH(r, strCheck, cout);
        }
        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int ModelBookTypeArticleTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameToIn())
                    {
                        
                        if (l.ForBookNameIn())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeEncyclopediaTH(r, strCheck, cout);
        }

         //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int ModelBookTypeEncyclopediaTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInDotEditor())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageAndBook())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {

                                    System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                    
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH( r,  strCheck,  cout);
        }

         //เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNamesNF())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNarrator())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookName())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameInDotEditor())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageAndBook())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPlaceEnd())
                                    {
                                        System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH2(r, strCheck, cout);
        }

         //เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryTH2(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNamesNF())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNarrator())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookName())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPlaceEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
                        }   
                    }
                }
            }
            return ModelJournalTypeArticlesTH(r,strCheck,cout);
        }

        //บทความทั่วไป เอกสารประเภทวารสาร
        private int ModelJournalTypeArticlesTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeReviewTH(r, strCheck, cout);
        }

         //บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร
        private int ModelJournalTypeReviewTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameES())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameReview(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeInterviewTH(r, strCheck, cout);
        }

        //บทสัมภาษณ์ เอกสารประเภทวารสาร
        private int ModelJournalTypeInterviewTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameES())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForNamesInterviewer(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทสัมภาษณ์ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeBookTH(r, strCheck, cout);
        }

        //หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeBookTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeColumnTH(r, strCheck, cout);
        }

        //กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeColumnTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForColumnEnd())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }

                }
            }
            l.sentence = strCheck;
            l.countLength = 0;
            l.countCutNotBold = 0;
            if (l.ForBookName())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForColumnEnd())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }

                }
            }
            return ModelNewspaperTypeInterviewTH(r, strCheck, cout);
        }

        //กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeInterviewTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForNamesInterviewer(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameEC())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEC())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForBookNameEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisAbstractBookTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ
        private int ModelThesisTypeThesisAbstractBookTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameToIn())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInDotEditor())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ");
                                    return l.countLength;
                                }
                            }
                        }
                    }

                }
            }
            return ModelThesisTypeThesisAbstractJournalTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร
        private int ModelThesisTypeThesisAbstractJournalTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelThesisTypeThesisOnlineTH(r, strCheck, cout);
        }

        //ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
         // Enter URL มันจะตัดทิ้ง
        private int ModelThesisTypeThesisOnlineTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameEC())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEC())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForBookName())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForSearch())
                                    {
                                        l.countCutNotBold = l.countLength;
                                        if (l.ForURL())
                                        {
                                            System.Windows.Forms.MessageBox.Show("ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                            return l.countLength;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeLetterTH(r, strCheck, cout);
        }

        //จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeLetterTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForMonthYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if(l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeBrochuresAndLeafletsTH(r, strCheck, cout);
        }

        //จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeBrochuresAndLeafletsTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracketBold())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBrochuresAndLeaflets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
            }
            return ModelOtherTypeArchivesTH(r, strCheck, cout);
        }

        //จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeArchivesTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracket())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNumber())
                    {
                    }
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameEndBold())
                    {

                        System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                        return l.countLength;

                    }
                }
            }
            return ModelOtherTypeGovernmentGazetteTH(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracket())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForAt())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd2())
                            {
                                System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeAudioTH(r, strCheck, cout);
        }

        //สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeAudioTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNameOnePrevious())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNameYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameESBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameNotPublished(0))
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeImageTH(r, strCheck, cout);
        }

        //ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeImageTH(Word.Range r, string strCheck, int cout)
        {
             LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameESBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameNotPublished(0))
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameDB(0))
                            {
                                System.Windows.Forms.MessageBox.Show("ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelOnlineTypeOnlineTH(r, strCheck, cout);
        }

        //บทความออนไลน์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOnlineTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            bool check = false;
            if (l.ForNames())
            {
                check = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    check = true;
                }
            }
            if (check)
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForSearch())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            return ModelOnlineTypeOtherTH(r, strCheck, cout);
        }

        //บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOtherTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            bool checkName = false;
            if (l.ForNames())
            {
                checkName = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    checkName = true;
                }
            }
            if (checkName)
            {
                bool check = false;
                string sentenceCopy = l.sentence;
                int countLengthCopy = l.countLength;
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    check = true;
                }
                else
                {
                    l.sentence = sentenceCopy;
                    l.countLength = countLengthCopy;
                    l.countCutNotBold = l.countLength;
                    if (l.ForMonthYear())
                    {
                        check = true;
                    }
                }
                if (check)
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBold())
                        {

                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
           
            return ModelOnlineTypeElectronicTH(r, strCheck, cout);
        }

        //บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeElectronicTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            bool check = false;
            if (l.ForNames())
            {
                check = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    check = true;
                }
            }
            if (check)
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNameYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            return 0;
        }

//===========================================================================================================================//
//===========================================================================================================================//
        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        private int ModelBookTypeBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    //l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForYearCreate())
                        {

                        }
                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                            return l.countLength;
                        }

                    }
                }
            }
            return ModelBookTypeArticleEN(r, strCheck, cout);
        }

        //บทความในหนังสือ เอกสารประเภทหนังสือ
        private int ModelBookTypeArticleEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameToIn())
                    {
                        if (l.ForBookNameIn())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeEncyclopediaEN(r, strCheck, cout);
        }

        //หนังสือสารานุกรม เอกสารประเภทหนังสือ
        private int ModelBookTypeEncyclopediaEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInDotEditor())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {

                                    System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                    return l.countLength;

                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryEN(r, strCheck, cout);
        }

        //เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNamesNF())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNarrator())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookName())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameInDotEditor())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPage())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPlaceEnd())
                                    {
                                        System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryEN2(r, strCheck, cout);
        }

        //เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ
        private int ModelBookTypeHandoutLibraryEN2(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNamesNF())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNarrator())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookName())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPlaceEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelJournalTypeArticlesEN(r, strCheck, cout);
        }

        //บทความทั่วไป เอกสารประเภทวารสาร
        private int ModelJournalTypeArticlesEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeReviewEN(r, strCheck, cout);
        }


        //บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร
        private int ModelJournalTypeReviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameES())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameReview(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelJournalTypeInterviewEN(r, strCheck, cout);
        }

        //บทสัมภาษณ์ เอกสารประเภทวารสาร
        private int ModelJournalTypeInterviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameES())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForNamesInterviewer(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทสัมภาษณ์ เอกสารประเภทวารสาร");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeBookEN(r, strCheck, cout);
        }

        //หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelNewspaperTypeColumnEN(r, strCheck, cout);
        }

        //กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeColumnEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForColumnEnd())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                                return l.countLength;
                            }
                        }
                    }

                }
            }
            else
            {
                l.sentence = strCheck;
                l.countLength = 0;
                l.countCutNotBold = 0;
                if (l.ForBookName())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForColumnEnd())
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelNewspaperTypeInterviewEN(r, strCheck, cout);
        }

        //กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์
        private int ModelNewspaperTypeInterviewEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForNamesInterviewer(0))
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameECBold())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInitials())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEC())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForBookNameEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }

            }
            return ModelThesisTypeThesisAbstractBookEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ
        private int ModelThesisTypeThesisAbstractBookEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameToIn())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInDotEditor())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEnd())
                                {
                                    System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ");
                                    return l.countLength;
                                }
                            }
                        }
                    }

                }
            }
            return ModelThesisTypeThesisAbstractJournalEN(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร
        private int ModelThesisTypeThesisAbstractJournalEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelThesisTypeThesisOnlineEN(r, strCheck, cout);
        }

        //ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง
        private int ModelThesisTypeThesisOnlineEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameEC())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEC())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForBookName())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForSearch())
                                    {
                                        l.countCutNotBold = l.countLength;
                                        if (l.ForURL())
                                        {
                                            System.Windows.Forms.MessageBox.Show("ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                            return l.countLength;
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeLetterEN(r, strCheck, cout);
        }


        //จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeLetterEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForMonthYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameECBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                System.Windows.Forms.MessageBox.Show("จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelOtherTypeBrochuresAndLeafletsEN(r, strCheck, cout);
        }

        //จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeBrochuresAndLeafletsEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracketBold())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBrochuresAndLeaflets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEnd())
                        {
                            System.Windows.Forms.MessageBox.Show("จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
            }
            return ModelOtherTypeArchivesEN(r, strCheck, cout);
        }

        //จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeArchivesEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracket())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNumber())
                    {
                    }
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameEndBold())
                        {

                            System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;

                        }
                }
            }
            return ModelOtherTypeGovernmentGazetteEN(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracket())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForAt())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
            return ModelMaterialNotPublishedTypeAudioEN(r, strCheck, cout);
        }

        //สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeAudioEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNameOnePrevious())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNameYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameESBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameNotPublished(0))
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }
           /* l.sentence = strCheck;
            l.countLength = 0;
            check = true;
            if (check)
            {
                if (l.ForBookNameES())
                {
                    if (l.ForBookNameReview(0))
                    {
                        if (l.ForNameYear())
                        {
                            if (l.ForBookNameEnd())
                            {
                                System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }
            }*/
            return ModelMaterialNotPublishedTypeImageEN(r, strCheck, cout);
        }

        //ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์
        private int ModelMaterialNotPublishedTypeImageEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameESBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameNotPublished(0))
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameDB(0))
                            {
                                System.Windows.Forms.MessageBox.Show("ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                return l.countLength;
                            }
                        }
                    }
                }

            }
            return ModelOnlineTypeOnlineEN(r, strCheck, cout);
        }

        //บทความออนไลน์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOnlineEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            bool check = false;
            if (l.ForNames())
            {
                check = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    check = true;
                }
            }
            if (check)
            {
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForSearch())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForURL())
                            {
                                System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                return l.countLength;
                            }
                        }

                    }
                }
            }
            else
            {
                l.countCutNotBold = 0;
                if (l.ForBookName())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForDate())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }

                        }
                    }
                }
            }
            return ModelOnlineTypeOtherEN(r, strCheck, cout);
        }

        //บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOtherEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            bool checkName = false;
            if (l.ForNames())
            {
                checkName = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    checkName = true;
                }
            }
            if (checkName)
            {
                bool check = false;
                string sentenceCopy = l.sentence;
                int countLengthCopy = l.countLength;
                l.countCutNotBold = l.countLength;
                if (l.ForDate())
                {
                    check = true;
                }
                else
                {
                    l.sentence = sentenceCopy;
                    l.countLength = countLengthCopy;
                    l.countCutNotBold = l.countLength;
                    if (l.ForMonthYear())
                    {
                        check = true;
                    }
                }
                if (check)
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            return ModelOnlineTypeElectronicEN(r, strCheck, cout);
        }

        //บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeElectronicEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            bool check = false;
            if (l.ForNames())
            {
                check = true;
            }
            else
            {
                l.sentence = strCheck;
                l.range = r;
                l.countCutNotBold = 0;
                l.countLength = 0;
                if (l.ForBookName())
                {
                    check = true;
                }
            }
            if (check)
            {
                l.countCutNotBold = l.countLength;
                if (l.ForNameYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookName())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBold())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            return 0;
        }


        private void CheckStringMatch(string strFromRange, string regex, ref int checkValue)
        {
            Match match = Regex.Match(strFromRange, regex);
            if (match.Success)
            {
                checkValue = match.Value.Length;
                return;
            }
                checkValue= -1;
        }


//===========================================================================================================================//
//===========================================================================================================================//
        private List<Word.Range> SubRange(Word.Range range, List<string> listReferences)
        {
            //foreach (Word.Range range 
            List<Word.Range> listReferencesRange = new List<Word.Range>();
            int countForCut = 0;
            foreach (string listReference in listReferences)
            {
                Word.Range rangeNew = range.Application.ActiveDocument.Range(countForCut, countForCut + listReference.Length);
               // range.
                listReferencesRange.Add(rangeNew);
               // int a = rangeNew.Text.Length;
                countForCut += listReference.Length+1;
            }
            return listReferencesRange;
        }

    }
}
