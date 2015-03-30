using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class FontManager
    {

        public struct Problem
        {
            public Word.Range range;
            public string comment;
            public int trueSize;
        }

        static List<Problem> listPb;
        static int index = 0;

        private void hightLightChar(Range rng, int size)
        {

          //  for (int x = rng.Start; x < rng.End; x++)
          //  {
                //if (Globals.ThisAddIn.Application.ActiveDocument.Range(x, x + 1).Font.Size != size)
                //{
                    rng.HighlightColorIndex = WdColorIndex.wdYellow;
                //}

           // }
        }

        public void CheckFontSize(int subStance, int subHeading, int topic, int nameChapter, int chapter, FontUC fc)
        {
            index = 0;
            Word.Application wordApp = Globals.ThisAddIn.Application;
            object missing = System.Reflection.Missing.Value;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int numberPang = document.ComputeStatistics(stat, missing);
            bool isFalse = false;
            listPb = new List<Problem>();
            fc.pgbFontSize.Minimum = 0;
            fc.pgbFontSize.Maximum = numberPang;
            fc.pgbFontSize.Value = 0;
            for (int i = 1; i <= numberPang; i++)
            {
                Problem pb = new Problem();
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";
                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;

                for (int ii = 1; ii <= range.Paragraphs.Count; ii++)
                {
                    if ((ii == 1 || ii == 2) && range.Paragraphs[ii].Alignment == WdParagraphAlignment.wdAlignParagraphCenter)
                    {
                        string buffRang = "";
                        if (range.Paragraphs[ii].Range.Text.Length > 4)
                        {
                            buffRang = range.Paragraphs[ii].Range.Text.Substring(0, 5);
                        }
                        else
                        {
                            buffRang = range.Paragraphs[ii].Range.Text;
                        }
                        if (buffRang == "บทที่")
                        {
                            if (range.Paragraphs[ii].Range.Font.Size != chapter)
                            {
                                pb.trueSize = chapter;
                                isFalse = true;
                                hightLightChar(range.Paragraphs[ii].Range, chapter);
                            }

                        }
                        else if (range.Paragraphs[ii].Range.Font.Size != nameChapter)
                        {
                            pb.trueSize = nameChapter;
                            isFalse = true;
                            hightLightChar(range.Paragraphs[ii].Range, nameChapter);
                        }
                    }
                    else if ((range.Paragraphs[ii].Alignment == WdParagraphAlignment.wdAlignParagraphLeft
                || range.Paragraphs[ii].Alignment == WdParagraphAlignment.wdAlignParagraphThaiJustify
                || range.Paragraphs[ii].Alignment == WdParagraphAlignment.wdAlignParagraphJustify)
                && (range.Paragraphs[ii].Range.Font.Bold == -1))
                    {
                        if (range.Paragraphs[ii].Range.Text[0] != '\b'
                            && range.Paragraphs[ii].Range.Text[0] != '\t'
                            && range.Paragraphs[ii].LeftIndent == 0
                            && range.Paragraphs[ii].FirstLineIndent == 0)
                        {
                            if (range.Paragraphs[ii].Range.Font.Size != topic)
                            {
                                isFalse = true;
                                pb.trueSize = topic;
                                hightLightChar(range.Paragraphs[ii].Range, topic);
                            }
                        }
                        else if (range.Paragraphs[ii].Range.Font.Size != subHeading)
                        {
                            isFalse = true;
                            pb.trueSize = subHeading;
                            hightLightChar(range.Paragraphs[ii].Range, subHeading);
                        }
                    }
                    else
                    {
                        if (range.Paragraphs[ii].Range.Font.Size != subStance)
                        {
                            isFalse = true;
                            pb.trueSize = subStance;
                            hightLightChar(range.Paragraphs[ii].Range, subStance);
                        }
                    }

                    if (isFalse)
                    {
                        //range.Paragraphs[ii].Range.HighlightColorIndex = WdColorIndex.wdYellow;
                        pb.range = range.Paragraphs[ii].Range;
                        pb.comment = "ขนาดอักษรควรจะเป็น " + pb.trueSize;
                        listPb.Add(pb);
                        isFalse = false;
                    }

                }

                fc.pgbFontSize.Increment(1);
                //Thread.Sleep(500);
            }
            //fc.pgbFontSize.Value = fc.pgbFontSize.Minimum;
            if (listPb.Count > 0)
            {
                fc.btn_lookError.Visible = true;
                fc.lblSizeFault.Text = listPb.Count + " ข้อผิดพลาด";
            }
            else
            {
                fc.btn_lookError.Visible = false;
                fc.lblSizeFault.Text = listPb.Count + " ข้อผิดพลาด";
                fc.pnlSizeCheck.Visible = false;
                fc.lblMainText.Text = "ข้อผิดพลาด";
            }
        }

        public void checkFontName(string font, FontUC fc)
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;
            object missing = System.Reflection.Missing.Value;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int numberPang = document.ComputeStatistics(stat, missing);
            int countError = 0;
            Word.Range rng;
            rng = Globals.ThisAddIn.Application.ActiveDocument.Content;
            fc.pgbFontName.Minimum = 0;
            fc.pgbFontName.Maximum = numberPang;
            fc.pgbFontName.Value = 0;
            for (int i = 1; i <= numberPang; i++)
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";
                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;

                if (range.Font.Name == "")
                {
                    for (int ii = 1; ii <= range.Paragraphs.Count; ii++)
                    {
                        if (range.Paragraphs[ii].Range.Font.Name == "")
                        {
                            //rng.Paragraphs[i].Range.HighlightColorIndex = WdColorIndex.wdRed;
                            //rng.Paragraphs[i].Application
                            foreach (Microsoft.Office.Interop.Word.Range rngWord in range.Paragraphs[ii].Range.Words)
                            {
                                if (rngWord.Font.Name == "")
                                {
                                    foreach (Microsoft.Office.Interop.Word.Range rngCharacter in rngWord.Characters)
                                    {
                                        if (rngCharacter.Font.Name != font)
                                        {
                                            rngCharacter.HighlightColorIndex = WdColorIndex.wdYellow;
                                            countError++;
                                        }
                                    }
                                }
                                else if (rngWord.Font.Name != font)
                                {
                                    rngWord.HighlightColorIndex = WdColorIndex.wdYellow;
                                    countError++;
                                }
                                
                            }
                            /*for (int j = rng.Paragraphs[i].Range.Start; j < rng.Paragraphs[i].Range.End; j++)
                            {
                                if (Globals.ThisAddIn.Application.ActiveDocument.Range(j, j + 1).Font.Name != font)
                                {
                                    Globals.ThisAddIn.Application.ActiveDocument.Range(j, j + 1).HighlightColorIndex = WdColorIndex.wdYellow;
                                }
                            }*/
                        }
                        else if (range.Paragraphs[ii].Range.Font.Name != font)
                        {
                            range.Paragraphs[ii].Range.HighlightColorIndex = WdColorIndex.wdYellow;
                            countError++;
                        }
                    }
                }
                else if (range.Font.Name != font)
                {
                    range.HighlightColorIndex = WdColorIndex.wdYellow;
                    countError++;
                }
                fc.pgbFontName.Increment(1);
                //Thread.Sleep(500);
            }
            if (countError>0)
            {
                fc.lblFontFault.Text = countError +" ข้อผิดพลาด";
                fc.btnEdit.Visible = true;
            }
            else
            {
                fc.btnEdit.Visible = false;
                fc.lblFontFault.Text = countError + " ข้อผิดพลาด";
            }
        }

        //Correct Font Size
        public void IndexFaultFontSize(FontUC fc)
        {
            listPb[index].range.Select();
            listPb[index].range.HighlightColorIndex = WdColorIndex.wdAuto;
            fc.lblMainText.Text = listPb[index].comment;
            index++;
        }
        //Correct all font in doc
        public void CorrectFont(string font, FontUC fc)
        {
            Word.Document app = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range range = app.Content;
            range.Font.Name = font;
            range.HighlightColorIndex = WdColorIndex.wdAuto;
            fc.lblFontFault.Text = "---";
            fc.btnEdit.Visible = false;
           // MessageBox.Show("แก้ไขเสร็จสิ้น");
        }

        public void ClearHightLight()
        {
            Word.Document app = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range range = app.Content;
            range.HighlightColorIndex = WdColorIndex.wdAuto;
        }
    }
}
