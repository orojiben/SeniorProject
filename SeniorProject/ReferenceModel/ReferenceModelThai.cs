using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    class ReferenceModelThai
    {
        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        public int ModelBookTypeBookTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
           // Word.Range r2 = r.Document.Range(r.Start, 10+r.Start);
          /*  foreach (Word.Range range in r.Document.StoryRanges)
            {
                string s = r.Text;
            }*/
            l.range = r;
            if (l.ForNames())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutBold = l.countLength;
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        //(ครั้งที่พิมพ์).
                        if (l.ForYearCreate())
                        {

                        }
                        l.countCutNotBold = l.countLength;
                        if (!l.ForBookTranslator())
                        {
                            return ModelBookTypeArticleTH(r, strCheck, cout);
                        }

                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEndColon())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPublishersEnd())
                            {
                                //l.countCutNotBold = l.countLength;
                                //if (!l.ForBookAddEnd())
                                //{
                                // return ModelBookTypeArticleTH(r, strCheck, cout);
                                //}
                                //System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
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
                                if (l.ForPlaceEndColon())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPublishersEnd())
                                    {
                                        //System.Windows.Forms.MessageBox.Show("บทความในหนังสือ เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
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
                    if (l.ForBookNameToIn())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameInSpace())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageAndBook())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEndColon())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPublishersEnd())
                                    {

                                        //System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return ModelBookTypeHandoutLibraryTH(r, strCheck, cout);
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
                        if (l.ForBookNameToIn())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameInSpace())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageAndBook())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPlaceEndColon())
                                    {
                                        l.countCutNotBold = l.countLength;
                                        if (l.ForPublishersEnd())
                                        {
                                            //System.Windows.Forms.MessageBox.Show("เอกสารประกอบการบรรยาย เอกสารที่จัดพิมพ์รวมเล่มและอ้างเฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                            return l.countLength;
                                        }
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
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldFullStop())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPlaceEndColon())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPublishersEnd())
                                {
                                    //System.Windows.Forms.MessageBox.Show("เอกสารที่จัดพิมพ์เฉพาะเรื่อง เอกสารประเภทหนังสือ");
                                    return l.countLength;
                                }
                            }
                        }
                    }
                }
            }
            return ModelJournalTypeArticlesTH(r, strCheck, cout);
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
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                //System.Windows.Forms.MessageBox.Show("บทความทั่วไป เอกสารประเภทวารสาร");
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
                    if (l.ForBookNameSpaceToSquareBrackets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameFistSquareBracketsOnBold())
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameBoldComma())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    //System.Windows.Forms.MessageBox.Show("บทวิจารณ์และบทความปริทัศน์หนังสือ เอกสารประเภทวารสาร");
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
                    if (l.ForBookNameSpaceToSquareBrackets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameFistSquareBracketsAndName())
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameBoldComma())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForYearAndNumber())
                                {
                                    //System.Windows.Forms.MessageBox.Show("บทสัมภาษณ์ เอกสารประเภทวารสาร");
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
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                //System.Windows.Forms.MessageBox.Show("หนังสือพิมพ์ทั่วไป เอกสารประเภทหนังสือพิมพ์");
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
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                //System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
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
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPageEnd())
                            {
                                //System.Windows.Forms.MessageBox.Show("กรณีบทความมีชื่อคอลัมน์ เอกสารประเภทหนังสือพิมพ์");
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
                    if (l.ForBookNameFullStopToSquareBrackets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameFistSquareBracketsAndName())
                        {
                            l.countCutBold = l.countLength;
                            if (l.ForBookNameBoldComma())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPageEnd())
                                {
                                    //System.Windows.Forms.MessageBox.Show("กรณีอ้างบทสัมภาษณ์จากหนังสือพิมพ์ เอกสารประเภทหนังสือพิมพ์");
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
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameCommas2AndFullstopEnd(0))
                        {
                            //System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                            return l.countLength;
                        }
                        /* if (l.ForBookNameEC())
                         {
                             l.countCutNotBold = l.countLength;
                             if (l.ForBookNameEC())
                             {
                                 l.countCutNotBold = l.countLength;
                                 if (l.ForBookNameEnd())
                                 {
                                     //System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                                     return l.countLength;
                                 }
                             }
                         }*/
                    }
                }

            }
            return ModelThesisTypeThesisAbstractBookTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ **สาราณุกรมเอกสารหนังสือ
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
                        if (l.ForBookNameInSpace())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPage())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPlaceEndColon())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForPublishersEnd())
                                    {
                                        //System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง หนังสือรวมบทคัดย่อ");
                                        return l.countLength;
                                    }
                                }
                            }
                        }
                    }

                }
            }
            return ModelThesisTypeThesisAbstractJournalTH(r, strCheck, cout);
        }

        //วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร **บทความทั่วไปเอกสารประเภทวารสาร
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
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                //System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง สิ่งพิมพ์ประเภทวารสาร");
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
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameComma())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForBookNameToRetrieved())
                                {
                                    l.countCutNotBold = l.countLength;
                                    if (l.ForSearch())
                                    {
                                        l.countCutNotBold = l.countLength;
                                        if (l.ForURL())
                                        {
                                            //System.Windows.Forms.MessageBox.Show("ฐานข้อมูลวิทยานิพนธ์ออนไลน์ เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
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
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldComma())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForYearAndNumber())
                            {
                                //System.Windows.Forms.MessageBox.Show("จดหมายข่าว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
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
            if (l.ForBookNameBoldFullStop())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForYear())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBrochuresAndLeaflets())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEndColon())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPublishersEnd())
                            {
                                //System.Windows.Forms.MessageBox.Show("จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                                return l.countLength;
                            }
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
            if (l.ForBookNameToBracketDate(0))
            {
                // l.countCutNotBold = l.countLength;
                // if (l.ForDate())
                // {
                l.countCutNotBold = l.countLength;
                if (l.ForBookNumber())
                {
                }
                l.countCutBold = l.countLength;
                if (l.ForBookNameBoldFullStopEnd())
                {

                    //System.Windows.Forms.MessageBox.Show("จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                    return l.countLength;

                }
                // }
            }
            return ModelOtherTypeGovernmentGazetteTH(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteTH(Word.Range r, string strCheck, int cout)
        {
            LexerTH l = new LexerTH();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracketDate(0))
            {
                //l.countCutNotBold = l.countLength;
                //if (l.ForDate())
                // {
                l.countCutBold = l.countLength;
                if (l.ForBookNameBoldFullStop())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForAt())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForPageEnd2())
                        {
                            //System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
                //}
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
                    if (l.ForBookNameSpaceBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameFistSquareBrackets())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPlaceEndColon())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForPublishersEnd())
                                {
                                    //System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
                                    return l.countLength;
                                }
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
                    if (l.ForBookNameSpaceBold())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameFistSquareBrackets())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForBookNameDB())
                            {
                                //System.Windows.Forms.MessageBox.Show("ฐานข้อมูลสำเร็จรูป เอกสารประเภทวัสดุไม่ตีพิมพ์");
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
            if (l.ForBookNameBoldFullStop())
            {
                l.countCutNotBold = l.countLength;
                if (l.ForSearch())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForURL())
                    {
                        //System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                        return l.countLength;
                    }
                }

            }
            l.sentence = strCheck;
            l.range = r;
            l.countCutNotBold = 0;
            l.countLength = 0;
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
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForSearch())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForURL())
                            {
                                //System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
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
            if (l.ForBookNameFullStopToBold())
            {
                l.countCutBold = l.countLength;
                if (l.ForBookNameBoldFullStop())
                {

                    l.countCutNotBold = l.countLength;
                    if (l.ForSearch())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForURL())
                        {
                            //System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
                            return l.countLength;
                        }
                    }
                }

            }
            l.sentence = strCheck;
            l.range = r;
            l.countCutNotBold = 0;
            l.countLength = 0;

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
                    else
                    {
                        if (l.ForYear())
                        {
                            check = true;
                        }
                    }
                }
                if (check)
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldFullStop())
                        {

                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    //System.Windows.Forms.MessageBox.Show("บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์");
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

            if (l.ForBookNameFullStopToBold())
            {
                l.countCutBold = l.countLength;
                if (l.ForBookNameBoldFullStop())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForSearch())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForURL())
                        {
                            //System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                            return l.countLength;
                        }
                    }
                }

            }
            l.sentence = strCheck;
            l.range = r;
            l.countCutNotBold = 0;
            l.countLength = 0;
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
                    if (l.ForBookNameFullStopToBold())
                    {
                        l.countCutBold = l.countLength;
                        if (l.ForBookNameBoldFullStop())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForSearch())
                            {
                                l.countCutNotBold = l.countLength;
                                if (l.ForURL())
                                {
                                    //System.Windows.Forms.MessageBox.Show("บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์");
                                    return l.countLength;
                                }
                            }
                        }

                    }
                }
            }
            return 0;
        }

    }
}
