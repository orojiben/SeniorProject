using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
namespace SeniorProject
{
    class ReferenceModelEN
    {
        //หนังสือทั่วไป เอกสารประเภทหนังสือ
        public int ModelBookTypeBookEN(Word.Range r, string strCheck, int cout)
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
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForYearCreate())
                        {

                        }
                        l.countCutNotBold = l.countLength;
                        if (l.ForPlaceEndColon())
                        {
                            l.countCutNotBold = l.countLength;
                            if (l.ForPublishersEnd())
                            {
                                //System.Windows.Forms.MessageBox.Show("หนังสือทั่วไป เอกสารประเภทหนังสือ");
                                return l.countLength;
                            }
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

                                        //System.Windows.Forms.MessageBox.Show("หนังสือสารานุกรม เอกสารประเภทหนังสือ");
                                        return l.countLength;

                                    }
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
                    if (l.ForBookNameBoldFullStop())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForBookNameCommas2AndFullstopEnd(0))
                        {
                            //System.Windows.Forms.MessageBox.Show("วิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง เอกสารประเภทวิทยานิพนธ์และการศึกษาค้นคว้าด้วยตนเอง");
                            return l.countLength;

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
            return ModelOtherTypeBrochuresAndLeafletsEN(r, strCheck, cout);
        }

        //จุลสารและแผ่นพับ เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeBrochuresAndLeafletsEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
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
            return ModelOtherTypeArchivesEN(r, strCheck, cout);
        }

        //จดหมายเหตุ คำสั่ง ประกาศ แผ่นปลิว เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeArchivesEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracketDate(0))
            {
                //  l.countCutNotBold = l.countLength;
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
                //  }
            }
            return ModelOtherTypeGovernmentGazetteEN(r, strCheck, cout);
        }

        //สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ
        private int ModelOtherTypeGovernmentGazetteEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            if (l.ForBookNameToBracketDate(0))
            {
                //  l.countCutNotBold = l.countLength;
                //   if (l.ForDate())
                //  {
                l.countCutBold = l.countLength;
                if (l.ForBookNameBoldFullStop())
                {
                    l.countCutNotBold = l.countLength;
                    if (l.ForAt())
                    {
                        l.countCutNotBold = l.countLength;
                        if (l.ForPageEnd())
                        {
                            //System.Windows.Forms.MessageBox.Show("สารสนเทศในราชกิจจานุเบกษา เอกสารประเภทสื่อสิ่งพืมพ์อื่นๆ");
                            return l.countLength;
                        }
                    }
                }
                // }
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
                                 //System.Windows.Forms.MessageBox.Show("สื่อประเภทบันทึกเสียง เอกสารประเภทวัสดุไม่ตีพิมพ์");
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
            return ModelOnlineTypeOnlineEN(r, strCheck, cout);
        }

        //บทความออนไลน์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOnlineEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
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
            /* else
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
                                     //System.Windows.Forms.MessageBox.Show("บทความออนไลน์ เอกสารประเภทสื่อออนไลน์");
                                     return l.countLength;
                                 }
                             }

                         }
                     }
                 }
             }*/
            return ModelOnlineTypeOtherEN(r, strCheck, cout);
        }

        //บทความในสื่อออนไลน์ประเภทต่างๆ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeOtherEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
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
            return ModelOnlineTypeElectronicEN(r, strCheck, cout);
        }

        //บทเรียนอิเล็กทรอนิกส์ เอกสารประเภทสื่อออนไลน์
        private int ModelOnlineTypeElectronicEN(Word.Range r, string strCheck, int cout)
        {
            LexerEN l = new LexerEN();
            l.sentence = strCheck;
            l.range = r;
            bool check = false;

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
