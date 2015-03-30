using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class LexerEN
    {
        public string sentence;
        public int countLength;
        public Word.Range range;
        string[] monthENs;
        int checkForBookName;
        public int countCutBold = 0;
        public int countCutNotBold = 0;
        public LexerEN()
        {
            checkForBookName = 0;
            countLength = 0;

            monthENs = new string[12];
            monthENs[0] = @"^[J][a][n][u][a][r][y]";
            monthENs[1] = @"^[F][e][b][r][u][a][r][y]";
            monthENs[2] = @"^[M][a][r][c][h]";
            monthENs[3] = @"^[A][p][r][i][l]";
            monthENs[4] = @"^[M][a][y]";
            monthENs[5] = @"^[J][u][n][e]";
            monthENs[6] = @"^[J][u][l][y]";
            monthENs[7] = @"^[A][u][g][u][s][t]";
            monthENs[8] = @"^[S][e][p][t][e][m][b][e][r]";
            monthENs[9] = @"^[O][c][t][o][b][e][r]";
            monthENs[10] = @"^[N][o][v][e][m][b][e][r]";
            monthENs[11] = @"^[D][e][c][e][m][b][e][r]";
        }

        private void CheckStringMatch(string strFromRange, string regex, ref int checkValue)
        {
            Match match = Regex.Match(strFromRange, regex);
            if (match.Success)
            {
                checkValue = match.Value.Length;
                return;
            }
            checkValue = -1;
        }

        //1 = Pass 0=fail
        private int CheckString()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return CheckString();
                }

                CheckStringMatch(this.sentence, @"^(\.)?\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                    if (checkValue == -1)
                    {
                        return 0;
                    }
                    return CheckString();
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return CheckString();
                }



            }
            return 1;
        }

        //0 = Pass 1=fail 2=next
        private int CheckStringNotColon()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^(\.)?\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                    if (checkValue == -1)
                    {
                        return 0;
                    }
                    return CheckString();
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return CheckString();
                }


            }
            return 1;
        }

        //สัญประกาศ 0 = ไม่มี 1=มี 2=ผิด 3=fullstop 4=colon
        private int CheckStringQuotationMarks()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^((\“)|(\”))", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckString() == 1)
                {
                    CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 7;
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 1;
                        }
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 3;
                        }
                        CheckStringMatch(this.sentence, @"^(\:\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        CheckStringMatch(this.sentence, @"^(\,\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 5;
                        }
                        CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 6;
                        }
                        CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -2;
                        }

                    }
                    return 2;
                }
                else
                {
                    if (CheckStringSingleQuotationMarks() == 1)
                    {
                        return CheckStringQuotationMarksEnd();
                    }
                    return 2;
                }
            }
            return 0;
        }

        //สัญประกาศ 0 = ไม่มี 1=มี 2=ผิด
        private int CheckStringQuotationMarksEnd()
        {
            int checkValue = -1;

            if (CheckString() == 1)
            {
                CheckStringMatch(this.sentence, @"^(\”)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 7;
                    }
                    CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 1;
                    }
                    CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 3;
                    }
                    CheckStringMatch(this.sentence, @"^(\:\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 4;
                    }
                    CheckStringMatch(this.sentence, @"^(\,\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 5;
                    }
                    CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 6;
                    }
                    CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return -2;
                    }
                }
                return 2;
            }
            else
            {
                if (CheckStringSingleQuotationMarks() == 1)
                {
                    return CheckStringQuotationMarksEnd();
                }
                return 2;
            }
        }

        //สัญประกาศ 0 = ไม่มี 1=มี 2=ผิด 3=fullstop 4=colon
        private int CheckStringParenthesis()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^(\()", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckString() == 1)
                {
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CutString(checkValue);
                        if (this.sentence == "")
                        {
                            return -2;
                        }
                        CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 7;
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 1;
                        }
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 3;
                        }
                        CheckStringMatch(this.sentence, @"^(\:\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        CheckStringMatch(this.sentence, @"^(\,\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 5;
                        }
                        CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 6;
                        }
                        CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -2;
                        }
                    }
                }
                return 2;
            }
            return 0;
        }

        private int CheckStringParenthesisEnd()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^(\()", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckString() == 1)
                {
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return 2;
                    }
                }

            }
            return 0;
        }

        //สัญประกาศ 0 = ไม่มี 1=มี 2=ผิด
        private int CheckStringSingleQuotationMarks()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^((\‘)|(\’))", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckString() == 1)
                {
                    CheckStringMatch(this.sentence, @"^(\’)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 1;
                        }
                    }
                    return 2;
                }
                else
                {
                    return 2;
                }
            }
            return 0;
        }

        private int CheckStringBoldPublicFullStop()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                    {
                        return 2;
                    }
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringBoldPublicFullStopEnd()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);

                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (this.sentence == "")
                    {
                        return 2;
                    }
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicFullStopToBold()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                    {
                        return 2;
                    }
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicFullStopToSquareBrackets()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                    if (checkValue != -1)
                    {
                        return 2;
                    }
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicFullStopToBracketDate()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue != -1)
                    {
                        return 2;
                    }
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicSpace()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 2;
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }


                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringBoldPublicComma()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.)?\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                    {
                        return 2;
                    }
                    return 0;
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicComma()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.)?\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 2;
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicCommaAndFullStop()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.)?\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 2;
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 3;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringBoldPublicSquareBracketsClose()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\].\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 2;
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        //สัญประกาศ 0 = ไม่มี 1=มี 2=จบ
        private int CheckStringPublicFullStopEnd()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (this.sentence == "")
                    {
                        return 2;
                    }
                    this.sentence = "." + this.sentence;
                }

                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicBracketsCloseEnd()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    if (this.sentence == "")
                    {
                        return 2;
                    }
                    else
                    {
                        int checkStringParenthesisEnd = CheckStringParenthesisEnd();
                        if (checkStringParenthesisEnd == 2)
                        {
                            if (this.sentence == "")
                            {
                                return 2;
                            }
                        }
                    }
                    this.sentence = ")" + this.sentence;
                }

                CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        private int CheckStringPublicColon()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[?]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^\,", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^[:]\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 2;
                }

                CheckStringMatch(this.sentence, @"^(\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return 0;
                }

            }
            return 1;
        }

        public bool ForBookNameBoldFullStop()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldFullStop();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldFullStop();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldFullStop();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldFullStop();
            }

            /*CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckString() == 1)
                {
                    CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                            {
                                return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                            }
                            return ForBookNameBold();
                        }
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameBold();
                        }
                    }
                }
                return false;
            }*/
            int valueCheck = CheckStringBoldPublicFullStop();
            if (valueCheck == 0)
            {
                return ForBookNameBoldFullStop();
            }
            else if (valueCheck == 2)
            {
                return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
            }
            return false;
        }

        public bool ForBookNameBoldFullStopEnd()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                if (this.sentence == "")
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldFullStopEnd();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldFullStopEnd();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                if (this.sentence == "")
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldFullStopEnd();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldFullStopEnd();
            }
            int valueCheck = CheckStringBoldPublicFullStopEnd();
            if (valueCheck == 0)
            {
                return ForBookNameBoldFullStopEnd();
            }
            else if (valueCheck == 2)
            {
                return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
            }
            return false;
        }

        public bool ForBookNameFullStopToBold()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameFullStopToBold();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameFullStopToBold();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameFullStopToBold();
            }
            int valueCheck = CheckStringPublicFullStopToBold();
            if (valueCheck == 0)
            {
                return ForBookNameFullStopToBold();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public bool ForBookNameFullStopToSquareBrackets()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameFullStopToSquareBrackets();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameFullStopToSquareBrackets();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameFullStopToSquareBrackets();
            }
            int valueCheck = CheckStringPublicFullStopToSquareBrackets();
            if (valueCheck == 0)
            {
                return ForBookNameFullStopToSquareBrackets();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public bool ForPlaceEndColon()
        {

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 4)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForPlaceEndColon();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 4)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForPlaceEndColon();
            }

            int valueCheck = CheckStringPublicColon();
            if (valueCheck == 0)
            {
                return ForPlaceEndColon();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;

        }

        public bool ForPublishersEnd()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == -2)
            {
                if (this.sentence == "")
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2)
            {
                return false;
            }
            else
            {
                return ForPublishersEnd();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == -2)
            {
                if (this.sentence == "")
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2)
            {
                return false;
            }
            else
            {
                return ForPublishersEnd();
            }
            int valueCheck = CheckStringPublicFullStopEnd();
            if (valueCheck == 0)
            {
                return ForPublishersEnd();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public bool ForPublishersEndBrackets()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 7)
            {
                if (this.sentence == "")
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForPublishersEndBrackets();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 7)
            {
                if (this.sentence == "")
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForPublishersEndBrackets();
            }
            int valueCheck = CheckStringPublicBracketsCloseEnd();
            if (valueCheck == 0)
            {
                return ForPublishersEndBrackets();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public bool ForBookNameSpaceBold()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 1)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameSpaceBold();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpaceBold();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 1)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameSpaceBold();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpaceBold();
            }

            int valueCheck = CheckStringPublicSpace();
            if (valueCheck == 0)
            {
                return ForBookNameSpaceBold();
            }
            else if (valueCheck == 2)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameSpaceBold();
            }
            return false;
        }

        public bool ForBookNameInSpace()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([I][n]\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (!CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                {
                    return false;
                }
                this.countCutBold = this.countLength;
                if (ForBookNameSpaceBold())
                {
                    return true;
                }

            }

            return false;
        }

        public bool ForBookNameBoldComma()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 5)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldComma();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldComma();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 5)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
                }
                return ForBookNameBoldComma();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldComma();
            }

            int valueCheck = CheckStringBoldPublicComma();
            if (valueCheck == 0)
            {
                return ForBookNameBoldComma();
            }
            else if (valueCheck == 2)
            {
                return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 2 + range.Start));
            }
            return false;
        }

        public bool ForBookNameSpaceToSquareBrackets()
        {
            int checkValue = -1;
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 1)
            {
                CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                if (checkValue != -1)
                {
                    //CutString(checkValue);
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpaceToSquareBrackets();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpaceToSquareBrackets();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 1)
            {
                CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                if (checkValue != -1)
                {
                    //CutString(checkValue);
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpaceToSquareBrackets();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpaceToSquareBrackets();
            }

            int valueCheck = CheckStringPublicSpace();
            if (valueCheck == 0)
            {
                return ForBookNameSpaceToSquareBrackets();
            }
            else if (valueCheck == 2)
            {
                CheckStringMatch(this.sentence, @"^\[", ref checkValue);
                if (checkValue != -1)
                {
                    //CutString(checkValue);
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpaceToSquareBrackets();
            }
            return false;
        }

        public bool ForBookNameFistSquareBracketsOnBold()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                {
                    this.countCutNotBold = this.countLength;
                    if (this.ForBookNameSpace())
                    {
                        this.countCutBold = this.countLength;
                        if (this.ForBookNameBoldSquareBracketsClose())
                        {
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public bool ForBookNameSpace()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 1)
            {
                if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpace();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpace();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 1)
            {
                if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpace();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSpace();
            }

            int valueCheck = CheckStringPublicSpace();
            if (valueCheck == 0)
            {
                return ForBookNameSpace();
            }
            else if (valueCheck == 2)
            {
                if (CheckBold(this.range.Document.Range(this.countLength + range.Start, this.countLength + 1 + range.Start)))
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameSpace();
            }
            return false;
        }

        public bool ForBookNameBoldSquareBracketsClose()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 6)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength - 3 + range.Start, this.countLength - 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 3 + range.Start));
                }
                return ForBookNameBoldSquareBracketsClose();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldSquareBracketsClose();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 6)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength - 3 + range.Start, this.countLength - 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 3 + range.Start));
                }
                return ForBookNameBoldSquareBracketsClose();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameBoldSquareBracketsClose();
            }

            int valueCheck = CheckStringBoldPublicSquareBracketsClose();
            if (valueCheck == 0)
            {
                return ForBookNameBoldSquareBracketsClose();
            }
            else if (valueCheck == 2)
            {
                if (CheckNotBold(this.range.Document.Range(this.countLength - 3 + range.Start, this.countLength - 1 + range.Start)))
                {
                    return CheckBold(this.range.Document.Range(this.countCutBold + range.Start, this.countLength - 3 + range.Start));
                }
            }
            return false;
        }

        public bool ForBookNameFistSquareBracketsAndName()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                {
                    this.countCutNotBold = this.countLength;
                    if (this.ForBookNameSquareBracketsCloseAndName())
                    {
                        return true;
                    }
                }
            }
            return false;
        }

        public bool ForBookNameSquareBracketsCloseAndName()
        {
            int checkValue = -1;
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 5)
            {
                CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[t][h][e]\s[e][d][i][t][o][r]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
                //return ForBookNameSquareBracketsClose();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSquareBracketsCloseAndName();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 5)
            {
                CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[t][h][e]\s[e][d][i][t][o][r]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
                //return ForBookNameSquareBracketsClose();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSquareBracketsCloseAndName();
            }

            int valueCheck = CheckStringPublicSpace();
            if (valueCheck == 0)
            {
                return ForBookNameSquareBracketsCloseAndName();
            }
            else if (valueCheck == 2)
            {
                CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[t][h][e]\s[e][d][i][t][o][r]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
                return ForBookNameSquareBracketsCloseAndName();
            }
            return false;
        }

        public bool ForBookNameCommas2AndFullstopEnd(int checkLoop)
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 5)
            {
                return ForBookNameCommas2AndFullstopEnd(++checkLoop);
            }
            else if (checkStringQuotationMarks == -2)
            {
                if (checkLoop >= 2)
                {
                    if (this.sentence == "")
                    {
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2)
            {
                return false;
            }
            else
            {
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 5)
            {
                return ForBookNameCommas2AndFullstopEnd(++checkLoop);
            }
            else if (checkStringQuotationMarks == -2)
            {
                if (checkLoop >= 2)
                {
                    if (this.sentence == "")
                    {
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2)
            {
                return false;
            }
            else
            {
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }

            int valueCheck = CheckStringPublicCommaAndFullStop();
            if (valueCheck == 0)
            {
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }
            else if (valueCheck == 2)
            {
                return ForBookNameCommas2AndFullstopEnd(++checkLoop);
            }
            else if (valueCheck == 3)
            {
                if (checkLoop >= 2)
                {
                    if (this.sentence == "")
                    {
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                return ForBookNameCommas2AndFullstopEnd(checkLoop);
            }
            return false;
        }

        public bool ForBookNameComma()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 5)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameComma();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 5)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameComma();
            }

            int valueCheck = CheckStringPublicComma();
            if (valueCheck == 0)
            {
                return ForBookNameComma();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public bool ForBookNameToRetrieved()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                CheckStringMatch(this.sentence, @"^[R][e][t][r][i][e][v][e][d]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToRetrieved();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameToRetrieved();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                CheckStringMatch(this.sentence, @"^[R][e][t][r][i][e][v][e][d]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToRetrieved();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameToRetrieved();
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์A-Za-z])+");
            int valueCheck = CheckStringBoldPublicFullStop();
            if (valueCheck == 0)
            {
                return ForBookNameToRetrieved();
            }
            else if (valueCheck == 2)
            {
                CheckStringMatch(this.sentence, @"^[R][e][t][r][i][e][v][e][d]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToRetrieved();
            }
            return false;
        }

        public bool ForBookNameToBracketDate(int checkLoop)
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;

            if (checkLoop > 0)
            {
                int checkStringQuotationMarks = CheckStringQuotationMarks();
                if (checkStringQuotationMarks == 3)
                {
                    int copyCountCutNotBold = this.countCutNotBold;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue != -1)
                    {
                        if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                        {
                            this.countCutNotBold = this.countLength;
                            if (this.ForDate())
                            {
                                return true;
                            }
                            this.countCutNotBold = copyCountCutNotBold;
                            return ForBookNameToBracketDate(checkLoop);
                        }
                        return false;
                    }
                    this.countCutNotBold = copyCountCutNotBold;
                    return ForBookNameToBracketDate(checkLoop);
                }
                else if (checkStringQuotationMarks == 0)
                {

                }
                else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
                {
                    return false;
                }
                else
                {
                    return ForBookNameToBracketDate(checkLoop);
                }

                int checkStringParenthesis = CheckStringParenthesis();
                if (checkStringParenthesis == 3)
                {
                    int copyCountCutNotBold = this.countCutNotBold;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue != -1)
                    {
                        if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                        {
                            this.countCutNotBold = this.countLength;
                            if (this.ForDate())
                            {
                                return true;
                            }
                            this.countCutNotBold = copyCountCutNotBold;
                            return ForBookNameToBracketDate(checkLoop);
                        }
                        return false;
                    }
                    this.countCutNotBold = copyCountCutNotBold;
                    return ForBookNameToBracketDate(checkLoop);
                }
                else if (checkStringParenthesis == 0)
                {

                }
                else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
                {
                    return false;
                }
                else
                {
                    return ForBookNameToBracketDate(checkLoop);
                }
            }
            int valueCheck = CheckStringPublicFullStopToBracketDate();
            if (valueCheck == 0)
            {
                return ForBookNameToBracketDate(checkLoop++);
            }
            else if (valueCheck == 2)
            {
                int copyCountCutNotBold = this.countCutNotBold;
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                    {
                        this.countCutNotBold = this.countLength;
                        if (this.ForDate())
                        {
                            return true;
                        }
                        this.countCutNotBold = copyCountCutNotBold;
                        return ForBookNameToBracketDate(++checkLoop);
                    }
                    return false;
                }
                this.countCutNotBold = copyCountCutNotBold;
                return ForBookNameToBracketDate(++checkLoop);
            }
            return false;
        }

        public bool ForBookNameFistSquareBrackets()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                {

                    this.countCutBold = this.countLength;
                    if (this.ForBookNameSquareBracketsClose())
                    {
                        return true;
                    }

                }
            }
            return false;
        }

        public bool ForBookNameSquareBracketsClose()
        {
            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 6)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSquareBracketsClose();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 6)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameSquareBracketsClose();
            }

            int valueCheck = CheckStringBoldPublicSquareBracketsClose();
            if (valueCheck == 0)
            {
                return ForBookNameSquareBracketsClose();
            }
            else if (valueCheck == 2)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            return false;
        }

        public int ForName()
        {
            int valueName2 = ForName2();
            if (valueName2 != -1)
            {
                return valueName2;
            }
            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsEN = ForNamelistInitialsEN();
                    if (valueForNamelistInitialsEN == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsEN >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 3;
                    }
                }
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return 4;
                    }
                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return 2;
                                }
                                return 1;
                            }
                            return 2;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsEN = ForNamelistInitialsEN();
            if (valueForNamelistInitialsEN == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsEN >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        return 2;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public int ForNameDontAnd()
        {
            int valueName2 = ForName2DontAnd();
            if (valueName2 != -1)
            {
                return valueName2;
            }

            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsEN = ForNamelistInitialsEN();
                    if (valueForNamelistInitialsEN == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsEN >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 8888;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return -1;
                            }
                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return -1;
                                }
                                return 1;
                            }
                            return -1;
                        }
                        return 1;
                    }
                    return -1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2DontAnd()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsEN = ForNamelistInitialsEN();
            if (valueForNamelistInitialsEN == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsEN >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -1;
                        }
                        return -1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public bool ForNames()
        {
            int checkValue = -1;
            int checkValueFormate = ForName();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAnd();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNames();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return ForNames();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }

            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNames();
            }

            if (checkValueFormate == 3)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }


            return false;
        }

        public bool ForNamesForCheck()
        {
            int checkValue = -1;
            int checkValueFormate = ForName();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAnd();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return true;
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNamesForCheck();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAnd();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return ForNamesForCheck();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return true;
            }

            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNamesForCheck();
            }

            if (checkValueFormate == 3)
            {
                return true;
            }


            return false;
        }

        public int ForNamelistInitialsEN()
        {
            Match match;

            for (char chars = 'A'; chars <= 'Z'; chars++)
            {
                match = Regex.Match(this.sentence, "^" + chars);
                if (match.Success)
                {
                    string listInitialsENNew = "^" + chars + @"([a-z])?\.\s";
                    match = Regex.Match(this.sentence, listInitialsENNew);
                    if (match.Success)
                    {
                        CutString(match.Length);
                        return ForNamelistInitialsEN() + 1;
                    }

                    listInitialsENNew = "^" + chars + @"([a-z])?\.\,\s";
                    match = Regex.Match(this.sentence, listInitialsENNew);
                    if (match.Success)
                    {
                        CutString(match.Length);
                        return 9999;
                    }
                }
            }


            return 0;
        }

        public bool ForYear()
        {
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^(\()", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9][0-9]{3})+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([A-Za-z])*", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([n]\.[d]\.)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
            }
            return false;
        }

        public bool ForYearForCheck()
        {
            int checkValue = -1;
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9][0-9]{3})+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([n]\.[d]\.)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([A-Za-z]\.)+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            return false;
        }

        public bool ForYearCreate()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\)\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
            }
            this.sentence = sentenceCopy;
            this.countLength = countLengthCopy;
            return false;
        }


        public bool ForBookName()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, "^[’]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookName();
                        }
                    }
                    return false;
                }
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^(\))", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookName();
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookName();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookName();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookName();
            }
            return false;
        }

        public bool ForBookNameIn()
        {
            int countCutNotBoldIn = this.countLength;
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([I][n]\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                if (this.ForNamesNF())
                {
                    CheckStringMatch(this.sentence, @"^(\()+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([E][d]([s])?)\.", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\)\,\s)+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);

                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
                        }
                    }
                    else
                    {
                        this.sentence = sentenceCopy;
                        this.countLength = countLengthCopy;
                    }
                }

                if (!CheckNotBold(this.range.Document.Range(countCutNotBoldIn + range.Start, this.countLength + range.Start)))
                {
                    return false;
                }
                this.countCutBold = this.countLength;
                if (ForBookNameSpaceBold())
                {
                    return true;
                }

            }

            return false;
        }



        public int ForNameNF()
        {
            int valueName2 = ForName2NF();
            if (valueName2 != -1)
            {
                return valueName2;
            }
            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsEN = ForNamelistInitialsEN();
                    if (valueForNamelistInitialsEN == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsEN >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 3;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }
                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    return 1;
                                }
                                return 1;
                            }
                            return 1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2NF()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsEN = ForNamelistInitialsEN();
            if (valueForNamelistInitialsEN == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsEN >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return 4;
                        }
                        return 1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public int ForNameDontAndNF()
        {
            int valueName2 = ForName2DontAndNF();
            if (valueName2 != -1)
            {
                return valueName2;
            }

            int checkValue = -1;


            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsEN = ForNamelistInitialsEN();
                    if (valueForNamelistInitialsEN == 0)
                    {
                        return 0;
                    }
                    else if (valueForNamelistInitialsEN >= 9999)
                    {
                        return 9999;
                    }
                    else
                    {
                        return 8888;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }
                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return -1;
                            }
                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    return 1;
                                }
                                return 1;
                            }
                            return 1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return 1;
            }
            return -1;
        }

        public int ForName2DontAndNF()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsEN = ForNamelistInitialsEN();
            if (valueForNamelistInitialsEN == 0)
            {
                return -1;
            }
            else if (valueForNamelistInitialsEN >= 9999)
            {
                this.sentence = sentenceCoppy;
                return -1;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return -1;
                        }
                        return 1;
                    }
                    return 1;
                }
                return -1;
            }

        }

        public bool ForNamesNF()
        {
            int checkValue = -1;
            int checkValueFormate = ForNameNF();
            if (checkValueFormate == -1)
            {
                return false;
            }
            else if (checkValueFormate == 4)
            {
                checkValueFormate = ForNameDontAndNF();
                if (checkValueFormate == 1)
                {
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue != -1)
                    {
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
                else if (checkValueFormate == 8888)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return false;
            }
            else if (checkValueFormate == 0 || checkValueFormate == 9999)
            {
                return ForNamesNF();
            }
            else if (checkValueFormate == 2)
            {
                CheckStringMatch(this.sentence, @"^[a][n][d]\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkValueFormate = ForNameDontAndNF();
                    if (checkValueFormate == 1)
                    {
                        CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                        if (checkValue != -1)
                        {
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^\,\s");
            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                return ForNamesNF();
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {

                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }

            CheckStringMatch(this.sentence, @"^[a][n][d]\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNamesNF();
            }

            if (checkValueFormate == 3)
            {
                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
            }


            return false;
        }

        public bool ForPage()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                bool checkNumber = false;
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(([4-9]|([1-9]([0-9])+))[t][h]\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkNumber = true;
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1])", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([s][t]\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            checkNumber = true;
                        }
                        else
                        {
                            return false;
                        }
                    }
                    else
                    {
                        CheckStringMatch(this.sentence, @"^([2])", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([n][d]\s)", ref checkValue);
                            if (checkValue != -1)
                            {
                                checkNumber = true;
                            }
                            else
                            {
                                return false;
                            }
                        }
                        else
                        {
                            CheckStringMatch(this.sentence, @"^([3])", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([r][d]\s)", ref checkValue);
                                if (checkValue != -1)
                                {
                                    checkNumber = true;
                                }
                                else
                                {
                                    return false;
                                }
                            }
                        }
                    }
                }

                if (checkNumber)
                {
                    CheckStringMatch(this.sentence, @"^([e][d]\.\,\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    else
                    {
                        return false;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([V][o][l]\.[1-9]([0-9])*\,\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    else
                    {
                        return false;
                    }
                }

                CheckStringMatch(this.sentence, @"^([p]\.\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([p][p]\.\s)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^((\s)?\-(\s)?[1-9]([0-9])*)", ref checkValue);
                            if (checkValue != -1)
                            {

                                CutString(checkValue);
                            }
                            else
                            {
                                return false;
                            }
                            CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                            if (checkValue != -1)
                            {

                                CutString(checkValue);
                                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                            }
                        }
                    }
                }
            }
            return false;
        }


        public bool ForNarrator()
        {
            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^(\()", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([N][a][r][r][a][t][o][r])", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }

            }
            return false;
        }

        public bool ForDate()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^(\()", ref checkValue);
            int countLengthCopy = this.countLength;
            string sentenceCopy = this.sentence;
            if (checkValue != -1)
            {
                CutString(checkValue);

                checkValue = ForNamemonthEN();
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                    }
                                    CheckStringMatch(this.sentence, @"^\)\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^[N][o]\.\s", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            CheckStringMatch(this.sentence, @"^([1-9][0-9]*)\.\s", ref checkValue);

                                            if (checkValue != -1)
                                            {
                                                CutString(checkValue);
                                                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            checkValue = ForNamemonthEN();
                            if (checkValue > 0)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        CheckStringMatch(this.sentence, @"^([A-Za-z]\.)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                    }
                                }
                            }
                            else
                            {

                            }
                        }
                    }
                }
            }

                this.sentence = sentenceCopy;
                this.countLength = countLengthCopy;
                this.countCutNotBold = this.countLength;
                return this.ForYear();
        }

        public bool ForDateForCheck()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                checkValue = ForNamemonthEN();
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            checkValue = ForNamemonthEN();
                            if (checkValue > 0)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        CheckStringMatch(this.sentence, @"^([A-Za-z]\.)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        return true;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public int ForNamemonthEN()
        {
            Match match;
            foreach (string monthEN in monthENs)
            {
                match = Regex.Match(this.sentence, monthEN);
                if (match.Success)
                {
                    return match.Length;
                }
            }


            return -1;
        }

        public bool ForYearAndNumber()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([0-9]([0-9])*)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([0-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[0-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([A-Z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^\)\,\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                            if (checkValue != -1)
                            {

                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                                if (checkValue != -1)
                                {

                                    CutString(checkValue);
                                }
                                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                                if (checkValue != -1)
                                {

                                    CutString(checkValue);
                                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                }
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool ForAt()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            else
                            {
                                return false;
                            }
                        }
                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));

                        }
                    }
                }
            }
            return false;
        }

        public bool ForPageEnd()
        {
            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
            if (checkValue != -1)
            {

                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([A-Z])", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([A-Z])", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    return true;
                }
            }


            return false;
        }

        public bool ForColumnEnd()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            if (this.ForPlaceEndColon())
            {
                this.countCutNotBold = this.countLength;
                if (this.ForBookNameFullStopToBold())
                {
                    return true;
                }
            }

            this.sentence = sentenceCopy;
            this.countLength = countLengthCopy;
            return false;
        }

        public bool ForBookNameToIn()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;

            int checkStringQuotationMarks = CheckStringQuotationMarks();
            if (checkStringQuotationMarks == 3)
            {
                CheckStringMatch(this.sentence, @"^\[I][n]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToIn();
            }
            else if (checkStringQuotationMarks == 0)
            {

            }
            else if (checkStringQuotationMarks == 2 || checkStringQuotationMarks == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameToIn();
            }

            int checkStringParenthesis = CheckStringParenthesis();
            if (checkStringParenthesis == 3)
            {
                CheckStringMatch(this.sentence, @"^\[I][n]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToIn();
            }
            else if (checkStringParenthesis == 0)
            {

            }
            else if (checkStringParenthesis == 2 || checkStringParenthesis == -2)
            {
                return false;
            }
            else
            {
                return ForBookNameToIn();
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์A-Za-z])+");
            int valueCheck = CheckStringBoldPublicFullStop();
            if (valueCheck == 0)
            {
                return ForBookNameToIn();
            }
            else if (valueCheck == 2)
            {
                CheckStringMatch(this.sentence, @"^[I][n]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
                return ForBookNameToIn();
            }
            return false;
        }

        public bool ForSearch()
        {

            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^[R][e][t][r][i][e][v][e][d]", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                        if (checkValue != -1)
                        {

                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            checkValue = ForNamemonthEN();
                            if (checkValue > 0)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            return true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else
                    {
                        checkValue = ForNamemonthEN();
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                                if (checkValue != -1)
                                {

                                    CutString(checkValue);
                                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                    }
                                    CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                                    if (checkValue != -1)
                                    {
                                        CutString(checkValue);
                                        CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                                        if (checkValue != -1)
                                        {
                                            CutString(checkValue);
                                            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                                            if (checkValue != -1)
                                            {
                                                CutString(checkValue);
                                                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
                
            }
            return false;
        }

        public bool ForURL()
        {

            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^[f][r][o][m]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([ก-ฮะ-์0-9a-zA-z./:=]|(\[)(\])(\))(\())+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                    }
                }
            }

            return false;
        }

        public bool ForMonthYear()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {

                CutString(checkValue);
                checkValue = ForNamemonthEN();
                if (checkValue > 0)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\-", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValue = ForNamemonthEN();
                        if (checkValue > 0)
                        {
                            CutString(checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                            }
                        }
                    }


                }
            }
            return false;
        }

        public bool ForBookNameToBracketForCheck()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                    CheckStringMatch(this.sentence, "^([’])", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForBookNameToBracketForCheck();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/?])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToBracketForCheck();
                }
                if (this.checkForBookName == 0)
                {
                    CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                    if (checkValue != -1)
                    {
                        this.checkForBookName++;
                        CutString(checkValue);
                    }
                }

                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameToBracketForCheck();
            }
            CheckStringMatch(this.sentence, @"^(\.\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^(\?\s\()", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return true;
            }

            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketForCheck();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracketForCheck();
            }

            return false;
        }

        public bool ForBrochuresAndLeaflets()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\[)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            bool checkPass = false;
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^[B][r][o][c][h][u][r][e]", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkPass = true;
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^[F][l][a][p]", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkPass = true;
                    }
                }
            }
            if (checkPass)
            {
                CheckStringMatch(this.sentence, @"^\]\.\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
            }
            return false;
        }

        public bool ForNamePrevious()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(([A-Za-z])+(\s)?)+", ref checkValue));
                var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    while (true)
                    {
                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^([,.])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^([ก-ฮะ-์])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            continue;
                        }
                        break;
                    }
                   
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                    CheckStringMatch(this.sentence, @"^\)\;\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                        if (checkValue != -1)
                        {
                            return false;
                        }
                        return true;
                    }
                }
            }
            return false;
        }

        public bool ForNameYear()
        {
            string sentanceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^([A-Za-z]\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            this.countLength = countLengthCopy;
            this.sentence = sentanceCopy;   
            return ForDate();
        }

        public bool ForNameYearForCheck()
        {
            string sentanceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^([A-Za-z]\.)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                }
                else
                {
                    CheckStringMatch(this.sentence, @"^([1-9]([0-9]){3})", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\)\.\s+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }
                    }
                }
            }
            this.countLength = countLengthCopy;
            this.sentence = sentanceCopy;
            return ForDateForCheck();
        }

        public bool ForNameOne()
        {
            bool valueName2 = ForNameOne2();
            if (valueName2)
            {
                return valueName2;
            }
            int checkValue = -1;

            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    int valueForNamelistInitialsEN = ForNamelistInitialsEN();
                    if (valueForNamelistInitialsEN == 0)
                    {
                        return false;
                    }
                    else if (valueForNamelistInitialsEN >= 9999)
                    {
                        return false;
                    }
                    else
                    {
                        return true;
                    }
                }

                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return false;
                    }
                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return false;
                            }
                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
                                    return true;
                                }
                                return false;
                            }
                            return true;
                        }
                        return false;
                    }
                    return true;
                }
                return false;
            }
            return false;
        }

        public bool ForNameOne2()
        {
            string sentenceCoppy = this.sentence;
            int checkValue = -1;

            int valueForNamelistInitialsEN = ForNamelistInitialsEN();
            if (valueForNamelistInitialsEN == 0)
            {
                return false;
            }
            else if (valueForNamelistInitialsEN >= 9999)
            {
                this.sentence = sentenceCoppy;
                return false;
            }
            else
            {
                CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                    return false;
                }
                return false;
            }

        }

        public bool ForNameOnePrevious()
        {
            if (ForNameOne())
            {
                if (ForNamePrevious())
                {
                    int checkValue = -1;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue == -1)
                    {

                        return ForNameOnePrevious();
                    }
                    return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                }
            }
            return false;
        }

        public bool ForNameOnePreviousForCheck()
        {
            if (ForNameOne())
            {
                if (ForNamePrevious())
                {
                    int checkValue = -1;
                    CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                    if (checkValue == -1)
                    {

                        return ForNameOnePreviousForCheck();
                    }
                    return true;
                }
            }
            return false;
        }

        public bool ForBookNameDB()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                if (CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start)))
                {
                    this.countCutNotBold = this.countLength;
                    if (this.ForPlaceEndColon())
                    {
                        this.countCutNotBold = this.countLength;
                        if (this.ForPublishersEndBrackets())
                        {
                            return true;
                        }

                    }
                }
            }
            return false;
        }

        public bool ForBookNumber()
        {
            int checkValue = -1;
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([N][o]\.)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^[1-9][0-9]*", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return CheckNotBold(this.range.Document.Range(this.countCutNotBold + range.Start, this.countLength + range.Start));
                        }
                    }
                }
            }
            return false;
        }

        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
        }

        bool CheckBold(Word.Range range)
        {
            if (range.Bold == -1)
            {
                return true;
            }
            return false;
        }
        bool CheckNotBold(Word.Range range)
        {
            if (range.Bold == 0)
            {
                return true;
            }
            return false;
        }
    }
}
