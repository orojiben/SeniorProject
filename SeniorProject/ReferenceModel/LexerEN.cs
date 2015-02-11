using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SeniorProject
{
    class LexerEN
    {
        public string sentence;
        public int countLength;
        string[] monthENs;
        int checkForBookName;
        bool checkC;//วงเล็บห
        public LexerEN()
        {
            checkC = false;
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

                return ForNames();
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
                return ForNames();
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
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([1-9][0-9]{3})+", ref checkValue);
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
                else
                {
                    CheckStringMatch(this.sentence, @"^([A-Za-z].)+", ref checkValue);
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
                        return true;
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
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

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
                return true;
            }
            CheckStringMatch(this.sentence, @"^(\?\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookName();
            }
            return false;
        }

        public bool ForBookNameEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForBookNameEnd();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            if (!checkC)
            {
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    checkC = true;
                }
            }

            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEnd();
                }

                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEnd();
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
                return ForBookNameEnd();
            }


            if (checkC)
            {
                CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                else
                {
                    return false;
                }
            }


            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence == "")
                {
                    return true;
                }
                return ForBookNameEnd();
            }

            CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                if (this.sentence != "")
                {
                    return ForBookNameEnd();
                }
            }

            return false;
        }

        public bool ForPlaceEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^(([‘]))", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForPlaceEnd();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForPlaceEnd();
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

                CheckStringMatch(this.sentence, @"^((\,)?\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }

                return ForPlaceEnd();
            }
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1 && this.checkForBookName > 0)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return true;
            }
            return false;
        }

        public bool ForBookTranslator()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                while (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);

                        CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                        }
                        CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return true;
                        }

                        CheckStringMatch(this.sentence, @"^(\,\s)+", ref checkValue);
                        if (checkValue != -1)
                        {
                            continue;
                        }
                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    }
                    else
                    {
                        return false;
                    }
                }
            }
            return true;
        }

        public bool ForBookAddEnd()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\s", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([๐-๙0-9ก-ฮะ-์-])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^[ฯ]", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
                            CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }

                            CheckStringMatch(this.sentence, @"^\)", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return true;
                            }

                            CheckStringMatch(this.sentence, @"^(\,\s)+", ref checkValue);
                            if (checkValue != -1)
                            {
                                continue;
                            }
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            return false;
                        }
                    }
                }
            }
            return true;
        }

        public bool ForBookNameIn()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([I][n]\s)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                string sentenceCopy = this.sentence;
                int countLengthCopy = this.countLength;
                if (ForNamesNF())
                {
                    CheckStringMatch(this.sentence, @"^(\()+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([E][d][s])\.", ref checkValue);
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



                if (ForBookNameNF())
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
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
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
                                CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                                if (checkValue != -1)
                                {
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

                return ForNamesNF();
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {

                return true;
            }

            CheckStringMatch(this.sentence, @"^[a][n][d]\s", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForNamesNF();
            }

            if (checkValueFormate == 3)
            {
                return true;
            }


            return false;
        }

        public bool ForBookNameNF()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                            }
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
                            return ForBookNameNF();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameNF();
                }
                CheckStringMatch(this.sentence, @"^\.", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
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
                return ForBookNameNF();
            }
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^\,\s\(", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue - 1);
                this.checkForBookName = 0;
                return true;
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
                            return true;
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
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }

        public bool ForBookNameInDotEditor()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^([I][n]\s)", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);

                if (ForBookNameNF())
                {
                    return true;
                }

            }

            return false;
        }

        public bool ForPageAndBook()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\(", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));

            if (checkValue != -1)
            {
                CutString(checkValue);
                string sentenceCoppy = this.sentence;
                if (!ForBookNameEC())
                {
                    this.sentence = sentenceCoppy;
                }

                CheckStringMatch(this.sentence, @"^([ห][น][้][า]\s)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                }
                else
                {
                    return false;
                }

                CheckStringMatch(this.sentence, @"^([1-9]([0-9])*)", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\-[1-9]([0-9])*)", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                    }
                    CheckStringMatch(this.sentence, @"^\)\.\s", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);
                        return true;
                    }
                }

            }
            return false;
        }

        public bool ForBookNameEC()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForBookNameEC();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameEC();
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
                return ForBookNameEC();
            }

            CheckStringMatch(this.sentence, @"^\.\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameEC();
            }


            CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return true;
            }

            CheckStringMatch(this.sentence, @"^\.", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameEC();
            }

            return false;
        }

        public bool ForBookNameInitials()
        {
            if (ForBookNameEC())
            {
                int checkValue = -1;
                CheckStringMatch(this.sentence, @"^([A-Z][a-z]?\.(\s)?)+", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
                    }
                    return false;
                }
                return true;
            }
            return false;
        }

        public bool ForNarrator()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^(\()", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
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
                        return true;
                    }
                }

            }
            return false;
        }

        public bool ForDate()
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


            return 0;
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
                                    return true;
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
                            return true;

                        }
                    }
                }
            }
            return false;
        }

        public bool ForBookNameReview(int checkSame)
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\[", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1 || checkSame > 0)
            {
                if (checkSame == 0)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, "^([‘])", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                    if (checkValue != -1)
                    {

                        CutString(checkValue);

                        CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        while (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                                return ForBookNameReview(1);
                            }
                        }
                        return false;
                    }
                }
                //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return ForBookNameReview(1);
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
                    return ForBookNameReview(1);
                }
                CheckStringMatch(this.sentence, @"^(\]\.\s)", ref checkValue);

                if (checkValue != -1)
                {
                    CutString(checkValue);
                    this.checkForBookName = 0;
                    return true;
                }
            }
            return false;
        }

        public int ForNameInterviewer()
        {
            int valueName2 = ForNameInterviewer2();
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
                    CheckStringMatch(this.sentence, @"^([a][n][d])\n", ref checkValue);
                    if (checkValue != -1)
                    {
                        return -1;
                    }

                    CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return 1;
                    }

                    CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^\s", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^([a][n][d])\n", ref checkValue);
                            if (checkValue != -1)
                            {
                                CutString(checkValue);
                                return 4;
                            }

                            CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                return 1;
                            }

                            CheckStringMatch(this.sentence, @"^[A-Z]([A-Za-z])+", ref checkValue);
                            if (checkValue != -1)
                            {
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

        public int ForNameInterviewer2()
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

        public int ForNameInterviewerDontAnd()
        {
            int valueName2 = ForNameInterviewer2DontAnd();
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

                    CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        return 1;
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

                            CheckStringMatch(this.sentence, @"^[b][y]\s", ref checkValue);
                            if (checkValue != -1)
                            {
                                return 1;
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

        public int ForNameInterviewer2DontAnd()
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
                        CheckStringMatch(this.sentence, @"^([a][n][d])\n", ref checkValue);
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

        public bool ForNamesInterviewer(int checkSame)
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^\[", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1 || checkSame > 0)
            {
                if (checkSame == 0)
                {
                    CutString(checkValue);
                }
                int checkValueFormate = ForNameInterviewer();
                if (checkValueFormate == -1)
                {
                    return false;
                }
                else if (checkValueFormate == 4)
                {
                    checkValueFormate = ForNameInterviewerDontAnd();
                    if (checkValueFormate == 1)
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
                                    return true;
                                }
                            }
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
                    return ForNamesInterviewer(1);
                }
                else if (checkValueFormate == 2)
                {
                    CheckStringMatch(this.sentence, @"^[a][n][d]\s", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        checkValueFormate = ForNameInterviewerDontAnd();
                        if (checkValueFormate == 1)
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
                                        return true;
                                    }
                                }
                            }
                        }
                        return false;
                    }
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
                                return true;
                            }
                        }
                    }
                }

                //Match match = Regex.Match(this.sentence, @"^\,\s");
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
                            return true;
                        }
                    }
                    return ForNamesInterviewer(1);
                }
                /* CheckStringMatch(this.sentence, @"^\(", ref checkValue);
                 if (checkValue != -1)
                 {
                     return true;
                 }*/

                CheckStringMatch(this.sentence, @"^[a][n][d]\s", ref checkValue);
                if (checkValue != -1)
                {
                    return ForNamesInterviewer(1);
                }

            }
            return false;
        }

        public bool ForBookNameES()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForBookNameES();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameES();
                }
                CheckStringMatch(this.sentence, @"^\,\s", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
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
                return ForBookNameES();
            }
            CheckStringMatch(this.sentence, @"^\[", ref checkValue);

            if (checkValue != -1)
            {
                this.checkForBookName = 0;
                return true;
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

        public bool ForPageEnd2()
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^([ห][น][้][า]\s)", ref checkValue);
            if (checkValue != -1)
            {

                CutString(checkValue);
            }
            else
            {
                return false;
            }
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
                    return true;
                }
            }


            return false;
        }

        public bool ForColumnEnd()
        {
            string sentenceCopy = this.sentence;
            int countLengthCopy = this.countLength;
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);

                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
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
                            return ForColumnEnd();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                        }
                        else
                        {
                            this.sentence = sentenceCopy;
                            this.countLength = countLengthCopy;
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
                            return true;
                        }
                        CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                        if (checkValue != -1)
                        {
                            CutString(checkValue);
                            return ForColumnEnd();
                        }
                    }
                    this.sentence = sentenceCopy;
                    this.countLength = countLengthCopy;
                    return false;
                }
            }

            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForColumnEnd();
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
                return ForColumnEnd();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1 && this.checkForBookName > 0)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return true;
            }
            this.sentence = sentenceCopy;
            this.countLength = countLengthCopy;
            return false;
        }

        public bool ForBookNameToIn()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);

                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForBookNameToIn();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToIn();
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
                return ForBookNameToIn();
            }
            CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);

            if (checkValue != -1)
            {

                CutString(checkValue);
                this.checkForBookName = 0;
                CheckStringMatch(this.sentence, @"^[I][n]", ref checkValue);
                if (checkValue != -1)
                {
                    return true;
                }
                return ForBookNameToIn();
            }
            CheckStringMatch(this.sentence, @"^(\?\s)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                this.checkForBookName = 0;
                return true;
            }
            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToIn();
            }
            return false;
        }

        public bool ForSearch()
        {

            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, @"^[R][e][t][r][i][e][v][e][d]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
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
                                                return true;
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
                    CheckStringMatch(this.sentence, @"^([0-9a-zA-z./:-=])+", ref checkValue);
                    if (checkValue != -1)
                    {
                        CutString(checkValue);
                        return true;
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
                                return true;
                            }
                        }
                    }


                }
            }
            return false;
        }

        public bool ForBookNameToBracket()
        {
            int checkValue = -1;
            var task = Task.Factory.StartNew(() => CheckStringMatch(this.sentence, "^[‘]", ref checkValue));
            var completedWithinAllotedTime = task.Wait(TimeSpan.FromMilliseconds(1000));
            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
                if (checkValue != -1)
                {

                    CutString(checkValue);
                    CheckStringMatch(this.sentence, @"^(\s)+", ref checkValue);
                    while (checkValue != -1)
                    {
                        CutString(checkValue);
                        CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);
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
                            return ForBookNameToBracket();
                        }
                    }
                    return false;
                }
            }
            //Match match = Regex.Match(this.sentence, @"^[A-Z]([A-Za-z])+");
            CheckStringMatch(this.sentence, @"^([0-9A-Za-z-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                CheckStringMatch(this.sentence, @"^(\.){2,}", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                    return ForBookNameToBracket();
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
                return ForBookNameToBracket();
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
                return ForBookNameToBracket();
            }

            CheckStringMatch(this.sentence, @"^(\.)", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                return ForBookNameToBracket();
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
                    return true;
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
                    return true;
                }
            }
            return false;
        }

        public bool ForBookNameDB(int check)
        {
            int checkValue = -1;
            CheckStringMatch(this.sentence, @"^\(", ref checkValue);
            if (checkValue != -1)
            {
                CutString(checkValue);
                check = 1;
            }
            else if (check == 1)
            {

            }
            else
            {
                return false;
            }
            //Match match = Regex.Match(this.sentence, @"^([ก-ฮะ-์])+");
            CheckStringMatch(this.sentence, @"^([A-Za-z0-9-/])+", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);

                CheckStringMatch(this.sentence, @"^(\:)", ref checkValue);
                if (checkValue != -1)
                {
                    this.checkForBookName++;
                    CutString(checkValue);
                }


                CheckStringMatch(this.sentence, @"^(\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                CheckStringMatch(this.sentence, @"^(\.\s)", ref checkValue);
                if (checkValue != -1)
                {
                    CutString(checkValue);
                }
                return ForBookNameDB(check);
            }
            CheckStringMatch(this.sentence, @"^\)", ref checkValue);

            if (checkValue != -1)
            {
                CutString(checkValue);
                if (this.checkForBookName > 0)
                {
                    this.checkForBookName = 0;
                    return true;
                }
            }

            return false;
        }


        void CutString(int strLength)
        {
            this.countLength += strLength;
            this.sentence = this.sentence.Remove(0, strLength);
        }
    }
}
