using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using Word = Microsoft.Office.Interop.Word;
namespace SeniorProject
{
    class NumberToThai
    {
        Dictionary<string, string> dictionaryNameWithNumber;

        public NumberToThai()
        {
            dictionaryNameWithNumber = new Dictionary<string, string>();
        }

        private string StringThaiZero(string txt)
        {
            if (txt[0] == '0')
            {
                string[] num = { "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า"};
                string thai = "";

                foreach (char c in txt)
                {
                    thai += num[Int32.Parse(c + "")];
                }

                return thai;
            }
            return "";
        }

        private string StringThai(string txt)
        {
            string stringThaiZero =  StringThaiZero(txt);
            if (stringThaiZero != "")
            {
                return stringThaiZero;
            }
            string bahtTxt, n, bahtTH = "";
            double amount;
            try { amount = Convert.ToDouble(txt); }
            catch { amount = 0; }
            bahtTxt = amount.ToString("####.00");
            string[] num = { "ศูนย์", "หนึ่ง", "สอง", "สาม", "สี่", "ห้า", "หก", "เจ็ด", "แปด", "เก้า", "สิบ" };
            string[] rank = { "", "สิบ", "ร้อย", "พัน", "หมื่น", "แสน" };
            string[] temp = bahtTxt.Split('.');
            string mil = "";
            string intVal = "";
            string decVal = "";
            if (temp[0].Length > 6)
            {
                mil = temp[0].Substring(0, temp[0].Length - 6);
                intVal = temp[0].Substring(temp[0].Length - 6);
            }
            else
            {
                intVal = temp[0];
            }
            decVal = temp[1];

            if (Convert.ToDouble(bahtTxt) == 1)
            {
                bahtTH = "หนึ่ง";
            }
            else if (Convert.ToDouble(bahtTxt) == 0)
                bahtTH = "ศูนย์";
            else
            {
                for (int i = 0; i < mil.Length; i++)
                {
                    n = mil.Substring(i, 1);
                    if (n != "0")
                    {
                        if (((i == (mil.Length - 1)) && (n == "1") && (mil.Length == 1)))
                            bahtTH += "หนึ่ง";
                        else if ((i == (mil.Length - 1)) && (n == "1"))
                            bahtTH += "เอ็ด";
                        else if ((i == (mil.Length - 2)) && (n == "2"))
                            bahtTH += "ยี่";
                        else if ((i == (mil.Length - 2)) && (n == "1"))
                            bahtTH += "";
                        else
                            bahtTH += num[Convert.ToInt32(n)];
                        bahtTH += rank[(mil.Length - i) - 1];
                    }
                }
                if (bahtTH != "")
                {
                    bahtTH += "ล้าน";
                }
                for (int i = 0; i < intVal.Length; i++)
                {
                    n = intVal.Substring(i, 1);
                    if (n != "0")
                    {
                        if (((i == (intVal.Length - 1)) && (n == "1") && (intVal.Length == 1)))
                            bahtTH += "หนึ่ง";
                        else if ((i == (intVal.Length - 1)) && (n == "1"))
                            bahtTH += "เอ็ด";
                        else if ((i == (intVal.Length - 2)) && (n == "2"))
                            bahtTH += "ยี่";
                        else if ((i == (intVal.Length - 2)) && (n == "1"))
                            bahtTH += "";
                        else
                            bahtTH += num[Convert.ToInt32(n)];
                        bahtTH += rank[(intVal.Length - i) - 1];
                    }
                }
                if (bahtTH != "")
                {

                }
                if (decVal == "00")
                {
                }
                else
                {
                    if (decVal == "01")
                    {
                        bahtTH += "หนึ่ง";
                    }
                    else
                    {
                        for (int i = 0; i < decVal.Length; i++)
                        {
                            n = decVal.Substring(i, 1);
                            if (n != "0")
                            {
                                if ((i == decVal.Length - 1) && (n == "1"))
                                    bahtTH += "เอ็ด";
                                else if ((i == (decVal.Length - 2)) && (n == "2"))
                                    bahtTH += "ยี่";
                                else if ((i == (decVal.Length - 2)) && (n == "1"))
                                    bahtTH += "";
                                else
                                    bahtTH += num[Convert.ToInt32(n)];
                                bahtTH += rank[(decVal.Length - i) - 1];
                            }
                        }
                    }
                }
            }
            return bahtTH;
        }

        public void NameToNumber(ref List<string> listReferences)
        {
            for (int i = 0; i < listReferences.Count; i++)
            {
                foreach (var key in dictionaryNameWithNumber.Keys)
                {
                    try
                    {
                        string copy = listReferences[i];
                        copy = copy.Substring(0, key.Length);
                        if (copy == key)
                        {
                            listReferences[i] = listReferences[i].Substring(copy.Length);
                            listReferences[i] = dictionaryNameWithNumber[key] + listReferences[i];
                            break;
                        }
                    }
                    catch { };
                }
            }
        }

        public void NameToNumber(Word.Range range)
        {
                foreach (var key in dictionaryNameWithNumber.Keys)
                {
                    try
                    {
                        string copyRange = range.Text;
                        string copyRangeSub = copyRange.Substring(0, key.Length);
                        if (copyRangeSub == key)
                        {
                            //listReferences[i] = dictionaryNameWithNumber[key] + listReferences[i];
                            FindAndReplace(range, key, dictionaryNameWithNumber[key]);
                            break;
                        }
                    }
                    catch { };
                }
        }

        public int NumberToName(ref string references,ref int lengthCut)
        {
            string buff = "";
            Match match = Regex.Match(references, @"[ก-ฮ]");
            if (match.Success)
            {
                match = Regex.Match(references, @"^([ก-ฮะ-์]\s)*");
                if (match.Success)
                {
                    buff += match.Value;
                    string copyString = references.Substring(0, match.Length);
                    references = references.Substring(match.Length);
                    match = Regex.Match(references, @"^[0-9]+(\.[0-9]+)?");
                    if (match.Success)
                    {
                        buff += match.Value;
                        string copyNumber = references.Substring(0, match.Length);
                        string stringThai = StringThai(copyNumber);
                        lengthCut = stringThai.Length;
                        dictionaryNameWithNumber.Add(stringThai, copyString + copyNumber);
                        references = references.Substring(match.Length);
                        references = copyString + stringThai + references;

                        return buff.Length;
                    }
                }
            }
            return 0;

        }

        public bool FindAndReplace(Word.Range range, object findText, object replaceWithText)
        {
            //options
            object matchCase = false;
            object matchWholeWord = false;
            object matchWildCards = false;
            object matchSoundsLike = false;
            object matchAllWordForms = false;
            object forward = true;
            object format = false;
            object matchKashida = false;
            object matchDiacritics = false;
            object matchAlefHamza = false;
            object matchControl = false;
            object read_only = false;
            object visible = true;
            object replace = 1;
            object wrap = 1;

            //execute find and replace


            return range.Find.Execute(
                ref findText,
                ref matchCase,
                ref matchWholeWord,
                ref matchWildCards,
                ref matchSoundsLike,
                ref matchAllWordForms,
                ref forward,
                ref wrap,
                ref format,
                ref replaceWithText,
                ref replace,
                ref matchKashida,
                ref matchDiacritics,
                ref matchAlefHamza,
                ref matchControl);
        }
    }
}
