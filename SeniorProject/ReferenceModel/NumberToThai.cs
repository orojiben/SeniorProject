using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    class NumberToThai
    {

        public static string StringThai(string txt)
        {
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
    }
}
