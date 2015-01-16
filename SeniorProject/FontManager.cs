using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class FontManager
    {
        public static List<string> GetText()
        {
            Word.Document app = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range range = app.Content;
            string xml = range.get_XML(false);
            List<string> str = new List<string>();
            int begin = 0, last = 0;
            while (xml.IndexOf("<w:t>") != -1)
            {
                begin = xml.IndexOf("<w:t>");
                begin += 5;
                last = xml.IndexOf("</w:t>");
                str.Add(xml.Substring(begin, last - begin));
                xml = xml.Substring(last + 6);
            }
            return str;
        }

        public static void CheckFontSize(int small, int medium, int large)
        {
            

        }

        public static void CheckFontName(string font)
        {
            try
            {
                Globals.ThisAddIn.Application.ActiveDocument.DeleteAllComments();
            }
            catch (Exception e)
            {

            }
            List<string> text = GetText();
            int begin = 0, comBegin = 0, comLast = 0;
            string lastFont = "", tmp = "";
            Word.Range rng;
            foreach (string i in text)
            {
                try
                {
                    rng = Globals.ThisAddIn.Application.ActiveDocument.Range(begin, begin + i.Length);

                    //rng.Select();
                    if (lastFont == "")
                    {
                        lastFont = rng.Font.Name;
                        comLast += i.Length;

                    }
                    else if (lastFont == rng.Font.Name)
                    {
                        comLast += i.Length;
                        //begin = comLast + 1;
                    }
                    else
                    {
                        tmp = lastFont;
                        lastFont = rng.Font.Name;
                        rng = Globals.ThisAddIn.Application.ActiveDocument.Range(comBegin, comLast);
                        rng.Select();
                        if (tmp != font)
                        {
                            rng.Comments.Add(rng, tmp);
                        }
                        comBegin = comLast;
                        comLast = comBegin + i.Length + 1;
                    }
                    begin = comLast;
                }
                catch { }
            }
            rng = Globals.ThisAddIn.Application.ActiveDocument.Range(comBegin, comLast);
            if (lastFont != font)
            {
                rng.Comments.Add(rng, lastFont);
            }
        }

        //Correct
        public static void CorrectFont(string font)
        {
            Word.Document app = Globals.ThisAddIn.Application.ActiveDocument;
            Word.Range range = app.Content;
            range.Font.Name = font;
            try
            {
                app.DeleteAllComments();
            }
            catch (Exception e)
            {

            }
        }
    }
}
