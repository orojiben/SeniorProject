using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using System.Windows.Forms;
using System.IO;

namespace SeniorProject
{
    public partial class ThisAddIn
    {
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }

        public List<string> GetText()
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


        public void GetFont(string font)
        {
            List<string> text = this.GetText();
            int begin = 0, comBegin = 0, comLast = 0;
            string lastFont = "", tmp = "";
            Word.Range rng;
            foreach (string i in text)
            {
                try
                {
                    rng = this.Application.ActiveDocument.Range(begin, begin + i.Length);

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
                        rng = this.Application.ActiveDocument.Range(comBegin, comLast);
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
            rng = this.Application.ActiveDocument.Range(comBegin, comLast);
            if (lastFont != font)
            {
                rng.Comments.Add(rng, lastFont);
            }
        }

        //Correct
        public void CorrectFont(string font)
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
            this.saveNewFile();
        }


        //Save File
        public void saveNewFile()
        {
            Word.Document app = Globals.ThisAddIn.Application.ActiveDocument;
            SaveFileDialog save = new SaveFileDialog();
            save.Filter = "Word Documents|*.docx";
            if (save.ShowDialog() == DialogResult.OK)
            {
                string path = Path.GetFullPath(save.FileName);
                app.SaveAs2(path);
            }
        }


        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }

        #endregion
    }
}
