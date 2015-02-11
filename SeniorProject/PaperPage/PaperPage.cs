using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    class PaperPage
    {
        private string paperSize;

        public PaperPage(string paperSize)
        {
            this.paperSize = paperSize;
        }

        public bool cheking()
        {
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            if (pageSetup.PaperSize.ToString() == this.paperSize)
            {
                return true;
            }
            return false;

        }
        public void changing()
        {
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            if (this.paperSize == "wdPaperA4")
            {
                pageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            }
        }
    }
}
