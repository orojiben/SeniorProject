using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace SeniorProject
{
    public class PaperPage
    {     
        public PaperPage()
        {
        }

        public bool cheking()
        {
            bool value = chekingNotShow();
            show();
            return value;

        }

        public bool chekingNotShow()
        {
            //Ribbon1.paperPageUC = new PaperPageUC();
            if (!Ribbon1.paperPageUC.checkSetClick)
            {
                Ribbon1.paperPageUC.btn_Edit.Click += new System.EventHandler(this.btn_Edit_Click);
                Ribbon1.paperPageUC.checkSetClick = true;
            }
            Word.PageSetup pageSetup =  Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            
            Ribbon1.paperPageUC.setPaperPageUC(pageSetup.PaperSize.ToString() == Ribbon1.styles.Paper, pageSetup.PaperSize.ToString());
            return pageSetup.PaperSize.ToString() == Ribbon1.styles.Paper;

        }

        public void changing()
        {
            Word.PageSetup pageSetup =  Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            if (Ribbon1.styles.Paper == "wdPaperA4")
            {
                pageSetup.PaperSize = Word.WdPaperSize.wdPaperA4;
            }
        }

        private void btn_Edit_Click(object sender, EventArgs e)
        {
            Ribbon1.saveFileAuto();
            this.changing();
            if (Globals.ThisAddIn.CustomTaskPanes[0].Visible == true)
            {

                this.chekingNotShow();
                this.showForAll();
                Ribbon1.showCheckAllUC.setButtonClickALL();
            }
            else
            {
                this.cheking();
            }
        }

        public void show()
        {
            Ribbon1.showCustomTaskPane(2);
            /*if (ThisAddIn.mainCustomTaskPane.Count > 0)
            {
                for (int i = 0; i < ThisAddIn.mainCustomTaskPane.Count; ++i)
                {
                    ThisAddIn.mainCustomTaskPane.RemoveAt(i);
                }
            }

            Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.paperPageUC, "Paper Page");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/
        }

        public void showForAll()
        {
            Ribbon1.showCustomTaskPane(2, true);
            /*if (ThisAddIn.mainCustomTaskPane.Count > 1)
            {
                for (int i = 1; i < ThisAddIn.mainCustomTaskPane.Count; ++i)
                {
                    ThisAddIn.mainCustomTaskPane.RemoveAt(i);
                }
            }
            this.chekingNotShow();
            Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.paperPageUC, "Paper Page");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/
        }
    }
}
