using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Threading;

namespace SeniorProject
{
    public class MarginPage
    {
        public bool leftCheck;
        public bool rightCheck;
        public bool topCheck;
        public bool bottomCheck;
        private Word.Document documentError;
        public MarginPage()
        {
        }

        /*public float inchToPoint(float inch)
        {
            return 72f * inch;
        }

        private float centimeterToPoint(float centimeter)
        {
            return 28.34645669291f * centimeter;
        }*/

        public bool cheking()
        {
            chekingNotShow();
            show();
            return this.cheked();
        }

        public bool chekingNotShow()
        {
            //Ribbon1.marginPageUC = new MarginPageUC();
            if (!Ribbon1.marginPageUC.checkSetClick)
            {
                Ribbon1.marginPageUC.btn_Edit.Click += new System.EventHandler(this.btn_Edit_Click);
                Ribbon1.marginPageUC.button1.Click += new System.EventHandler(this.btn_Edit_Click);
                Ribbon1.marginPageUC.checkSetClick = true;
            }
            //NextPange();
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;

            this.leftCheck = pageSetup.LeftMargin == Ribbon1.styles.LeftMargin;
            this.rightCheck = pageSetup.RightMargin == Ribbon1.styles.RightMargin;
            this.topCheck = pageSetup.TopMargin == Ribbon1.styles.TopMargin;
            this.bottomCheck = pageSetup.BottomMargin == Ribbon1.styles.BottomMargin;
            Ribbon1.marginPageUC.setMarginPageUC(this.cheked(), this.leftCheck, this.rightCheck,
                 this.topCheck, this.bottomCheck
                , pageSetup.LeftMargin, pageSetup.RightMargin, pageSetup.TopMargin, pageSetup.BottomMargin);
            return this.cheked();
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

        public bool cheked()
        {
            return this.leftCheck && this.rightCheck && this.topCheck && this.bottomCheck;
        }
        
        public void changing()
        {
            //NextPangeChenging();
            try
           {
               Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
                pageSetup.LeftMargin = Ribbon1.styles.LeftMargin;
                pageSetup.RightMargin = Ribbon1.styles.RightMargin;
                pageSetup.TopMargin = Ribbon1.styles.TopMargin;
                pageSetup.BottomMargin = Ribbon1.styles.BottomMargin;
            }
            catch {
          //      System.Windows.Forms.MessageBox.Show("ไม่สามารตแก้ไขได้");
                Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
                this.documentError = new Word.Document();
                this.documentError.PageSetup.PaperSize = pageSetup.PaperSize;
                this.documentError.PageSetup.LeftMargin = Ribbon1.styles.LeftMargin;
                this.documentError.PageSetup.RightMargin = Ribbon1.styles.RightMargin;
                this.documentError.PageSetup.TopMargin = Ribbon1.styles.TopMargin;
                this.documentError.PageSetup.BottomMargin = Ribbon1.styles.BottomMargin;
                
                pageSetup = this.documentError.PageSetup;
                object missing = System.Reflection.Missing.Value;
                object fileName = "bin.docx";
                this.documentError.SaveAs(ref fileName,
        ref missing, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref missing,
        ref missing, ref missing, ref missing, ref missing, ref missing);

                Thread threadDeleteFile = new Thread(this.deleteFileClose);
                threadDeleteFile.Start();

            };
        }

        public void deleteFileClose()
        {
            Thread.Sleep(1000);
            this.documentError.Close();
            deleteFile();
        }

        public void deleteFile()
        {
              try
           {
               string pathDoc = Environment.GetFolderPath(Environment.SpecialFolder.Personal);
               string path = pathDoc + "\\" + "bin.docx";
               System.IO.File.Delete(path);
                
           }
           catch 
           {
               deleteFile();
           }
        }

        public void show2(){
            Word.PageSetup pageSetup = Globals.ThisAddIn.Application.ActiveDocument.PageSetup;
            pageSetup.Orientation = Microsoft.Office.
    Interop.Word.WdOrientation.wdOrientLandscape;

            pageSetup.Orientation = Microsoft.Office.
    Interop.Word.WdOrientation.wdOrientPortrait;
            pageSetup.LeftMargin = Ribbon1.styles.LeftMargin;
            pageSetup.RightMargin = Ribbon1.styles.RightMargin;
            pageSetup.TopMargin = Ribbon1.styles.TopMargin;
            pageSetup.BottomMargin = Ribbon1.styles.BottomMargin;
        }


        public void show()
        {
            Ribbon1.showCustomTaskPane(1);
            /*if (ThisAddIn.mainCustomTaskPane.Count > 0)
            {
                int count = ThisAddIn.mainCustomTaskPane.Count;
                for (int i = 0; i < count; ++i)
                {
                    ThisAddIn.mainCustomTaskPane.RemoveAt(0);
                }
            }

            Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.marginPageUC, "Margin Page");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/
        }

        public void showForAll()
        {
            Ribbon1.showCustomTaskPane(1,true);
            /*if (ThisAddIn.mainCustomTaskPane.Count > 1)
            {
                int count = ThisAddIn.mainCustomTaskPane.Count;
                for (int i = 1; i < count; ++i)
                {
                    ThisAddIn.mainCustomTaskPane.RemoveAt(1);
                }
            }
            this.chekingNotShow();
            /*Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.showCheckAllUC, "Check All");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/
            /*Ribbon1.myCustomTaskPane = ThisAddIn.mainCustomTaskPane.Add(Ribbon1.marginPageUC, "Margin Page");
            Ribbon1.myCustomTaskPane.DockPosition = Microsoft.Office.Core.MsoCTPDockPosition.msoCTPDockPositionRight;
            Ribbon1.myCustomTaskPane.Width = 300;
            Ribbon1.myCustomTaskPane.Visible = true;*/

        }


        private void NextPange()
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;
            object missing = System.Reflection.Missing.Value;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int numberPang = document.ComputeStatistics(stat, missing);
            for (int i = 1; i <= numberPang; i++)
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";

                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;
                Word.PageSetup pageSetup = range.Document.PageSetup;
                    this.leftCheck = pageSetup.LeftMargin == Ribbon1.styles.LeftMargin;
                    this.rightCheck = pageSetup.RightMargin == Ribbon1.styles.RightMargin;
                        this.topCheck = pageSetup.TopMargin == Ribbon1.styles.TopMargin;
                    this.bottomCheck = pageSetup.BottomMargin == Ribbon1.styles.BottomMargin;
                    Ribbon1.marginPageUC.setMarginPageUC(this.cheked(), this.leftCheck, this.rightCheck,
         this.topCheck, this.bottomCheck
        , pageSetup.LeftMargin, pageSetup.RightMargin, pageSetup.TopMargin, pageSetup.BottomMargin);
            }
        }
        private void NextPangeChenging()
        {
            Word.Application wordApp = Globals.ThisAddIn.Application;
            object missing = System.Reflection.Missing.Value;
            Word.Document document = Globals.ThisAddIn.Application.ActiveDocument;
            Word.WdStatistic stat = Word.WdStatistic.wdStatisticPages;
            int numberPang = document.ComputeStatistics(stat, missing);
            for (int i = 1; i <= numberPang; i++)
            {
                object what = Word.WdGoToItem.wdGoToPage;
                object which = null;
                object counts = i;
                object name = null;
                object Page = "\\Page";

                wordApp.Selection.GoTo(ref what, ref which, ref counts, ref name);
                Word.Range range = wordApp.ActiveDocument.Bookmarks.get_Item(ref Page).Range;
                Word.PageSetup pageSetup = range.Document.PageSetup;
                pageSetup.LeftMargin = Ribbon1.styles.LeftMargin;
                System.Windows.Forms.MessageBox.Show("ไม่ตรง2" + Ribbon1.styles.LeftMargin);
                pageSetup.RightMargin = Ribbon1.styles.RightMargin;
                pageSetup.TopMargin = Ribbon1.styles.TopMargin;
                pageSetup.BottomMargin = Ribbon1.styles.BottomMargin;
            }
        }


    }
}
