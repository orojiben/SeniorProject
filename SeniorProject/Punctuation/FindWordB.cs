using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;


namespace SeniorProject
{
    class FindWordB
    {
       public class RangeNum
        {
            public Word.Range rang;
            public bool pos; 
        }


       public FindWordB()
        {

        }


        public bool find(   object findText, 
                            object replaceWithText,
                            object matchCase,
                            object matchWholeWord,
                            object matchWildCards,
                            object matchSoundsLike,
                            object matchAllWordForms,
                            object format,
                            object matchKashida,
                            object matchDiacritics,
                            object matchAlefHamza,
                            object matchControl,
                            object forward, 
                            object replace, // 0-2
                            object wrap,    // 0-2
                                   Word.Find f){
             f.ClearFormatting();
            
            f.Execute(  ref findText,
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
            return f.Found;
        }
         public bool findR(   object findText, 
                            object replaceWithText,
                            object matchCase,
                            object matchWholeWord,
                            object matchWildCards,
                            object matchSoundsLike,
                            object matchAllWordForms,
                            object format,
                            object matchKashida,
                            object matchDiacritics,
                            object matchAlefHamza,
                            object matchControl,
                            object forward, 
                            object replace, // 0-2
                            object wrap,    // 0-2
                                   Word.Range r){

            r.Find.ClearFormatting();
            //r.Find.Replacement.ClearFormatting();
       
            r.Find.Execute(  ref findText,
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
            return r.Find.Found;
        }
               

        /* 1.findText       wdStyleHyperlink
           2.replaceWithText  
           3.matchCase  
           4.matchWholeWord  
           5.matchWildCards  
           6.matchSoundsLike  
           7.matchAllWordForms  // Object 
           8.format              // Object 
           9.matchKashida  
           10.matchDiacritics  
           11.matchAlefHamza  
           12.matchControl  
           13.forward  
           14.replace  //wdReplaceNone wdReplaceOne wdReplaceAll
           15.wrap  //wdFindStop  wdFindContinue wdFindAsk 
                                   1           2       3  4  5  6  7  8  9 10 11 12 13 14 15            */
        //    return this.find(findText, Type.Missing, 1, 1, 1, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0 



        public bool find21(object findText, Word.Find f){
            return this.find(findText, Type.Missing, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, f);
        }
        public bool find31(object findText, Word.Find f)//not wildcard
        {
           
            return this.find(findText, Type.Missing, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, f);
        }


        public bool find2(object findText, Word.Range r){
            Word.Range rang = null;
            rang = makeRange(r);
            return this.findR(findText, Type.Missing, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, rang);
        }

        public bool find3(object findText, Word.Range r)//not wildcard
        {
           // r.Select();
            Word.Range rang = null;
            rang = makeRange(r);
            
            return this.findR(findText, Type.Missing, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, rang);
        }

                  
        public void findReplace(object findText,object ReplaceText,Word.Find f){
 //                      1           2       3  4  5  6  7  8  9 10 11 12 13 14 15  
            this.find(findText, ReplaceText, 1, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, f);
        }

        public void findReplace2(object findText,object ReplaceText,Word.Find f){
 //                      1           2       3  4  5  6  7  8  9 10 11 12 13 14 15  
            this.find(findText, ReplaceText, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, f);
        }
        public void findReplace22(object findText,object ReplaceText,Word.Range f){
 //                      1           2       3  4  5  6  7  8  9 10 11 12 13 14 15  
            this.findR(findText, ReplaceText, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 1, f);
        }

        public int FindCount(object findText,Word.Find f)
        {
            int count = 0;
            while (find21(findText, f) == true)++count;
            return count;
        }



        public Word.Range makeRange(Word.Range range)
        {

            Word.Range r = Globals.ThisAddIn.Application.ActiveDocument.Range(range.Start, range.End);
            //r.Select();
            return r;
        }




        public List<Word.Range> findRange(object findText, Word.Range rangeAll){ // use find 2
            List<Word.Range> lRange = new List<Word.Range>();

            while (this.find21(findText, rangeAll.Find) == true){
                lRange.Add(this.makeRange(rangeAll));
            }
            return lRange;
        }

        public List<Word.Range> findRange2(object findText, Word.Range rangeAll){// use find 3
            List<Word.Range> lRange = new List<Word.Range>();
            while (this.find31(findText, rangeAll.Find) == true)
            {
                lRange.Add(this.makeRange(rangeAll));
            }
            return lRange;
        }




       public List<Word.Range> ReSetRange(List<Word.Range> rang,String sign){
           List<Word.Range> lRange = new List<Word.Range>();
           foreach(Word.Range r in rang){
               if(find31((object)sign,r.Find))lRange.Add(r);
               else {}
           }
          return lRange;
       }

    }
}
