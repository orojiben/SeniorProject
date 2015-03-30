using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace SeniorProject
{
    public class Punctuation
    {   
        public Punctuation()
        {
        }

        public void show()
        {
            //Ribbon1.punctuationUC.checkAll();
            //Ribbon1.showCustomTaskPane(6);
            //Thread newThread = new Thread(Ribbon1.punctuationUC.checkAll);
            //newThread.Start();
            Ribbon1.punctuationUC.enableAll();
            Ribbon1.showCustomTaskPane(6);
           // Ribbon1.punctuationUC.checkAll();
            //Ribbon1.showCustomTaskPane(6);
            
        }

        public void showForAll()
        {
            Ribbon1.showCustomTaskPane(6, true);
        }
    }
}
