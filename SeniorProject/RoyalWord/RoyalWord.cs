using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SeniorProject
{
    public class RoyalWord
    {

        public void show()
        {
            //Ribbon1.royalWordUC.check_incorrect_word();
            Ribbon1.royalWordUC.enableAll();
            Ribbon1.showCustomTaskPane(5);
        }

        public void showForAll()
        {
            Ribbon1.showCustomTaskPane(5, true);
        }
    }
}
