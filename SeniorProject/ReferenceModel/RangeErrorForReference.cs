using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Word = Microsoft.Office.Interop.Word;

namespace SeniorProject
{
    class RangeErrorForReference
    {
        private int numberReferenceError;
        private Word.Range rangeReferenceError;
        public bool rangeErrorModel;
        public bool rangeErrorNameYear;
        public bool rangeErrorMargin;
        public RangeErrorForReference(int numberReferenceError, Word.Range rangeReferenceError, int typeError)
        {
            this.rangeErrorModel = false;
            this.rangeErrorNameYear = false;
            this.rangeErrorMargin = false;
            this.numberReferenceError = numberReferenceError;
            this.rangeReferenceError = rangeReferenceError;
            this.setRangeError(typeError);
        }

        public int getNumberReferenceError()
        {
            return this.numberReferenceError;
        }

        public Word.Range getRangeReferenceError()
        {
            return this.rangeReferenceError;
        }

        public void setRangeError(int typeError)
        {
            if (typeError == 1)
            {
                this.rangeErrorModel = true;
            }
            else if (typeError == 2)
            {
                this.rangeErrorNameYear = true;
            }
            else if (typeError == 3)
            {
                this.rangeErrorMargin = true;
            }
        }

        public bool checkBtnEditEnabled()
        {
            return this.rangeErrorNameYear || this.rangeErrorMargin;
        }
    }
}
