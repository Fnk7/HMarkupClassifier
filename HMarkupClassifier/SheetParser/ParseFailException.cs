using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace HMarkupClassifier.SheetParser
{
    class ParseFailException : Exception
    {
        public ParseFailException(string message) : base(message)
        {
        }

        public ParseFailException(string message, Exception innerException) : base(message, innerException)
        {
        }
    }
}
