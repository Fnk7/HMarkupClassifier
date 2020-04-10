using ClosedXML.Excel;
using System;

namespace HMarkupClassifier.SheetParser
{
    class XContentFactory
    {
        public XContent GetXContent(IXLCell cell)
        {
            if (cell.IsEmpty())
                return new XContent();
            try
            {
                string content = cell.GetString();
                return new XContent(content);
            }
            catch (Exception ex)
            {
                throw new ParseFailException("Parse Content Fail!", ex);
            }
        }
    }
}
