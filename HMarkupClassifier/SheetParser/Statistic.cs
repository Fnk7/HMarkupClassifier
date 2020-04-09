using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

using HMarkupClassifier.SheetParser.Styles;

namespace HMarkupClassifier.SheetParser
{
    class Statistic
    {
        Dictionary<Style, Style> styles = new Dictionary<Style, Style>();

        public Style GetStyle(IXLStyle style)
        {
            Style temp = new Style(style);
            if (styles.ContainsKey(temp))
                temp = styles[temp];
            else styles.Add(temp, temp);
            temp.count++;
            return temp;
        }


    }
}
