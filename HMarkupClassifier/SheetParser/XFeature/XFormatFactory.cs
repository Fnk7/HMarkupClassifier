using ClosedXML.Excel;
using System.Collections.Generic;
using System.Linq;

namespace HMarkupClassifier.SheetParser
{
    class XFormatFactory
    {
        public int MaxStyles = 100;
        public Dictionary<XStyle, XStyle> styles = new Dictionary<XStyle, XStyle>();
        public List<string> FontNames = new List<string>();

        public XFormat GetXFormat(IXLCell cell)
        {
            XFormat xFormat = new XFormat(cell);
            xFormat.Style = GetXStyle(cell.Style);
            return xFormat;
        }

        public XStyle GetXStyle(IXLStyle style)
        {
            XStyle xStyle = new XStyle(style);
            var FontName = style.Font.FontName;
            if (!FontNames.Contains(FontName))
                FontNames.Add(FontName);
            xStyle.Font.NameIndex = FontNames.IndexOf(FontName);
            if (styles.ContainsKey(xStyle))
                xStyle = styles[xStyle];
            else
            {
                styles.Add(xStyle, xStyle);
                if (styles.Count >= MaxStyles)
                {
                    throw new ParseFailException($"Styles Count over {MaxStyles}");
                }
            }
            xStyle.count++;
            return xStyle;
        }

        public void SetOrderedIndex()
        {
            Dictionary<int, int> FontNameCnt = new Dictionary<int, int>();
            Dictionary<int, int> FontColorCnt = new Dictionary<int, int>();
            Dictionary<int, int> PtnColorCnt = new Dictionary<int, int>();
            Dictionary<int, int> BckColorCnt = new Dictionary<int, int>();
            foreach (var style in styles.Keys)
            {
                if (!FontNameCnt.ContainsKey(style.Font.NameIndex))
                    FontNameCnt.Add(style.Font.NameIndex, 0);
                if (!FontColorCnt.ContainsKey(style.Font.Color))
                    FontColorCnt.Add(style.Font.Color, 0);
                if (!PtnColorCnt.ContainsKey(style.Fill.PtnColor))
                    PtnColorCnt.Add(style.Fill.PtnColor, 0);
                if (!BckColorCnt.ContainsKey(style.Fill.BckColor))
                    BckColorCnt.Add(style.Fill.BckColor, 0);
                FontNameCnt[style.Font.NameIndex] += style.count;
                FontColorCnt[style.Font.Color] += style.count;
                PtnColorCnt[style.Fill.PtnColor] += style.count;
                BckColorCnt[style.Fill.BckColor] += style.count;
            }
            var OrderedFontName = (from x in FontNameCnt orderby x.Value descending select x.Key).ToList();
            var OrderedFontColor = (from x in FontColorCnt orderby x.Value descending select x.Key).ToList();
            var OrderedPtnColor = (from x in PtnColorCnt orderby x.Value descending select x.Key).ToList();
            var OrderedBckColor = (from x in BckColorCnt orderby x.Value descending select x.Key).ToList();
            foreach (var style in styles.Keys)
            {
                style.Font.NameIndex = OrderedFontName.IndexOf(style.Font.NameIndex);
                style.Font.Color = OrderedFontColor.IndexOf(style.Font.Color);
                style.Fill.PtnColor = OrderedPtnColor.IndexOf(style.Fill.PtnColor);
                style.Fill.BckColor = OrderedBckColor.IndexOf(style.Fill.BckColor);
            }
        }
    }
}
