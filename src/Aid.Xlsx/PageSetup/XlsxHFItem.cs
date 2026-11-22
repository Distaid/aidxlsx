#nullable disable

using System.Text;
using Aid.Xlsx.Extensions;
using Aid.Xlsx.RichText;

namespace Aid.Xlsx.PageSetup
{
    internal class XlsxHFItem : IXlsxHFItem
    {
        internal readonly XlsxHeaderFooter HeaderFooter;
        public XlsxHFItem(XlsxHeaderFooter headerFooter)
        {
            HeaderFooter = headerFooter;
        }
        public XlsxHFItem(XlsxHFItem defaultHFItem, XlsxHeaderFooter headerFooter)
            : this(headerFooter)
        {
            defaultHFItem.texts.ForEach(kp => texts.Add(kp.Key, kp.Value));
        }
        private readonly Dictionary<XlsxHFOccurrence, List<XlsxHFText>> texts = new Dictionary<XlsxHFOccurrence, List<XlsxHFText>>();
        public String GetText(XlsxHFOccurrence occurrence)
        {
            var sb = new StringBuilder();
            if (texts.TryGetValue(occurrence, out List<XlsxHFText> hfTexts))
            {
                foreach (var hfText in hfTexts)
                    sb.Append(hfText.GetHFText(sb.ToString()));
            }

            return sb.ToString();
        }

        public IXlsxRichString AddText(String text)
        {
            return AddText(text, XlsxHFOccurrence.AllPages);
        }
        public IXlsxRichString AddText(XlsxHFPredefinedText predefinedText)
        {
            return AddText(predefinedText, XlsxHFOccurrence.AllPages);
        }

        public IXlsxRichString AddText(String text, XlsxHFOccurrence occurrence)
        {
            // TODO Styles: This doesn't update source when API object changes
            var richText = new XlsxRichString(text, HeaderFooter.Worksheet.GetFormat().Font, this, HeaderFooter.Worksheet.Workbook.Styles, null);

            var hfText = new XlsxHFText(richText, this);
            if (occurrence == XlsxHFOccurrence.AllPages)
            {
                AddTextToOccurrence(hfText, XlsxHFOccurrence.EvenPages);
                AddTextToOccurrence(hfText, XlsxHFOccurrence.FirstPage);
                AddTextToOccurrence(hfText, XlsxHFOccurrence.OddPages);
            }
            else
            {
                AddTextToOccurrence(hfText, occurrence);
            }

            return richText;
        }

        public IXlsxRichString AddNewLine()
        {
            return AddText(Environment.NewLine);
        }

        public IXlsxRichString AddImage(String imagePath, XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages)
        {
            throw new NotImplementedException();
        }

        private void AddTextToOccurrence(XlsxHFText hfText, XlsxHFOccurrence occurrence)
        {
            if (texts.TryGetValue(occurrence, out List<XlsxHFText> hfTexts))
                hfTexts.Add(hfText);
            else
                texts.Add(occurrence, new List<XlsxHFText> { hfText });

            this.HeaderFooter.Changed = true;
        }

        public IXlsxRichString AddText(XlsxHFPredefinedText predefinedText, XlsxHFOccurrence occurrence)
        {
            String hfText;
            switch (predefinedText)
            {
                case XlsxHFPredefinedText.PageNumber: hfText = "&P"; break;
                case XlsxHFPredefinedText.NumberOfPages: hfText = "&N"; break;
                case XlsxHFPredefinedText.Date: hfText = "&D"; break;
                case XlsxHFPredefinedText.Time: hfText = "&T"; break;
                case XlsxHFPredefinedText.Path: hfText = "&Z"; break;
                case XlsxHFPredefinedText.File: hfText = "&F"; break;
                case XlsxHFPredefinedText.SheetName: hfText = "&A"; break;
                case XlsxHFPredefinedText.FullPath: hfText = "&Z&F"; break;
                default: throw new NotImplementedException();
            }
            return AddText(hfText, occurrence);
        }

        public void Clear(XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages)
        {
            if (occurrence == XlsxHFOccurrence.AllPages)
            {
                ClearOccurrence(XlsxHFOccurrence.EvenPages);
                ClearOccurrence(XlsxHFOccurrence.FirstPage);
                ClearOccurrence(XlsxHFOccurrence.OddPages);
            }
            else
            {
                ClearOccurrence(occurrence);
            }
        }

        private void ClearOccurrence(XlsxHFOccurrence occurrence)
        {
            if (texts.ContainsKey(occurrence))
                texts.Remove(occurrence);
        }
    }
}
