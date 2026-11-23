#nullable disable

using Aid.Xlsx.Infrastructure.Extensions;
using Aid.Xlsx.Model.Worksheets;

namespace Aid.Xlsx.Model.PageSetup;

using System.Linq;

internal class XlsxHeaderFooter : IXlsxHeaderFooter
{
    public XlsxHeaderFooter(XlsxWorksheet worksheet)
    {
        this.Worksheet = worksheet;
        Left = new XlsxHFItem(this);
        Right = new XlsxHFItem(this);
        Center = new XlsxHFItem(this);
        SetAsInitial();
    }

    public XlsxHeaderFooter(XlsxHeaderFooter defaultHF, XlsxWorksheet worksheet)
    {
        this.Worksheet = worksheet;
        defaultHF.innerTexts.ForEach(kp => innerTexts.Add(kp.Key, kp.Value));
        Left = new XlsxHFItem((XlsxHFItem)defaultHF.Left, this);
        Center = new XlsxHFItem((XlsxHFItem)defaultHF.Center, this);
        Right = new XlsxHFItem((XlsxHFItem)defaultHF.Right, this);
        SetAsInitial();
    }

    internal readonly XlsxWorksheet Worksheet;

    public IXlsxHFItem Left { get; private set; }
    public IXlsxHFItem Center { get; private set; }
    public IXlsxHFItem Right { get; private set; }

    public String GetText(XlsxHFOccurrence occurrence)
    {
        //if (innerTexts.ContainsKey(occurrence)) return innerTexts[occurrence];

        var retVal = String.Empty;
        var leftText = Left.GetText(occurrence);
        var centerText = Center.GetText(occurrence);
        var rightText = Right.GetText(occurrence);
        retVal += leftText.Length > 0 ? "&L" + leftText : String.Empty;
        retVal += centerText.Length > 0 ? "&C" + centerText : String.Empty;
        retVal += rightText.Length > 0 ? "&R" + rightText : String.Empty;
        if (retVal.Length > 255)
            throw new ArgumentOutOfRangeException(null, "Headers and footers cannot be longer than 255 characters (including style markup).");
        return retVal;
    }

    private readonly Dictionary<XlsxHFOccurrence, String> innerTexts = [];
    internal void SetInnerText(XlsxHFOccurrence occurrence, String text)
    {
        var parsedElements = ParseFormattedHeaderFooterText(text);

        if (parsedElements.Any(e => e.Position == 'L'))
            this.Left.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'L').Select(e => e.Text).ToArray()), occurrence);

        if (parsedElements.Any(e => e.Position == 'C'))
            this.Center.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'C').Select(e => e.Text).ToArray()), occurrence);

        if (parsedElements.Any(e => e.Position == 'R'))
            this.Right.AddText(string.Join("\r\n", parsedElements.Where(e => e.Position == 'R').Select(e => e.Text).ToArray()), occurrence);

        innerTexts[occurrence] = text;
    }

    private struct ParsedHeaderFooterElement
    {
        public char Position;
        public string Text;
    }

    private static List<ParsedHeaderFooterElement> ParseFormattedHeaderFooterText(string text)
    {
        bool IsAtPositionIndicator(int i) => i < text.Length - 1 && text[i] == '&' && (sourceArray.Contains(text[i + 1]));

        var parsedElements = new List<ParsedHeaderFooterElement>();
        var currentPosition = 'L'; // default is LEFT
        var hfElement = "";

        for (int i = 0; i < text.Length; i++)
        {
            if (IsAtPositionIndicator(i))
            {
                if (hfElement.Length > 0) parsedElements.Add(new ParsedHeaderFooterElement()
                {
                    Position = currentPosition,
                    Text = hfElement
                });

                currentPosition = text[i + 1];
                i += 2;
                hfElement = "";
            }

            if (i < text.Length)
            {
                if (IsAtPositionIndicator(i))
                    i--;
                else
                    hfElement += text[i];
            }
        }

        if (hfElement.Length > 0)
            parsedElements.Add(new ParsedHeaderFooterElement()
            {
                Position = currentPosition,
                Text = hfElement
            });
        return parsedElements;
    }

    private Dictionary<XlsxHFOccurrence, String> _initialTexts;

    private Boolean _changed;
    internal Boolean Changed
    {
        get
        {
            return _changed || _initialTexts.Any(it => GetText(it.Key) != it.Value);
        }
        set { _changed = value; }
    }

    internal static readonly char[] sourceArray = ['L', 'C', 'R'];

    internal void SetAsInitial()
    {
        _initialTexts = [];
        foreach (var o in Enum.GetValues<XlsxHFOccurrence>())
        {
            _initialTexts.Add(o, GetText(o));
        }
    }


    public IXlsxHeaderFooter Clear(XlsxHFOccurrence occurrence = XlsxHFOccurrence.AllPages)
    {
        Left.Clear(occurrence);
        Right.Clear(occurrence);
        Center.Clear(occurrence);
        return this;
    }
}
