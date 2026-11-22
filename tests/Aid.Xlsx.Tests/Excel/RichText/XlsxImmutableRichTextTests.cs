using Aid.Xlsx.RichText;
using Aid.Xlsx.Workbooks;
using NUnit.Framework;

namespace Aid.Xlsx.Tests.Excel.RichText
{
    [TestFixture]
    public class XlsxImmutableRichTextTests
    {
        [Test]
        public void Equals_compares_text_runs_phonetic_runs_and_properties()
        {
            using var wb = new XlsxWorkbook();
            var ws = wb.AddWorksheet();
            var richText = (XlsxRichText)ws.Cell("A1").CreateRichText();
            richText
                .AddText("こんにち").SetBold(true) // Hello in hiragana
                .AddText("は,").SetBold(false) // object marker
                .AddText("世界").SetFontSize(15); // world in kanji
            richText.Phonetics
                .SetAlignment(XlsxPhoneticAlignment.Distributed)
                .Add(@"konnichi wa", 0, 6); // world in hiragana

            // Assert equal
            var immutableRichText = XlsxImmutableRichText.Create(richText);
            var equalImmutableRichText = XlsxImmutableRichText.Create(richText);
            Assert.AreEqual(immutableRichText, equalImmutableRichText);

            // Different font of a first run
            richText.ElementAt(0).SetBold(false);
            var withDifferentTextRunFont = XlsxImmutableRichText.Create(richText);
            Assert.AreNotEqual(immutableRichText, withDifferentTextRunFont);
            richText.ElementAt(0).SetBold(true);

            // Different phonetic properties
            richText.Phonetics.SetAlignment(XlsxPhoneticAlignment.Left);
            var withDifferentPhoneticsProps = XlsxImmutableRichText.Create(richText);
            Assert.AreNotEqual(immutableRichText, withDifferentPhoneticsProps);
            richText.Phonetics.SetAlignment(XlsxPhoneticAlignment.Distributed);

            // Different phonetic runs
            richText.Phonetics.Add("せかい", 6, 8);
            var withDifferentTextPhonetics = XlsxImmutableRichText.Create(richText);
            Assert.AreNotEqual(immutableRichText, withDifferentTextPhonetics);
        }
    }
}
