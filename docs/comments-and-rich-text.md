# Comments and rich text

A comment is the legacy Excel note attached to a cell: a small popup with text, an
author, and a shape. Rich text gives individual runs inside one cell their own font,
size, colour, and emphasis. Both are part of the cell model, so they survive a load
and save unless you replace the content that holds them.

## Comments

Use a comment to leave an explanation on a cell without changing its value. Each cell
can hold exactly one comment. The API creates the legacy notes that Excel shows on
hover or pins in the sheet; it does not model the newer threaded comments.

```csharp
using System;
using Aid.Xlsx.Model.Comments;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Budget");

var cell = sheet.Cell("B2");
cell.Value = 1250;
IXlsxComment note = cell.CreateComment();
note.Author = "Finance";
note.AddSignature();
note.AddText("Draft figure. ");
note.AddText("Pending approval.").SetItalic().SetFontColor(XlsxColor.DarkRed);
note.Visible = true;
note.Style.Size.Width = 30;
note.Style.Size.Height = 110;
note.Position.Row = 1;
note.Position.Column = 4;

workbook.SaveAs("budget.xlsx");
```

### Author and signature

A new comment starts with the author of its worksheet, which is initialised from
`XlsxWorkbook.Author` (`Environment.UserName` for a new workbook). `SetAuthor` changes
the name on that comment only and returns the comment, so it can be chained.
`AddSignature` appends the current author with a colon, in bold, followed by a newline.

```csharp
var audit = sheet.Cell("C3").CreateComment().SetAuthor("Audit");
audit.AddSignature();
audit.AddText("Checked against invoice 4711.");
```

Authors are stored once per comments part. Two comments with the same author share an
author entry; setting the same name to both is enough. Keep `Author` non-null: the
comments writer uses it as a dictionary key and throws while saving if it is `null`.

### Plain and formatted runs

`AddText` appends a run and returns it. The one-argument overload uses the comment's
default font, Tahoma 9, not the font of the cell. Call methods on the returned
`IXlsxRichString` to format that run only. `AddText(text, font)` copies the formatting
of any `IXlsxFontBase`, for example another run or a style font.

```csharp
var note = sheet.Cell("D4").CreateComment();
note.AddText("Totals: ");
note.AddText("1,250").SetBold().SetFontSize(11);
note.AddNewLine();
note.AddText("Source: ledger").SetFontColor(XlsxColor.Green);

// Container-level setters change every existing run.
note.AddText("Recheck.");
note.SetItalic();
```

`Text` concatenates all runs, `Count` is the number of runs, and `Length` is the
character count. `ClearText` removes all runs; `ClearFont` keeps the text but resets
every run to the default comment font.

### Size, position, and visibility

| Member | Meaning |
| --- | --- |
| `Position.Row`, `Position.Column` | 1-based anchor of the note's top-left corner |
| `Position.RowOffset`, `Position.ColumnOffset` | Offset inside the anchor cell |
| `Style.Size.Width` | Width in the legacy 7.5-point unit; default is 19.2 (144 pt) |
| `Style.Size.Height` | Height in points; default is 59.25 |
| `Visible` | `false` (default) shows the note on hover; `true` pins it |
| `ZOrder` | Ordering among drawings on the worksheet |

A new note is anchored one column to the right and one row above its cell, with small
offsets. When the cell is in the first row, the anchor row is clamped to the first
row. Setting `Position.Row` and `Position.Column` moves the note independently of
the cell it comments.

### Reading comments

`HasComment` tells you whether a note exists. `GetComment` returns the existing note,
but it creates a new one when the cell has none, so check `HasComment` first if you
only want to read.

```csharp
using Aid.Xlsx.Model.Comments;
using Aid.Xlsx.Model.RichText;

var cell = sheet.Cell("B2");
if (cell.HasComment)
{
    IXlsxComment note = cell.GetComment();
    Console.WriteLine($"{note.Author}: {note.Text}");
    foreach (IXlsxRichString run in note)
        Console.WriteLine($"{run.Text} bold={run.Bold} size={run.FontSize}");
}
```

The iteration order is the order in which runs were added, including runs loaded from
a file. Reading does not remove or rewrite the note.

### Deleting comments

| Call | Effect |
| --- | --- |
| `comment.Delete()` | Removes the note from its cell |
| `cell.Clear(XlsxClearOptions.Comments)` | Removes the note and leaves the value |
| `range.DeleteComments()` | Removes notes from every cell of the range |
| `cells.DeleteComments()` | Removes notes from the enumerated cells |

```csharp
using Aid.Xlsx.Model.Grid.Cells;

sheet.Cell("B2").GetComment().Delete();
sheet.Range("B3:B20").DeleteComments();
sheet.Cell("C3").Clear(XlsxClearOptions.Comments);
```

`CreateComment` replaces an existing note rather than adding a second one. A cell
cannot hold two comments, so there is no collection of comments per cell and no
worksheet-level comment collection; enumerate cells and test `HasComment` when you
need to inspect all notes.

### Threaded comments

Modern threaded comments live in a separate package part with authors and replies.
Aid.Xlsx does not read, create, reply to, or update that part. Comment edits made
through this API therefore do not change threaded comments, and the legacy comment API
above only sees notes. If a workbook uses threaded comments only, `HasComment` is
`false` and the note is not exposed.

### What survives a round trip

| Content | Load and save | Notes |
| --- | --- | --- |
| Note text, runs, author | Preserved | Runs are read from the comments part |
| Note size, position, visibility | Preserved | Read from and written to the VML drawing |
| Note colours, line style, margins | Preserved | Part of the drawing style |
| Threaded comments | Not modeled | Separate part, not read or updated |
| VML attributes outside the model | Dropped on rewrite | Comment shapes are regenerated on save |

The comments part is rewritten whenever the worksheet contains notes. Text is written
from the modeled runs, so rich text properties the model does not carry (for example
`outline`, `condense`, and `extend`) are not emitted, and empty runs are skipped. The
VML part is regenerated for comment shapes as well; any shape attribute the model does
not load is lost, while non-comment shapes in the same VML part are kept. See
[Compatibility](compatibility.md) for the general preservation rules.

## Rich text

Use rich text when one cell must mix formatting, such as a bold figure inside a
sentence. The cell value stays a text value; formatting is attached to runs, not to
the cell style.

### Create rich text

`CreateRichText` replaces the cell value with a rich-text object built from the
current formatted string. Calling `ClearText` on the returned object discards that
first run, which is the simplest way to build the content from scratch.

```csharp
using System;
using Aid.Xlsx.Model.RichText;
using Aid.Xlsx.Model.Styles.Colors;
using Aid.Xlsx.Model.Workbooks;

using var workbook = new XlsxWorkbook();
var sheet = workbook.Worksheets.Add("Revenue");

var cell = sheet.Cell("A4");
cell.Value = "Revenue grew 12% in Q1";
IXlsxRichText rich = cell.CreateRichText();
rich.ClearText();
rich.AddText("Revenue ");
rich.AddText("grew 12%").SetBold().SetFontColor(XlsxColor.Green);
rich.AddText(" in Q1").SetItalic();

workbook.SaveAs("revenue.xlsx");
```

Every `AddText` call appends a run. A new run starts with the default font of the
rich text, which is the cell font at the time `CreateRichText` was called. Changes to
a run are written back to the cell immediately.

### Read rich text and plain text

`GetRichText` returns a mutable view of the runs. If the cell has no rich text, the
method creates it from the current value, exactly like `CreateRichText`, so
`HasRichText` is the test for "already formatted".

```csharp
using Aid.Xlsx.Model.RichText;

var cell = sheet.Cell("A4");
Console.WriteLine(cell.GetText());
Console.WriteLine(cell.GetFormattedString());

if (cell.HasRichText)
{
    foreach (IXlsxRichString run in cell.GetRichText())
        Console.WriteLine($"{run.Text} / bold={run.Bold} / size={run.FontSize}");
}
```

For a text cell, `GetText` and `GetFormattedString` return the same concatenated text:
the number format does not apply formatting to strings. Neither method reports the
run formatting; iterate the runs for that. `GetRichText` returns a new wrapper each
time, and two wrappers for the same cell can overwrite each other's whole content, so
finish edits through one wrapper.

### Run formatting reference

| Method | Property | Effect |
| --- | --- | --- |
| `SetBold(bool)` | `Bold` | Bold on or off |
| `SetItalic(bool)` | `Italic` | Italic on or off |
| `SetUnderline()`, `SetUnderline(value)` | `Underline` | Underline style, `Single` by default |
| `SetStrikethrough(bool)` | `Strikethrough` | Strike-through line |
| `SetShadow(bool)` | `Shadow` | Shadow effect |
| `SetVerticalAlignment(value)` | `VerticalAlignment` | Superscript or subscript |
| `SetFontSize(double)` | `FontSize` | Size in points |
| `SetFontColor(XlsxColor)` | `FontColor` | Run colour |
| `SetFontName(string)` | `FontName` | Font name |
| `SetFontFamilyNumbering(value)` | `FontFamilyNumbering` | Font family classification |
| `SetFontCharSet(value)` | `FontCharSet` | Character set |
| `SetFontScheme(value)` | `FontScheme` | Major or minor theme font |

`IXlsxFormattedText<T>` exposes setters for the same properties without the character
set and scheme. Calling one of those setters on the rich text object applies the value
to all existing runs; calling it on a run changes only that run.

### Split, rebuild, and copy

There is no "format characters 8 to 15" method. Read the plain text, clear the runs,
then add the pieces with the formatting you want. `ClearText` empties both the runs and
the cell text, so capture the old text first and add the new runs immediately.

```csharp
var cell = sheet.Cell("A4");
string text = cell.GetText();
IXlsxRichText rich = cell.GetRichText();
rich.ClearText();
rich.AddText(text[..8]);
rich.AddText(text[8..]).SetBold();
```

To duplicate formatted text on another cell, create rich text there and copy into it.
`CopyFrom` replaces the target's runs and keeps the target cell value in sync.

```csharp
var source = sheet.Cell("A4").GetRichText();
sheet.Cell("B4").CreateRichText().CopyFrom(source);
```

### Phonetic runs

Phonetic runs (furigana) display reading text above characters. Indices are 0-based
character positions into the cell text. They must be ascending, must not overlap, and
must not exceed the text length; otherwise writing the rich text throws
`ArgumentException`. Turn on `ShowPhonetic` on the cell to display them.

```csharp
var cell = sheet.Cell("A6");
cell.Value = "東京";
IXlsxRichText rich = cell.CreateRichText();
rich.Phonetics.Type = XlsxPhoneticType.Hiragana;
rich.Phonetics.Alignment = XlsxPhoneticAlignment.Center;
rich.Phonetics.Add("とうきょう", 0, 2);
cell.ShowPhonetic = true;
```

`HasPhonetics` reports whether phonetics exist without creating the property object.
The same phonetics object is available on comments.

### Shared strings and file size

Text cells are shared by default (`ShareString` is `true`). The shared-string table
deduplicates identical content, including the full run list: two cells with the same
text and the same run formatting occupy one entry, while the same visible text split
into different runs occupies two entries. A rich string is larger than plain text
because each run writes its own `rPr` element, so a sentence split into many small
runs grows the file noticeably.

```csharp
sheet.Cell("A4").ShareString = false;
```

With `ShareString = false` the rich text is written inline in the worksheet part. That
is useful when a consumer reads the sheet part directly or when strings should not be
collected in the shared table, but it repeats the full text for every cell that uses
it. Formula results that are text are always inline. Plain strings still use the
shared table unless you opt out, so leave the default for normal data.

Rich text and the cell value are stored separately. Assign the value first, call
`CreateRichText`, and do not assign `Value` again while you still hold the rich-text
object. The run list and the value can get out of sync: the cell keeps the old runs in
memory, while the new value determines what most of the writer emits. Rebuild the rich
text after changing the value instead of editing the old object.

### What survives a round trip

| Content | Load and save | Notes |
| --- | --- | --- |
| Run text, fonts, colours, sizes | Preserved | Written as shared-string runs |
| Run underline, strike, shadow, alignment | Preserved | |
| Phonetic runs and their properties | Preserved | Written as `rPh` and `phoneticPr` |
| Empty runs | Dropped | A run with empty text is not written |
| `outline`, `condense`, `extend` run flags | Dropped | Not represented by the model |
| Plain cell text | Preserved | Shared string entries are rebuilt |

Loading a workbook keeps existing runs as long as the cell remains a text cell.
Assigning a number to the cell replaces the text and therefore the runs in the saved
file. Calculation and cell copying preserve runs; clearing the cell contents does not.
See [Cells and ranges](cells-and-ranges.md) for clearing and copying, and
[Formatting and worksheet features](formatting-and-worksheets.md) for cell fonts that
act as the default for new runs.
