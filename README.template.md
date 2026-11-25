# Avig.Ord

Read, write and edit Word documents (.docx, .dotx, .docm and Flat OPC) in .NET, without Word and
without other packages than [Avig.Result](https://www.nuget.org/packages/Avig.Result) and
[Avig.OpenXml](https://www.nuget.org/packages/Avig.OpenXml), the package layer it shares with
[Avig.Kalkyl](https://www.nuget.org/packages/Avig.Kalkyl) for Excel. Trimming and Native AOT safe.

- **Read** paragraphs, runs, tables, lists with their labels, headings, pictures, hyperlinks,
  fields, footnotes, comments, headers and footers, content controls and text boxes, with the
  formatting a run really has: the document defaults, the style chain and its own.
- **Write** a document block by block with constant memory, or **build** one in memory.
- **Edit** a document and keep what the package does not understand (SmartArt, charts,
  equations, other programs' parts) as it was.
- **Search and replace** also where Word has split the text over runs, with regex, and replace
  with text, a picture, paragraphs or a table.
- **Track changes**: read them as accepted, original or marked; accept or reject them; make your
  own edits as tracked changes.
- **Templates**: `{{Name}}` placeholders with formats, `{{#Lines}}…{{/Lines}}` sections that
  repeat table rows or paragraphs, conditions, MERGEFIELD fields and content controls, filled from
  a dictionary or from a class through a source generator.
- **Text out**: plain text, or GitHub Flavored Markdown with headings, lists, tables and pictures.
- **Files from outside** are opened with limits on size, parts and XML depth, without DTDs, and
  every failure has a kind you can show a user.

What it does not do: the old binary .doc format, PDF or page layout, and updating fields and
tables of contents (fields keep the result Word saved).

## Write a document

`DocxWriter` writes straight to the file: only the block being written, and the pictures, are in
memory.

<!-- example: write -->
```csharp
@@write@@
```

## Read a document

Paragraphs and tables are thin views over the document: `Text`, `Runs`, `Format` (as Word shows
it), `DirectFormat` (what the run itself sets), `StyleName`, `OutlineLevel`, `List` with the label
Word shows, and for a table `Rows`, `Cells`, `ColumnSpan` and `VerticalMerge`. Every index is
0-based, and every length is a `Length` (`Length.Cm(2.5)`, `Length.Pt(11)`, `.ToCm()`), so the
unit is never a guess.

<!-- example: read -->
```csharp
@@read@@
```

Every failure you should expect is a `Result`, not an exception: a missing file is `NotFound`, a
file another program holds is `Conflict`, and a file that is not a Word document, is damaged or is
too large is `Invalid`. `ErrorKind()` says which, in words you can show a user:

<!-- example: upload -->
```csharp
@@upload@@
```

## Text and Markdown

A paragraph with an outline level (Heading 1, Rubrik 1 or a style of your own) is a heading; lists
come out as `-` and `1.`, tables as GFM tables with the first row as header, emphasis as `**` and
`*`, and pictures as links into `media/`.

<!-- example: text-and-markdown -->
```csharp
@@text-and-markdown@@
```

## A large document

`DocxReader` reads one paragraph or table at a time, as read-only objects of the same types.
`ReadEventsAsync` goes further, element by element, for tables too large to hold at once.

<!-- example: stream -->
```csharp
@@stream@@
```

## Edit a document

<!-- example: edit -->
```csharp
@@edit@@
```

## Build a document in memory

<!-- example: build -->
```csharp
@@build@@
```

## Templates

A template is an ordinary Word document with `{{Name}}` where values go. `{{Name:format}}`
formats a number or date in the culture you give (the machine's when you give none), and
`{{Customer.Name}}` reaches into a nested value. `{{#Name}}` … `{{/Name}}` repeats what is
between for each item of a list — table rows when the tags are in a row, whole paragraphs when
they stand on paragraphs of their own — or keeps it once for `true`; `{{^Name}}` … `{{/Name}}`
keeps it when the value is false, empty or missing. Word's own MERGEFIELD fields (with `\@`,
`\#`, `\*`, `\b` and `\f`) and content controls (by tag, or title) are filled by name too, and a
`Replacement` as a value puts in a picture or a table.

<!-- example: template -->
```csharp
@@template@@
```

With a class instead of a dictionary, the source generator makes the mapping at compile time:

<!-- example: typed-template-type -->
```csharp
@@typed-template-type@@
```

<!-- example: typed-template -->
```csharp
@@typed-template@@
```

## Tracked changes

<!-- example: revisions -->
```csharp
@@revisions@@
```

## Files from outside

Avig.Ord never runs a macro, follows a link or updates a field, but a document it saves keeps
whatever it had. `ActiveContent` says what in a file can act on its own or fetch something when
Word opens it — the ways malware gets into documents: macros, ActiveX controls, embedded OLE
objects, fields such as DDE and INCLUDEPICTURE, and external references such as a template
attached from a server. `RemoveActiveContent` takes all of it out, and Word opens the result
without repair. Zip files that say one thing to Office and another to a scanner are refused when
they are opened.

<!-- example: clean -->
```csharp
@@clean@@
```

## The API at a glance

| Task | API |
| --- | --- |
| Open, save | `OrdDocument.OpenAsync`, `SaveAsync` (with `SaveOptions.Format` and `FlatOpc`), `OrdDocument.Create` |
| Read in a stream | `DocxReader.OpenAsync`, `ReadBlocksAsync`, `ReadEventsAsync`, `ReadPlainTextAsync`, `WriteMarkdownAsync` |
| Write in a stream | `DocxWriter.CreateAsync`, `ParagraphAsync`, `HeadingAsync`, `BulletAsync`, `NumberedAsync`, `TableAsync`, `PictureAsync`, `PageBreakAsync`, `SectionBreakAsync`, `Header`, `Footer` |
| Content | `Blocks`, `Paragraphs`, `Tables`, `Sections`, `Headers`, `Footers`, `Footnotes`, `Endnotes`, `Comments`, `Properties`, `Styles` |
| A paragraph | `Text`, `Inlines`, `Runs`, `StyleId`, `StyleName`, `Format`, `DirectFormat`, `OutlineLevel`, `List` |
| Inlines | `Run`, `Hyperlink`, `Field`, `Picture` (`ReadAsync`), `Break`, `NoteReference`, `TextBox`, `ContentControl` |
| Change | `AddParagraph`, `AddHeading`, `AddTable`, `InsertParagraphAfter`/`Before`, `InsertTableAfter`, `AppendText`, `SetText`, `SetStyle`, `Remove`, `AddRow`, `InsertCopyAfter` |
| Search | `Find`, `Replace` (text, regex, `Replacement`), `SearchOptions` (case, whole word, scope) |
| Changes | `TrackChanges`, `StopTrackingChanges`, `AcceptAllRevisions`, `RejectAllRevisions`, `ReadOptions.Revisions` |
| Templates | `FillTemplate`, `TemplateValues`, `[OrdTemplate]`, `[TemplateName]`, `TemplateResult.Missing` |
| Out | `ToPlainText`, `ToMarkdown`, `SaveMarkdownAsync` |
| Macros | `HasMacros`, `RemoveMacros` (macros are kept or removed, never run) |
| Active content | `ActiveContent`, `ExternalReferences`, `RemoveActiveContent`; `DocxReader.GetActiveContentAsync`, `GetExternalReferencesAsync` |
| Limits and errors | `OrdLimits`, `ErrorKind()`, `OrdErrorKind` |

## License

MIT
