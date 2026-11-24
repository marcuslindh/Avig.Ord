# Avig.Ord.Fuzz

Fuzz harness for Avig.Ord, built on [SharpFuzz](https://github.com/Metalnem/sharpfuzz) and
libFuzzer the way `Avig.Kalkyl.Fuzz` is. It has three modes:

- **smoke**: seeded random input with no coverage feedback, and nothing to install.
- **libFuzzer**: coverage-guided mutation; publish the harness, instrument `Avig.Ord.dll` with
  `sharpfuzz`, and run it under `libfuzzer-dotnet` as `Avig.Kalkyl.Fuzz/fuzz.ps1` does.
- **repro**: runs one saved input again and prints the whole exception.

The project is not part of the package: `publish.ps1` packs only `src/Avig.Ord`.

## Targets

Every seed comes from a document the package's own `DocxWriter` wrote (`Template.cs`): the
every-feature document of `WriterTests.EveryFeatureAsync`, with a nested table, notes, fields,
an SVG picture and two sections. The rich template adds what the writer does not make: tracked
changes, complex and nested fields, a symbol, content controls, a text box in alternate content,
a comment and a comments part.

| Target | Kind | What it checks |
|---|---|---|
| `docx-part` | crash and invariant | One part of the rich template (the body, styles, numbering, notes, header, footer, comments, core properties, settings, relationships, content types) is replaced by the input. The file is read with `DocxReader` (blocks, events, Markdown, text) in a revision view, opened as an `OrdDocument` and read through (every property of every block and inline, headers, notes, comments, sections, properties, styles), changed, saved as zip or Flat OPC, and opened again: the change must be there. |
| `zip-raw` | crash | The bytes of the rich template are mutated where the zip structure is: header fields, a truncated file, junk in front, repeated chunks. Same harness as `docx-part`. |
| `flat-opc` | crash and invariant | The every-feature document as `DocxWriter` writes it in Flat OPC, mutated as text, with Flat OPC tokens (duplicate parts, names outside the package, broken base64, two roots in `pkg:xmlData`). Same harness. |
| `roundtrip` | invariant | One part of the plain or rich template mutated. Open, save untouched (unread parts byte for byte) and save after reading (parts from XML): both read back with the same `ToPlainText` and `ToMarkdown`. When the input passes Microsoft's schema validator, the saved files do too, also after the changes every target makes. |
| `render` | crash | A body from the grammar in `BodyGenerator.cs`, now and then mutated, through the harness of `docx-part` in all three revision views. |
| `events` | crash | `ReadEventsAsync` on a generated, mutated body, in every view: no exception but the documented ones, no negative table depth, no empty text event. |
| `reader-model` | invariant | A generated body: `DocxReader` blocks and list labels, `ReadPlainTextAsync`, `WriteMarkdownAsync` and the paragraphs built from `ReadEventsAsync` equal what `OrdDocument` gives, in every view. |
| `revisions` | invariant | A generated body: `AcceptAllRevisions` gives the Accepted view's text (also after saving), `RejectAllRevisions` the Original view's. |
| `image-size` | crash and invariant | The header of a PNG, GIF, BMP, JPEG or SVG with edge sizes, cut short or with a byte changed, through `ImageInfo.TryGetSize`: no exception, a size in range, and the size a PNG header holds. |
| `format` | crash and invariant | `w:rPr` and `w:pPr` with edge values through `FormatReader`: no exception, and a whole number of twips gives that many twips or nothing, never a length of the other sign. |

`BodyGenerator` makes paragraphs with styles and list numbering (also levels and ids that do not
exist), runs with formatting, tabs, breaks, symbols and hyphens, hyperlinks, simple and complex
fields (nested, and now and then without their begin, separate or end), tracked insertions,
deletions and moves (also of rows and paragraph marks), content controls, smart tags, custom XML,
text boxes with blocks of their own, pictures, note references, bookmarks, and tables with merged
cells and tables inside them. Random bytes never become a body that gets past its first element.

### What counts as a finding

Nothing is swallowed by the runner. Each target catches only what the package documents:

- a refused file is a Result failure, and its `OrdError.ErrorKind()` must not be `Unknown`;
- while `DocxReader` reads the body, `InvalidDataException` or `XmlException`, again with a kind;
- once `OrdDocument.OpenAsync` has succeeded, **nothing** may throw: the file was accepted.

Every other exception is a finding, and so is an input that takes more than two seconds. An
invariant target that sees a wrong answer throws `InvariantViolationException`, its own type.

A smoke run ends with a line `Known: …` when Flat OPC saves hit the known Avig.OpenXml bug
below; the harness saves those as zip instead, so that the rest of the input is still tested.

## Findings

The first rounds (2026-11-24) found these. All are fixed and have regression tests in
`tests/Avig.Ord.Tests/FuzzFindingsTests.cs`:

| Target | Bug | Effect |
|---|---|---|
| `docx-part`, `zip-raw`, `roundtrip` | The model loads the parts other than the body when they are asked for. Broken XML in a header, the notes, the comments or the styles was not seen by `OpenAsync`. | `OpenAsync` succeeded, and then `Headers`, `Endnotes`, `Paragraphs` or `SaveAsync` as Flat OPC threw `XmlException`. Every XML part is now read through when the file is opened. |
| `render` | A complex field that began before a hyperlink and ended inside it took the hyperlink's list off the stack, and then the paragraph's own. | `InvalidOperationException: Stack empty` from `Paragraph.Text`, and the rest of the hyperlink came before its start. |
| `image-size` | `Math.Abs` of a BMP height of `int.MinValue`. | `OverflowException` when such a picture was added. |
| `reader-model` | `ReadEventsAsync` did not leave out rows marked deleted (or inserted, in the Original view), since it skipped `w:trPr` unread. | The events had the text of rows the view does not have. |
| `reader-model` | `ReadEventsAsync` had no event for `w:noBreakHyphen` and `w:sym`, gave an equation's `m:t` as text, and never ended an empty `<w:tbl/>`, `<w:tr/>` or `<w:tc/>`. | Text missing or extra; `TableDepth` wrong for the rest of the body. |
| `reader-model` | A field left open in one paragraph went on in the events, and a field begun inside a hyperlink did not end with it, while `Paragraph.Inlines` reads fields a paragraph and a container at a time. | Whole paragraphs missing from the events. |
| `reader-model` | `ReadPlainTextAsync` trimmed every line end at the end, `ToPlainText` only the last. | An empty last paragraph or a final line break read differently. |
| `reader-model` | `WriteMarkdownAsync` started a new renderer for each block, so it did not know a list had come before. | An empty paragraph between list items split the list in two. |
| `revisions` | `ReadDrawing` took every `w:txbxContent` under a drawing, also those inside its own text boxes. | A text box in a text box was read twice, and in rows the view leaves out. |
| `flat-opc` | Two relationships with one id in a part's relationships. | The file opened, and `SaveAsync` threw `InvalidOperationException`. It is now a damaged file. |
| `flat-opc` | A `pkg:xmlData` with two root elements: Avig.OpenXml 0.1.2 throws `InvalidOperationException` from its `XmlWriter` instead of giving Invalid. | `OpenAsync` threw. Worked around in `WordPackage` until it is fixed in Avig.OpenXml. |
| `roundtrip` | A section whose margins are wider than its page. | A table added to it had negative column widths, which the schema forbids. |
| `roundtrip` | New ids were the largest in use plus one. | With `int.MaxValue` in the file, a new list (and a note or a picture) got a negative id. |
| `render` | `Enumerable.Sum` over `gridSpan`, which a file can set to `int.MaxValue`; and Markdown padded every row to the `w:tblGrid`. | `OverflowException` from `ColumnCount` and `ToMarkdown`; a few megabytes of grid gave gigabytes of Markdown. Spans are now at most 63, as in Word. |
| `format` | Twips and half-points were multiplied into EMUs without a check. | `20000000000000000` twips came back as a negative length. |

Found by `HostileFilesTests` while the targets were being built, and fixed the same way:
XML nested past `OrdLimits.MaxXmlDepth` was `Damaged` instead of `TooLarge`; a list level with a
letter format and a start of `int.MaxValue` gave an 82-million-letter label for every item; and
20 000 complex fields in a row, each in the result of the one before, nested `Field` objects that
deep, so reading the text ended the process with a stack overflow (fields now nest at most 32
deep, in both readers).

Found by the Word check (`tools/Avig.Ord.WordCheck`): `Footnotes` and `Endnotes` listed notes
that nothing refers to any more, which Word does not show. They are now the notes the body refers
to, in its order.

### Open: in Avig.OpenXml

`OpcPackageWriter` re-reads every XML part for Flat OPC with its own depth limit of 100, while
`OrdLimits.MaxXmlDepth` lets a document of 256 levels open. Such a document cannot be saved as
Flat OPC (`XmlException`). It has to be fixed in Avig.OpenXml; the regression test
`A_document_nested_deeper_than_100_levels_saves_as_flat_opc` is skipped until then. The same fix
should give Invalid for a `pkg:xmlData` with two roots, so that the work-around can go.

## Checking that a target is alive

A target that cannot find anything reports zero findings just as confidently as a target that
tests bug-free code. Each one was checked by putting a fixed bug back, with `--seed 11`
(R01 … are the findings as numbered in `docs/testplan.md`, section 0):

| Target | Sabotage | Result |
|---|---|---|
| `docx-part` | do not read the XML parts through in `OrdDocument.ReadPackageAsync` (R01) | 5 findings in 7 inputs |
| `zip-raw` | the same | 5 findings in 93 inputs |
| `flat-opc` | accept two relationships with one id | 5 findings in 7,128 inputs |
| `roundtrip` | do not read the XML parts through (R01) | 5 findings in 6 inputs |
| `roundtrip` | let `TextWidth` be negative (R09) | 1 finding in 20,000 inputs |
| `roundtrip` | let `NextFreeId` wrap past `int.MaxValue` (R10) | **0 findings** in 20,000 inputs |
| `render` | the old `End` and `Close` of `InlineReader.State` (R02) | 5 findings in 744 inputs |
| `events` | `char.ConvertFromUtf32` of `w:sym` without the range check | 5 findings in 6 inputs |
| `reader-model` | no event for `w:noBreakHyphen` | 5 findings in 6 inputs |
| `reader-model` | do not leave out rows marked deleted | 5 findings in 28 inputs |
| `revisions` | look for text boxes in all of a drawing's descendants (R06) | 5 findings in 281 inputs |
| `image-size` | `Math.Abs` of the BMP height as an int (R03) | 5 findings in 625 inputs |
| `format` | no range check on twips (R12) | 5 findings in 45 inputs |

Two attempts showed where a target is weak:

- **The first `render` sabotage** took out only the guard in `End` and left the new `Close`,
  which already keeps containers on the stack: zero findings in 20,000 inputs, since nothing was
  broken any more. Putting the old pair back found the bug at once.
- **`roundtrip`'s schema invariant** says something only about inputs that are still valid
  after mutation, and a value such as `int.MaxValue` lands on the one attribute that matters
  rarely. The id bug (R10) was found once in a run of 2,000 and not in the sabotage run of
  20,000; the margin bug once in 20,000. A sweep over the integer attributes of the template, as
  Kalkyl's `attribute-sweep` does, would find them every time (docs/testplan.md, FUZ-02). Both
  have regression tests.

## Running

```powershell
# Every target, 20 000 inputs each
dotnet run --project tools/Avig.Ord.Fuzz -c Release -- all --smoke

# One target with a fixed seed, so that a finding can be reproduced
dotnet run --project tools/Avig.Ord.Fuzz -c Release -- reader-model --smoke --iter 50000 --seed 42

# Print the whole exception for a saved input; --dump writes the built file to open in Word
dotnet run --project tools/Avig.Ord.Fuzz -c Release -- docx-part --repro findings/docx-part/crash-00000001-000003.bin --dump C:\temp\dump
```

Findings are saved in `findings/<target>/` as `crash-<seed>-<iteration>.bin`, and an input that
takes more than two seconds as `slow-…`. A stack overflow ends the process before any `catch`
runs: run again with the same `--seed` and `--trace`, which writes every input to
`findings/<target>/last-input.bin` before it runs. `findings/` is not under version control; a
real finding becomes a regression test instead. `--write-corpus 16` writes seeds to
`corpus/<target>/` for libFuzzer.

## Adding a target

1. Write a `FuzzTarget` with a harness and a generator. The generator should mostly give valid
   input with a dense sprinkling of what tends to go wrong.
2. Add it to `Program.Targets`.
3. Run `--write-corpus 16` for it.
4. **Break the code on purpose** and check that the target finds it. Add the sabotage to the
   table above.
