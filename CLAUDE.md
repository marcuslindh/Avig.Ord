# CLAUDE.md

Knowledge-projekt: avigord
Dokumentation, beslut och agentlogg för projektet finns i Knowledge — se ~/.claude/CLAUDE.md.
Planen är "Avig.Ord — Arkitektur: Plan till 1.0"; besluten om API:t står i projektets beslut.

## Vad det här är

NuGet-paketet Avig.Ord: läsa, skriva och redigera Word-filer (.docx, .dotx, .docm, Flat OPC),
som ett eget alternativ till DocX, NPOI och de kommersiella biblioteken. Systerpaket till
Avig.Kalkyl (`G:\OneDrive\SkunkWorks\Prod\Lib\Avig.Kalkyl`), som det följer i allt.

## Kommandon

```powershell
dotnet build Avig.Ord.slnx
dotnet test Avig.Ord.slnx
./publish.ps1                 # testar, packar, formkontrollerar, lägger i G:\LocalNugetRepo
```

## Arkitektur i korthet

- **Avig.OpenXml** (eget paket, `G:\OneDrive\SkunkWorks\Prod\Lib\Avig.OpenXml`) gör paketlagret:
  zip och Flat OPC, gränser, relationer, content types, XmlReader utan DTD och med djupgräns.
  Det är 0.x i `G:\LocalNugetRepo` (se `nuget.config`).
- **Modellen** (`Blocks.cs`, `Inlines.cs`) är tunna vyer över `XElement`: `Paragraph`, `Table`,
  `TableRow`, `TableCell`, `Run`, `Hyperlink`, `Field`, `Picture` … Samma typer i
  `DocxReader` (skrivskyddade, `Context.Owner` är null) och `OrdDocument`.
- **`DocumentContext`** bär det en vy behöver: namnen (Transitional eller Strict, från delens
  rot), stilarna, numreringen, delens relationer, revisionsvyn och ägaren.
- **`InlineReader`** läser ett stycke till inlines i revisionsvyn: körningar (tabb och radbrytning
  i texten), fält (komplexa fält är en stack), hyperlänkar, bilder, textrutor, innehållskontroller.
- **`StyleSheet`** räknar fram verklig formatering: docDefaults → styckestilens kedja →
  teckenstilen → direkt formatering. En `basedOn`-loop bryts.
- **`OrdDocument`** håller alla delar i minnet; XML-delar laddas som `XDocument` först när de
  behövs och sparas från XML bara om de lämnats ut. Allt annat sparas byte för byte.
- **Byggarna** (`Builders.cs`) gör XML genom `IContentSink`, som både `DocxWriter` och
  `OrdDocument` har. Stilar som används läggs till vid sparning från `BuiltInStyles`.

## Regler som betyder mest

- **Inga beroenden utom Avig.Result och Avig.OpenXml.** `publish.ps1` failar annars. Testerna
  får använda DocumentFormat.OpenXml (schemavalidering, och som oberoende källa till filer).
- **0-baserat** i det publika API:t, **Result för förväntade fel**, **bara async för I/O**,
  `CancellationToken` förs vidare.
- **Säkert för trimning och AOT.** Ingen reflektion; den typade mallfyllningen är en källgenerator.
- **Läs på lokalt namn i delens namnrymd** via `WordNames`, aldrig med hårdkodad Transitional-namnrymd,
  så att Strict fungerar och sparas som Strict.
- **En fil utifrån får inte kunna ta slut på minne eller stack.** All läsning går via
  `OpcXml.CreateReader` med `OrdLimits.MaxXmlDepth`. Rekursion över XML är begränsad.
- **Felkoder:** meddelanden om skadade filer börjar med `The file is damaged`
  (`OpcErrors.DamagedPrefix`), om gränser innehåller `the limit set in OrdLimits.`, och om filer
  som inte är Word börjar med `The file is not a Word document`. `OrdError.ErrorKind()` läser av dem.
- **Det som skrivs ska klara Microsofts schemavalidator** (`TestFiles.Validate`), i schemats
  ordning för barnelement (rPr, pPr, style …).
- **Ett fel om en fil blir ett test** (`HostileFilesTests.cs`, `FuzzFindingsTests.cs`).
- **README-exemplen är tester.** Skriv README i `README.template.md`, där `@@namn@@` står för
  koden i `#region example: namn` i `tests/Avig.Ord.Tests/ReadmeExamples.cs`, och gör README.md
  med `dotnet run tools/readme.cs`. `Every_README_example_is_code_that_runs_here` failar om de
  skiljer sig. Ett nytt publikt API ska få ett exempel och en rad i "The API at a glance".
- **Mallar och sök och ersätt går genom `ParagraphText`**, som ger styckets synliga text och
  vilka w:t den kommer från, och `ReplaceIn`, som delar körningar och spårar ändringar. Gör inga
  egna textsökningar i XML:en.
- **Spårade ändringar:** allt som ändrar innehåll frågar `Context.Owner?.Tracker` och märker då
  det nya som infogat och det gamla som borttaget (`Tracking`). Kopior görs rena med
  `Revisions.Accept` innan de märks, så att revisions-id inte dubbleras.

## Konventioner

- **Explicita typer** — aldrig `var`
- **Filomfattande namnrymder**, **klammrar alltid**, explicita bool-jämförelser (`== true`)
- Inga namn på en bokstav (AVIG0002), ingen `?:` (AVIG0024), ingen metod som blandar null och
  värden (AVIG0007): använd `Try…` eller `Result`
- Kommentarer och XML-dokumentation på engelska, CRLF
