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

# Fuzzning (tools/Avig.Ord.Fuzz, se dess README)
dotnet run --project tools/Avig.Ord.Fuzz -c Release -- all --smoke

# Riktiga Word (Windows med Word installerat)
dotnet run --project tools/Avig.Ord.WordCheck -- all      # create + check
```

`tools/Avig.Ord.Fuzz` är en SharpFuzz-harness som Avig.Kalkyl.Fuzz, och den ingår inte i
paketet. Tio mål: delar av en docx från `DocxWriter` som byts ut (`docx-part`), zip-bytes
(`zip-raw`), Flat OPC (`flat-opc`), öppna–spara–öppna med schemavalidering (`roundtrip`),
genererade kroppar (`render`, `events`, `reader-model`, `revisions`) och de små tolkarna
(`image-size`, `format`). Ett fynd är ett undantag utöver de dokumenterade, en avvisning utan
`OrdErrorKind`, en hängning eller ett fel svar. Kör smoke-läget efter ändringar i läsarna,
`InlineReader`, `EventReader`, renderarna, sparningen eller `FormatReader`. Ett fynd blir ett
regressionstest i `FuzzFindingsTests.cs`, och ett nytt mål ska vara verifierat genom sabotage
(tabellen i README:n).

`tools/Avig.Ord.WordCheck` styr Word via COM (sen bindning med `dynamic`). `create` låter Word
göra dokument med riktiga Word-funktioner och sparar dem med Words egen läsning (.json) i
`tests/Avig.Ord.Tests/WordFiles`, som `WordFilesTests` läser; kör det igen och checka in om
filerna ska ändras. `check` skriver dokument med paketet på alla sätt (DocxWriter, Create,
öppnad–ändrad–sparad, Flat OPC, mall, Word-filerna sparade om), öppnar dem i Word utan
reparation och jämför texten. Kör det efter ändringar i skrivaren, byggarna eller sparningen.
Jämförelsens regler (`Report.cs`, `OrdSide.cs`) länkas in i testerna. Testplanen är
`docs/testplan.md`.

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
- **När `OpenAsync` har lyckats får inget kasta.** Varje XML-del läses igenom vid öppning, även
  de som modellen laddar senare. Det som filen styr och som inte begränsas av XML-djupet har en
  egen gräns: fält nästlas högst `InlineReader.MaxFieldNesting` (32), `gridSpan` och
  Markdown-tabeller högst 63 kolumner, bokstavsetiketter högst 32 767. Nya id:n tas med
  `OrdDocument.NextFreeId`, aldrig största plus ett.
- **`DocxReader` ska ge samma svar som `OrdDocument`**: block, listetiketter, klartext, Markdown
  och händelsernas stycketext (`EventReader` följer `InlineReader`:s regler för fält, rader och
  textrutor). Fuzzmålet `reader-model` kontrollerar det.

## Konventioner

- **Explicita typer** — aldrig `var`
- **Filomfattande namnrymder**, **klammrar alltid**, explicita bool-jämförelser (`== true`)
- Inga namn på en bokstav (AVIG0002), ingen `?:` (AVIG0024), ingen metod som blandar null och
  värden (AVIG0007): använd `Try…` eller `Result`
- Kommentarer och XML-dokumentation på engelska, CRLF
