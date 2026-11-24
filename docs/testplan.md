# Avig.Ord — testplan för kvalitet och säkerhet

Version 1, 2026-11-24. Testplanen gäller NuGet-paketet Avig.Ord 0.1.0: strömläsaren
(`DocxReader`, med `ReadBlocksAsync`, `ReadEventsAsync`, `ReadPlainTextAsync` och
`WriteMarkdownAsync`), skrivaren (`DocxWriter`), dokumentmodellen (`OrdDocument`) och text- och
Markdown-utdata. Den följer upplägget i Avig.Kalkyls testplan (`G:\OneDrive\SkunkWorks\Prod\Lib\Avig.Kalkyl\docs\testplan.md`).

Varje testfall hör till en av tre grupper:

- **[A] Vanliga filer.** Testet körs med riktiga eller genererade Word-filer, i vanlig CI.
- **[B] Manipulerade filer.** Testet kräver filer som byggs för att vara skadliga eller
  trasiga. De byggs i minnet i testkoden (`HostilePackage` i `HostileFilesTests.cs`).
- **[C] Isolerad säkerhetsmiljö.** Testet kräver en container eller VM utan nätverk och med
  begränsat minne och CPU: zip-bomber i full storlek, långvarig belastning, tester som kan
  krascha processen.

Kör alla attacker kontrollerat, i en isolerad miljö och bara mot system du har behörighet att
testa. Paketet har inga nätverksanrop, och inget i planen kräver att det anropar något utanför
testmiljön.

---

## 0. Utgångsläge och fynd 2026-11-24

| Finns | Var | Täcker |
|---|---|---|
| Enhetstester | `tests/Avig.Ord.Tests` (`ReaderTests`, `WriterTests`) | Funktion: arv genom stilkedjan, listetiketter, revisionsvyer, fält, innehållskontroller, Strict, makron, mallar, Flat OPC. |
| Schemavalidering | `TestFiles.Validate`, `A_file_with_every_feature_passes_Microsofts_schema_validation` | Det som skrivs klarar Microsofts OpenXmlValidator. |
| Attackfiler | `HostileFilesTests.cs` | Zip-bomber, ljugande zip-huvuden, för många delar, nästling, DTD i varje del, stilloopar, absurda listor, relationer ut ur paketet, saknade delar, trasiga fält, symboler, Flat OPC, kalkylblad som huvuddel. |
| Fuzz-harness | `tools/Avig.Ord.Fuzz` | Tio mål med SharpFuzz, alla verifierade genom sabotage. Fynden står i dess README och har tester i `FuzzFindingsTests.cs`. |
| Word-kontroll | `tools/Avig.Ord.WordCheck`, `tests/Avig.Ord.Tests/WordFiles`, `WordFilesTests.cs` | Riktiga Word öppnar det paketet skriver, utan reparation, och läser samma text. Word gör egna filer som testerna läser. |
| `OrdLimits` | `src/Avig.Ord/OrdLimits.cs` | Högst 512 MiB uppackat, 10 000 delar och XML 256 nivåer djupt. |

Fynden från arbetet med milstolpe 5 (alla rättade, med regressionstester):

| # | Fynd | Hittat av | Rättning |
|---|---|---|---|
| R01 | `OpenAsync` lyckades med trasig XML i en del som laddas senare (sidhuvud, noter, kommentarer, stilar), och `Headers`, `Endnotes`, `Paragraphs` eller en Flat OPC-sparning kastade `XmlException`. | `docx-part`, `zip-raw`, `roundtrip` | Varje XML-del läses igenom när filen öppnas (`OrdDocument.CheckXml`). |
| R02 | Ett fält som började före en hyperlänk och slutade i den gav `InvalidOperationException: Stack empty` och fel ordning på texten. | `render` | `InlineReader.State` håller reda på sina listor; ett avslutat fälts resultat tas bort först när behållarna i det är slut. |
| R03 | `Math.Abs(int.MinValue)` för höjden i en BMP. | `image-size` | Via `long`. |
| R04 | `ReadEventsAsync` tog med rader markerade som borttagna, saknade `noBreakHyphen` och `sym`, tog med ekvationers `m:t`, avslutade aldrig tomma `<w:tbl/>`, `<w:tr/>`, `<w:tc/>`, och lät ett fält gälla över styckegränser och behållare. | `reader-model` | `EventReader` läser `w:trPr` innan raden börjar och följer `Paragraph.Inlines` regler. |
| R05 | `ReadPlainTextAsync` och `WriteMarkdownAsync` gav annan text än `ToPlainText` och `ToMarkdown` (radslut på slutet, en lista som delades av ett tomt stycke). | `reader-model` | Samma trimning; en renderare för hela kroppen (`MarkdownRenderer.ForStream`). |
| R06 | En textruta i en textruta lästes två gånger och utan hänsyn till sin rads revisionsvy. | `revisions` | `ReadDrawing` går inte in i textrutor. |
| R07 | Två relationer med samma id gav `InvalidOperationException` när filen sparades. | `flat-opc` | Skadad fil vid öppning. |
| R08 | Flat OPC med två rotelement i `pkg:xmlData` gav `InvalidOperationException` från Avig.OpenXml 0.1.2. | `flat-opc` | Rättat i Avig.OpenXml 0.1.4. |
| R09 | Marginaler bredare än sidan gav negativa kolumnbredder i en ny tabell. | `roundtrip` | Textbredden faller tillbaka på A4:s. |
| R10 | Nya id:n för listor, noter och bilder blev negativa när filen redan använde `int.MaxValue`. | `roundtrip` | `OrdDocument.NextFreeId`. |
| R11 | `gridSpan` på `int.MaxValue` gav `OverflowException`, och en `w:tblGrid` med 100 000 kolumner gav gigabyte Markdown. | `render` | Högst 63 kolumner per cell och i Markdown, som i Word. |
| R12 | 20000000000000000 twips blev en negativ längd. | `format` | Tal som inte ryms blir ingen längd. |
| R13 | XML djupare än `OrdLimits.MaxXmlDepth` gav `Damaged` i stället för `TooLarge`. | `HostileFilesTests` | `OrdError` känner igen djupgränsen. |
| R14 | En listnivå med bokstäver och start `int.MaxValue` gav en etikett på 82 miljoner tecken per punkt. | `HostileFilesTests` | Över 32 767 skrivs talet med siffror. |
| R15 | 20 000 komplexa fält i rad, vart och ett i resultatet av det förra, gav stackspill (processen dog). | `HostileFilesTests` | Fält nästlas högst 32 nivåer, i båda läsarna. |
| R16 | `Footnotes` och `Endnotes` listade noter som inget längre hänvisade till, vilket Word inte visar. | Word-kontrollen | Noterna kroppen hänvisar till, i dess ordning. |
| R17 | Efter sammanslagningen med mallarbetet: händelseläsaren gav ett tomt stycke för ett stycke som var helt borttaget i vyn, vilket blocken hoppar över. Och `revisions` jämförde text som `AcceptAllRevisions` och `RejectAllRevisions` med rätta slår ihop. | `reader-model`, `revisions` | `EventReader` läser stycket i minnet och frågar `Block.IsGoneInView`; `revisions` jämför styckenas text utan brytningar och etiketter, och hoppar över jämförelsen när ett fält går över en spårad styckemarkering. |

Rättat i Avig.OpenXml 0.1.4: `OpcPackageWriter` läste om varje XML-del för Flat OPC med en egen
djupgräns på 100, medan `OrdLimits.MaxXmlDepth` släpper igenom 256. Nu tar skrivaren
`MaxXmlDepth`, som Avig.Ord sätter till samma gräns som vid läsning.

---

## 1. Funktionella tester

### F-01 Filformat: DOCX, DOTX, DOCM, DOTM, Flat OPC och DOC [A]
- **Syfte:** paketet ska känna igen vad det kan läsa och tydligt avvisa resten.
- **Testdata:** .docx från Word och från paketet, .dotx, .docm med `vbaProject.bin`, Flat OPC
  (.xml) från Word och från paketet, Strict, .doc (OLE), en lösenordsskyddad .docx, .xlsx och
  .pptx med namnet .docx, en HTML-fil, en tom fil.
- **Förväntat:** Word-formaten läses; OLE ger `NotWord` med tipset om .doc och lösenord; ett
  kalkylblad ger `NotWord`; en tom eller avklippt fil ger en felkod. Makron sparas med eller
  tas bort med `RemoveMacros`, och körs aldrig.
- **Täckt av:** `ReaderTests` (makron, mallar, Strict, kalkylblad), `WriterTests` (Flat OPC,
  mall), `HostileFilesTests` (Flat OPC-varianter, kalkylblad, avklippta filer), Word-kontrollen
  (.dotx och Flat OPC öppnas i Word). **Saknas:** .pptx och en krypterad fil från Word.

### F-02 Tomma dokument [A]
- **Testdata:** en kropp utan stycken, `<w:body/>`, ett dokument utan `w:body`, tomma stycken,
  celler och tabeller.
- **Förväntat:** tom text, `Blocks` tomt; `AddParagraph` fungerar och filen sparas giltig.
- **Täckt av:** `HostileFilesTests` (tomma `<w:tbl/>`), `FuzzFindingsTests`. **Saknas:** ett
  dokument utan `w:body` som eget test.

### F-03 Stora dokument [A][C]
- **Syfte:** strömläsaren ska läsa med minne som motsvarar det största blocket; dokumentmodellen
  med minne proportionellt mot dokumentet.
- **Testdata:** 100 000 stycken, en tabell med 1 miljon celler, 1 000 bilder.
- **Observera:** toppminne och tid för `ReadBlocksAsync`, `ReadEventsAsync` och `OpenAsync`.
- **Saknas.** Prestandamätning med BenchmarkDotNet som i Kalkyl.

### F-04 Formatering genom stilkedjan [A]
- **Testdata:** docDefaults, styckestilar i flera led, teckenstilar, direkt formatering, stilar
  som saknas, loopar, kedjor på 5 000 stilar.
- **Förväntat:** Words värden; en loop eller kedja längre än 64 bryts utan fel.
- **Täckt av:** `ReaderTests`, `HostileFilesTests.Styles_based_on_each_other_…`.

### F-05 Listor [A]
- **Testdata:** flernivålistor från Word (1), a), i)), punktlistor, `startOverride`, listor över
  flera `w:num`, nivåer och start utanför gränserna.
- **Förväntat:** samma etikett som Words `ListString`; korta etiketter för absurda värden.
- **Täckt av:** `ReaderTests`, `WordFilesTests` (`lists.docx`), `HostileFilesTests`.

### F-06 Spårade ändringar [A]
- **Testdata:** infogat, borttaget, flyttat, raderade och infogade rader, ändrade styckemärken,
  formateringsändringar; från Word och genererat.
- **Förväntat:** varje vy (`Accepted`, `Original`, `Marked`) ger samma text som Word i
  motsvarande läge; `AcceptAllRevisions` ger `Accepted`-vyn och `RejectAllRevisions`
  `Original`-vyn.
- **Täckt av:** `ReaderTests`, `WordFilesTests` (`tracked-changes.docx`), fuzzmålet `revisions`.

### F-07 Fält, hyperlänkar, innehållskontroller, textrutor och bilder [A]
- **Täckt av:** `ReaderTests`, `WordFilesTests` (`fields-headers-footers`,
  `content-controls-text-box`, `pictures-symbols`), `FuzzFindingsTests`.

### F-08 Sidhuvuden, sidfötter, noter och kommentarer [A]
- **Täckt av:** `ReaderTests`, `WriterTests`, `WordFilesTests` (`comments-and-notes`).

### F-09 Specialtecken och Unicode [A]
- **Testdata:** tecken XML inte tillåter, ensamma surrogat, `w:sym` med alla slags värden, hårt
  bindestreck, tabb och radbrytning, emoji.
- **Täckt av:** `WriterTests.Text_with_characters_xml_cannot_hold…`,
  `HostileFilesTests.A_symbol_that_is_no_character…`, fuzzmålens tecken.

### F-10 Strömläsaren och dokumentmodellen ger samma svar [A]
- **Förväntat:** samma blocktext, listetiketter, klartext och Markdown, och händelsernas
  stycketext lika med styckenas.
- **Täckt av:** `ReaderTests`, fuzzmålet `reader-model`, `FuzzFindingsTests`.

### F-11 Det som skrivs öppnas i Word [A]
- **Förväntat:** Word öppnar filer från `DocxWriter`, `OrdDocument.Create`, öppnade-ändrade-sparade
  filer, Flat OPC och mallar utan reparation och läser samma text som paketet.
- **Täckt av:** `tools/Avig.Ord.WordCheck check` (avsnitt 6.4). Inte i CI, eftersom det kräver Word.

### F-12 Korrupta och felformaterade filer [B]
- Se avsnitt 2 och 4.

---

## 2. Säkerhetstester och attackscenarier

### 2.1 DoS och överbelastning

**S-01 Zip-bomb och falskt zip-huvud [B][C]**
- **Attack:** en del som packas 1000:1, eller ett huvud som anger en annan storlek än den verkliga.
- **Test:** 600 MB nollor (under en megabyte packat) mot standardgränsen; ett huvud på 2 MB som
  anger 0, 100 och 0xFFFFFFFF byte, med gränsen 1 MB.
- **Förväntat:** `TooLarge` innan något packas upp, under 200 MB allokerat. Ett huvud som anger
  för lite läses aldrig längre än det säger (delen blir kort); ett som anger för mycket ger `TooLarge`.
- **Täckt av:** `HostileFilesTests.A_zip_bomb_past_the_limit…`, `A_part_whose_headers_lie…`.
  **Saknas:** 10 GB i [C].

**S-02 För många delar [B]**
- **Test:** 10 001 delar. **Förväntat:** `TooLarge`. **Täckt av:** `HostileFilesTests`.

**S-03 Djup nästling [B]**
- **Attack:** tabeller i tabeller, hyperlänkar i hyperlänkar, textrutor i textrutor, fält i fält.
- **Förväntat:** djupare än `OrdLimits.MaxXmlDepth` ger `TooLarge`; precis innanför läses utan
  stackspill; logisk nästling som XML-djupet inte begränsar (fält) har en egen gräns (32).
- **Täckt av:** `HostileFilesTests` (nästlade tabeller, hyperlänkar, textrutor, 20 000 fält).
- **Not:** `ToPlainText` skriver tabeller åtta nivåer djupt och utelämnar djupare, avsiktligt.

**S-04 Förstärkning i utdata [B]**
- **Attack:** liten indata som ger stor utdata: listetiketter med bokstäver, `gridSpan`,
  `w:tblGrid` med 100 000 kolumner.
- **Förväntat:** etiketter under 200 tecken, Markdown och text proportionella mot indata.
- **Täckt av:** `HostileFilesTests.List_definitions_with_absurd…`, `FuzzFindingsTests.Spans_and_grids…`.

**S-05 Samtidighet [C]**
- **Saknas.** Parallella uppladdningar med `dotnet-counters`, som i Kalkyl SEC-05.

### 2.2 XML- och zip-attacker

**S-06 DTD, XXE och billion laughs [B]**
- **Test:** en DTD med nästlade entiteter i varje XML-del: kropp, stilar, numrering, sidhuvud,
  fotnoter, kommentarer, inställningar, core-egenskaper, relationer och content types, och i Flat OPC.
- **Förväntat:** filen avvisas med en felkod inom en sekund; entiteter expanderas aldrig.
- **Täckt av:** `HostileFilesTests.A_dtd_in_any_xml_part…`, `Flat_opc_files_made_to_hurt…`.

**S-07 Relationer ut ur paketet och till delar som saknas [B]**
- **Test:** `../../../../Windows/win.ini`, `/../secret.xml`, `C:\…`, `file:///…`, UNC, `%2e%2e/`,
  en tom målsökväg, en del som saknas; för sidhuvud, bild, stilar, numrering och fotnoter.
- **Förväntat:** läses som saknade; `Picture.ReadAsync` ger `NotFound`; ingen fil utanför läses.
- **Täckt av:** `HostileFilesTests.Relationships_that_climb_out…`, `A_picture_whose_part_is_missing…`.

**S-08 Dubbla relations-id och dubbla delar [B]**
- **Förväntat:** `Damaged`. **Täckt av:** `FuzzFindingsTests`, `HostileFilesTests` (Flat OPC).

**S-09 Trasiga zip-arkiv [B]**
- **Täckt av:** fuzzmålet `zip-raw`, `HostileFilesTests.Truncated_and_empty_files…`.

### 2.3 Word-specifika attacker

**S-10 Makron [A]**
- Makron sparas eller tas bort men körs aldrig (`ReaderTests.Macros_are_kept_or_removed…`).

**S-11 Fält [B]**
- Fält utan slut eller början, okända typer, fält i instruktioner: läses utan fel och ger
  samma text i båda läsarna. Paketet räknar aldrig om fält och följer aldrig `INCLUDETEXT`,
  `INCLUDEPICTURE` eller `HYPERLINK`. **Täckt av:** `HostileFilesTests`, fuzzmålen.

**S-12 Externa bilder och länkar [B]**
- En bild med `r:link` eller extern relation läses aldrig från nätet; `ReadAsync` ger `NotFound`.
  **Saknas:** ett eget test med `TargetMode="External"` för en bild.

### 2.4 Övriga risker

**S-13 Felmeddelanden [A]**
- Meddelandena är loggtext och kan citera filen; `OrdErrorKind` är det som visas för användare.
  Fuzzmålen kräver en felkod för varje avvisad fil.

**S-14 Undantag utanför kontraktet [B]**
- Efter lyckad `OpenAsync` får inget kasta. Strömläsaren får kasta `InvalidDataException` eller
  `XmlException`, med felkod. Kontrolleras av alla fuzzmål.

---

## 3. Prestanda och belastning

**Saknas.** Gör som Kalkyl: BenchmarkDotNet med `[MemoryDiagnoser]` för öppning, strömläsning,
`ToPlainText`, `ToMarkdown` och sparning av ett dokument på 100 sidor och ett på 10 000, och
spara baslinjer i `docs/benchmarks/`.

---

## 4. Fuzzning och negativa tester

`tools/Avig.Ord.Fuzz` (se dess README för mål, sabotage och fynd):

| ID | Område | Finns | Saknas |
|---|---|---|---|
| Z-01 | Muterade delar i en giltig docx | `docx-part`: 13 delar av en rik mall från `DocxWriter`. | Fler startfiler: Word-filerna i `WordFiles`. |
| Z-02 | Trasiga zip-arkiv | `zip-raw`. | — |
| Z-03 | Flat OPC | `flat-opc`. | — |
| Z-04 | Öppna, spara, öppna | `roundtrip`, med schemavalidering. | Invarianten för validering är svag för sällsynta värden (sabotaget av `NextFreeId` gav inget på 20 000 indata); ett svep över heltalsattribut som Kalkyls `attribute-sweep`. |
| Z-05 | Genererade kroppar | `render`, `events`, `reader-model`, `revisions` med `BodyGenerator`. | Matematik (`m:oMath`), `mc:AlternateContent` med annat än ritningar. |
| Z-06 | Små tolkare | `image-size`, `format`. | — |
| Z-07 | Täckningsstyrning | Harnessen fungerar med libFuzzer. | Skript (`fuzz.ps1`, `fuzz-night.ps1`, `corpus-check.ps1`) som i Kalkyl, och en natts körning. |

Grundregeln: **varje mål ska kunna hitta ett fel som läggs in med avsikt.**

---

## 5. Beroenden och leveranskedja

| Område | Läge |
|---|---|
| Beroenden | Bara Avig.Result och Avig.OpenXml; `publish.ps1` avbryter annars. Verktygen och testerna får använda DocumentFormat.OpenXml och SharpFuzz. |
| Kända sårbarheter | NuGet Audit på nivån `low` med varningar som fel, och låsfiler (`Directory.Build.props`). |
| Avig.OpenXml | 0.x från den lokala feeden. Två fel hittade här ska rättas där (R08 och det öppna djupfelet), och då ska både Kalkyls och Ords tester köras. |

---

## 6. Automatisering och verktyg

### 6.1 Attackfiler (xUnit) [B]
`HostileFilesTests.cs` med `HostilePackage`, som bygger paket del för del. Klassen körs inte
parallellt med andra, så att allokeringen den mäter är dess egen.

### 6.2 Zip-bomb i full storlek [C]
Generera en del som strömmas genom `DeflateStream` (som i Kalkyls 6.2) och öppna den i en
container med `--memory=2g --network none`.

### 6.3 Krascher i egen process [C]
**Saknas:** `corpus-check.ps1` som i Kalkyl, som kör varje sparad indata i en egen process med
tidsgräns och `DOTNET_GCHeapHardLimit`. Ett stackspill (som R15) syns bara så.

### 6.4 Word-kontrollen [A]
`dotnet run --project tools/Avig.Ord.WordCheck -- all` (Windows med Word):

- `create`: Word gör sju dokument via COM och sparar dem med Words egen läsning som .json i
  `tests/Avig.Ord.Tests/WordFiles`.
- `check`: paketet skriver dokument på alla sina sätt; Word öppnar varje fil skrivskyddat utan
  reparation (`OpenAndRepair: false`, `DisplayAlerts = 0`) och texten jämförs stycke för stycke,
  med listetiketter, sidhuvuden, sidfötter, noter, kommentarer, textrutor och antal bilder.
  Word stängs även när något går fel.

Jämförelsens konventioner står i `Report.cs` och `OrdSide.cs`, som testerna länkar in: Words
kontrolltecken översätts, tomma stycken räknas inte, en bild räknas i stället för Words "/", en
symbol jämförs som Words "(", och tal i sidhuvuden och sidfötter maskeras eftersom Word räknar om
PAGE och NUMPAGES.

---

## 7. Riskmatris

| # | Attack eller feltyp | Sannolikhet | Konsekvens | Risk | Test | Status |
|---|---|---|---|---|---|---|
| 1 | Stackspill av logiskt nästlade fält (R15) | Medel | Hög: processen dör | **Kritisk** | `HostileFilesTests` | Åtgärdad 2026-11-24 |
| 2 | Minnesbomb via listetiketter och tabellnät (R11, R14) | Medel | Hög: OOM från en liten fil | **Kritisk** | `HostileFilesTests`, `FuzzFindingsTests` | Åtgärdad 2026-11-24 |
| 3 | Undantag efter lyckad öppning (R01, R02, R07, R11) | Hög | Medel: begäran kraschar utanför kontraktet | Hög | alla fuzzmål | Åtgärdad 2026-11-24 |
| 4 | Zip-bomb eller falskt huvud | Hög | Hög | Hög | S-01 | Skyddat av `OrdLimits` och Avig.OpenXml, testat |
| 5 | XXE och billion laughs | Hög | Hög | Låg, redan skyddat | S-06 | Verifierat i varje del |
| 6 | Relationer ut ur paketet | Medel | Hög | Låg | S-07 | Verifierat |
| 7 | Fel felkod (R13) | Medel | Låg | Låg | `HostileFilesTests` | Åtgärdad |
| 8 | Strömläsaren och modellen ger olika text (R04–R06) | Hög | Medel: fel data i sökindex och export | Medel | `reader-model`, `revisions` | Åtgärdad |
| 9 | Flat OPC djupare än 100 kan inte sparas | Låg | Medel | Låg | `FuzzFindingsTests` | Rättad i Avig.OpenXml 0.1.4 |
| 10 | Samtidiga uppladdningar | Hög | Medel | Medel | S-05 | Inte testat |

---

## 8. Acceptanskriterier för 1.0

1. **Inga okontrollerade krascher.** Varje fil i korpusen, attackfilerna och Word-filerna ger
   framgång, ett `Result` med felkod, eller (bara i strömläsaren) `InvalidDataException` eller
   `XmlException` med felkod. Inget stackspill och ingen hängning över 30 sekunder, mätt i en
   egen process. Uppfyllt för smoke-körningarna 2026-11-24; `corpus-check.ps1` saknas.
2. **Ingen oavsiktlig filåtkomst och inget nätverk.** Relationer ut ur paketet läses aldrig
   (uppfyllt, S-07). Saknas: körning med Process Monitor eller `strace`.
3. **Begränsat minne och begränsad tid.** Med `OrdLimits.Default` ska ingen fil under 10 MB ge
   OOM med 1 GB heap eller ta mer än 10 sekunder. Uppfyllt för attackfilerna.
4. **Felkoder.** Varje avvisad fil har en `OrdErrorKind` som inte är `Unknown`, och en gräns i
   `OrdLimits` ger `TooLarge`. Uppfyllt.
5. **Samma svar från båda läsarna.** `reader-model` och `revisions` utan fynd på 20 000 indata.
   Uppfyllt.
6. **Word.** `WordCheck all` är grön: Word öppnar allt paketet skriver utan reparation och läser
   samma text, och `WordFilesTests` är gröna. Uppfyllt 2026-11-24 med Word 16.0.20625.
7. **Schemat.** Det som skrivs klarar OpenXmlValidator, också när en giltig fil öppnas, ändras
   och sparas (`roundtrip`). Uppfyllt.
8. **Fuzzning.** Varje mål är verifierat genom sabotage och har körts minst 20 000 indata i
   smoke-läge utan fynd. Före 1.0: täckningsstyrt minst 8 timmar per mål, och korpusen incheckad.
9. **Avig.OpenXml.** Felen R08 och djupgränsen för Flat OPC är rättade där, arbetsgången i
   `WordPackage` borttagen och testet utan `Skip`. Uppfyllt 2026-11-24 (Avig.OpenXml 0.1.4).
10. **Prestanda.** Baslinjer finns (avsnitt 3).

---

## 9. Testplan i prioritetsordning

| Test-ID | Testnamn | Risk | Grupp | Steg | Förväntat | Åtgärd vid brist |
|---|---|---|---|---|---|---|
| SEC-02 | `corpus-check.ps1` | Hög | C | Varje sparad indata i egen process med 1 GiB heap och 30 s. | Inga krascher. | Gör om fyndet till ett test. |
| SEC-03 | Zip-bomb i full storlek | Hög | C | 10 GB i container. | `TooLarge` direkt. | Kontrollera Avig.OpenXml. |
| SEC-04 | Samtidighet | Medel | C | 32 parallella öppningar av blandade filer. | Stabilt minne. | Dokumentera begränsning. |
| SEC-05 | Extern bild | Låg | B | `TargetMode="External"` för en bild. | `NotFound`, inget nätverk. | Rätta `Picture`. |
| FUZ-01 | Täckningsstyrd fuzzning | Medel | C | libFuzzer 8 timmar per mål. | Inga fynd. | Test i `FuzzFindingsTests`. |
| FUZ-02 | Svep över heltalsattribut | Medel | B | Varje heltalsattribut i mallen gånger varje gränsvärde. | Inga fynd. | Validera vid läsning. |
| FUN-01 | Fler Word-filer | Medel | A | `WordCheck create` med fler funktioner: avsnitt, kolumner, SmartArt, diagram, ekvationer. | Samma text som Word. | Rätta och lägg till test. |
| PRF-01 | Baslinjer | Låg | A | BenchmarkDotNet. | Baslinjer finns. | Profilera. |
