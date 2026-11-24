# Avig.Ord — plan till 1.0

Planen och besluten står i Knowledge, projekt `avigord`: "Avig.Ord — Arkitektur: Plan till 1.0"
och projektets beslut. Den här filen är lägesbilden i repot.

| # | Milstolpe | Status |
| --- | --- | --- |
| 0 | Avig.OpenXml: paketlagret utbrutet ur Avig.Kalkyl | Klar 2026-11-24 (Kalkyl på grenen `avig-openxml`) |
| 1 | Läsare: block och händelser, stilarv, numrering, tabeller, noter, fält, spårade ändringar, ToPlainText, ToMarkdown | Klar 2026-11-24 |
| 2 | Skrivare: DocxWriter, byggarna, bilder, länkar, noter, sektioner, sidhuvud och sidfot, Flat OPC | Klar 2026-11-24 |
| 3 | Dokumentmodell: öppna, ändra, spara, sök och ersätt över körningar, spårade ändringar | Klar 2026-11-24 |
| 4 | Mallar: `{{…}}`, avsnitt, villkor, MERGEFIELD, innehållskontroller, källgenerator | Klar 2026-11-24 |
| 5 | Robusthet: fuzzning, fientliga filer, kontroll i Word, testplan | Pågår (grenen `robustness`) |

## Arkitektur i korthet

Se `CLAUDE.md`. Modellen är tunna vyer över `XElement`; `DocumentContext` bär namnrymd, stilar,
numrering, relationer och revisionsvy; `InlineReader` läser ett stycke, `ParagraphText` ger den
sökbara texten; byggarna skriver XML genom `IContentSink`.

## Medvetna begränsningar

- **.doc, PDF och sidlayout** ingår inte. Fält räknas inte om; de visar resultatet Word sparade.
- **Ett stycke vars styckemarkering ensam är borttagen** (spårat) visas i Accepted-vyn som ett
  eget stycke, inte ihopslaget med nästa. `AcceptAllRevisions` slår ihop dem.
- **Händelseläsaren** (`ReadEventsAsync`) ger start och slut även för stycken som i vyn är helt
  borttagna; blockläsaren och dokumentmodellen hoppar över dem.
- **Avsnitt inuti ett stycke** (`{{#X}}…{{/X}}` på samma rad) fylls som text, i formateringen där
  avsnittet börjar.
- **Sök och ersätt** går inte över styckegränser.
- **Formateringsarv:** växlande egenskaper (fet, kursiv) följer "sist satt gäller", inte Words
  XOR-regel mellan styckestil och teckenstil.
- **Bilder:** PNG, JPEG, GIF, BMP och SVG (med PNG som reserv) kan läggas in; EMF/WMF bevaras men
  läggs inte in. Diagram, SmartArt och ekvationer bevaras men blir inte objekt.
