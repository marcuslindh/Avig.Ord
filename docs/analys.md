# Avig + Word — val för API:t

Underlaget om andra bibliotek finns i Knowledge: "World — Referens: Word-bibliotek för .NET
(NuGet) – läget november 2026". Besluten registrerades i Knowledge 2026-11-24.

## Valt (2026-11-24)

| Fråga | Val |
| --- | --- |
| Namn | **Avig.Ord** (namnrymd `Avig.Ord`) |
| Omfång i 1.0 | **Läsa och skriva**: läsare, skrivare, dokumentmodell och mallar |
| Delad kod med Avig.Kalkyl | **Avig.OpenXml**, eget paket, 1.0 samma dag som Kalkyl 1.0 |
| Strömmande läsning | **Block** (`ReadBlocksAsync`) och **händelser** (`ReadEventsAsync`) |
| Skapa dokument | **Ström** (`DocxWriter`) och **dokumentmodell** (`OrdDocument.Create`) med samma byggare |
| Typer | **Samma typer** i läsaren och dokumentmodellen; läsarens är skrivskyddade |
| Formatering vid läsning | **Verklig** (`Format`) som standard, rå (`DirectFormat`) finns |
| Enheter | **`Length`**, lagrad som EMU; i Avig.OpenXml från 0.2.0, delad med Kalkyl och Bildspel |
| Spårade ändringar vid läsning | **Som godkända** som standard; Original och Marked finns |
| Ändringar i dokumentmodellen | **Kan spåras** med `TrackChanges(author)` |
| Sök och ersätt | Text och regex, alla delar, ersätt med innehåll; inte över styckegränser |
| Textutvinning | `ToPlainText()` och `ToMarkdown()` (GFM); rubriker efter dispositionsnivå |
| Markdown-bilder | Länkar till `media/`, `SaveMarkdownAsync` sparar dem |
| Mallsyntax | **`{{…}}`** och Words **MERGEFIELD** och **innehållskontroller** |
| Saknat värde | **Tom sträng**, och listas i `TemplateResult.Missing` |
| Tal och datum i mallar | `TemplateOptions.Culture` (maskinens som standard) och format i platshållaren |
| Typade mallar | **Källgenerator** `[OrdTemplate]` i 1.0 |
| Filtyper | .docx, .dotx, .docm (läsa, ta bort makron) och Flat OPC |
| Bildformat | PNG, JPEG, GIF, BMP, SVG med PNG som reserv |
| Beroenden | net10.0, Avig.Result och Avig.OpenXml; säkert för trimning och AOT |

## Hur förebilderna gör, och varför Avig.Ord gör annorlunda

- **DocumentFormat.OpenXml** ger varje element som klass. Fullständigt, men man skriver själv
  arvet av formatering, numreringen, texten över körningar och mallarna. Avig.Ord gör det åt en
  och använder SDK:t bara i testerna, som schemavalidator och oberoende källa till filer.
- **DocX** och **OfficeIMO** har enkla API:n men undantag i stället för resultat, och DocX får
  bara användas icke-kommersiellt. Avig.Ord använder `Result` och `Try…` som resten av Avig.
- **NPOI** följer Apache POI:s Java-stil och har avgift för NuGet-binären.
- **Kommersiella bibliotek** (Aspose, Syncfusion, GemBox) gör också .doc, PDF och layout, vilket
  ligger medvetet utanför Avig.Ord.

## Följder av valen

- Modellen är vyer över `XElement` (System.Xml.Linq ingår i .NET och är AOT-säkert), så att allt
  okänt följer med orört när dokumentet sparas.
- Sök och ersätt och mallarna delar samma maskin: `ParagraphText` ger texten och var den kommer
  ifrån, och `ReplaceIn` delar körningar vid gränserna och byter innehållet.
- En del som lämnats ut som XML sparas från XML; resten byte för byte.
