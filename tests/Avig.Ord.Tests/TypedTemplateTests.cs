using System.Globalization;

namespace Avig.Ord.Tests;

[OrdTemplate]
public partial class Invoice
{
    public string Customer { get; init; } = string.Empty;

    [TemplateName("Förfallodag")]
    public DateOnly Due { get; init; }

    public IReadOnlyList<InvoiceLine> Lines { get; init; } = [];

    public bool Paid { get; init; }

    public decimal Total => Lines.Sum(line => line.Amount);
}

[OrdTemplate]
public partial class InvoiceLine
{
    public string Item { get; init; } = string.Empty;

    public decimal Amount { get; init; }
}

[OrdTemplate]
public partial record Receipt(string Number);

public class TypedTemplateTests
{
    [Fact]
    public async Task A_class_marked_as_a_template_fills_by_its_properties()
    {
        const string body = """
            <w:p><w:r><w:t>Faktura till {{Customer}}, förfaller {{Förfallodag:yyyy-MM-dd}}</w:t></w:r></w:p>
            <w:tbl><w:tblPr/><w:tblGrid><w:gridCol w:w="3000"/><w:gridCol w:w="3000"/></w:tblGrid>
              <w:tr><w:tc><w:p><w:r><w:t>{{#Lines}}{{item}}</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>{{Amount:0.00}}{{/Lines}}</w:t></w:r></w:p></w:tc></w:tr>
            </w:tbl>
            <w:p><w:r><w:t>Summa {{Total:0.00}}{{^Paid}} – obetald{{/Paid}}</w:t></w:r></w:p>
            """;
        OrdDocument document = await TestFiles.OpenAsync(TestFiles.Docx(body));
        Invoice invoice = new()
        {
            Customer = "Anna AB",
            Due = new DateOnly(2026, 12, 31),
            Lines = [new InvoiceLine { Item = "Te", Amount = 25 }, new InvoiceLine { Item = "Kaffe", Amount = 30.5m }],
        };

        TemplateResult result = document.FillTemplate(invoice, new TemplateOptions { Culture = CultureInfo.InvariantCulture });

        Assert.Empty(result.Missing);
        Assert.Equal("Faktura till Anna AB, förfaller 2026-12-31\nTe\t25.00\nKaffe\t30.50\nSumma 55.50 – obetald", document.ToPlainText());
    }

    [Fact]
    public void A_record_gets_its_values_too()
    {
        ITemplateData data = new Receipt("K-17");

        Assert.True(data.TryGetValue("Number", out object? value));
        Assert.Equal("K-17", value);
        Assert.True(data.TryGetValue("number", out _));
        Assert.False(data.TryGetValue("Missing", out _));
    }
}
