using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using PdfSharpCore.Pdf;

namespace PdfSharpDemo.Invoices.SimpleInvoice;

public static class SimpleInvoiceGenerator
{
    public static byte[] GenerateReport(SimpleInvoiceData data)
    {
        var document = new Document();

        var style = document.Styles[StyleNames.Normal];
        style.ParagraphFormat.SpaceBefore = Unit.FromPoint(3);
        style.ParagraphFormat.SpaceAfter = Unit.FromPoint(3);

        var section = document.AddSection();
        section.PageSetup.PageWidth = Unit.FromCentimeter(41);
        section.PageSetup.PageHeight = Unit.FromCentimeter(29.7);
        section.PageSetup.LeftMargin = Unit.FromCentimeter(1);
        section.PageSetup.TopMargin = Unit.FromCentimeter(1);
        section.PageSetup.BottomMargin = Unit.FromCentimeter(1);

        var pdfRenderer = new PdfDocumentRenderer
        {
            Document = document,
            PdfDocument = new PdfDocument
            {
                PageLayout = PdfPageLayout.SinglePage,
                ViewerPreferences =
                {
                    FitWindow = true
                }
            }
        };

        pdfRenderer.RenderDocument();

        using var stream = new MemoryStream();
        pdfRenderer.Save(stream, false);
        return stream.ToArray();
    }

    private static void AddBannerRow(Table table, string field1, string value1, string field2, string value2, ParagraphAlignment value2Alignment = ParagraphAlignment.Left)
    {
        var row = table.AddRow();
        if (field1 is not null)
        {
            row[0].AddParagraph(AddColon(field1));
            row[0].Format.Font.Bold = true;

            row[1].AddParagraph(value1 ?? string.Empty);
        }

        row[2].AddParagraph(AddColon(field2));
        row[2].Format.Font.Bold = true;

        row[3].AddParagraph(value2 ?? string.Empty);
        row[3].Format.Alignment = value2Alignment;
    }

    private static string AddColon(string value)
    {
        return $"{value}:";
    }

    private static void AddPartHeader(Table table, decimal? totalHours)
    {
        var row = table.AddRow();
        var pmText = totalHours.HasValue ? $"PM (Hours {totalHours:F})" : "PM";
        row[0].AddParagraph(pmText);

        row[1].AddParagraph("SMCS Codes");
        row[2].AddParagraph("Description");
        row[3].AddParagraph("Additional Notes");
        row[4].AddParagraph("Custom Notes");

        row[5].AddParagraph("Group Number");

        row[6].AddParagraph("Part Number");
        row[7].AddParagraph("Part Description");
        row[8].AddParagraph("Replacement %");
        row[9].AddParagraph("Quantity");
        row[10].AddParagraph("Notes");

        for (var i = 0; i < 11; i++)
        {
            row[i].Format.Font.Bold = true;
        }
    }

    private static void AddSpace(Section sec, Unit height)
    {
        var p = sec.AddParagraph();
        p.Format.LineSpacingRule = LineSpacingRule.Exactly;
        p.Format.LineSpacing = "0mm";
        p.Format.SpaceBefore = height;
    }
}