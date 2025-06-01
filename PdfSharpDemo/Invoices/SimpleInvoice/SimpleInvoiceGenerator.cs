using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Shapes;
using MigraDocCore.DocumentObjectModel.Tables;
using MigraDocCore.Rendering;
using PdfSharpCore.Drawing;
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
        
        var header = section.AddParagraph("INVOICE");
        header.Format.Font.Bold = true;
        
        var issuedToHeader = section.AddParagraph("ISSUED TO:");
        issuedToHeader.Format.Font.Bold = true;
        section.AddParagraph(data.IssuedTo.Name);
        section.AddParagraph(data.IssuedTo.AddressLine1);
        section.AddParagraph(data.IssuedTo.AddressLine2);
        
        var payToHeader = section.AddParagraph("PAY TO:");
        payToHeader.Format.Font.Bold = true;
        section.AddParagraph(data.PayTo.BankName);
        section.AddParagraph($"Account Name: {data.PayTo.AccountName}");
        section.AddParagraph($"Account No.: {data.PayTo.AccountNumber}");

        var invoiceDetailsFrame = section.AddTextFrame();
        var invoiceDetails = invoiceDetailsFrame.AddParagraph($"INVOICE NO: {data.InvoiceNumber}");
        invoiceDetails.Format.Font.Bold = true;
        invoiceDetailsFrame.AddParagraph($"DATE: {data.InvoiceDate:dd/MM/yyyy}");
        invoiceDetailsFrame.AddParagraph($"DUE DATE: {data.DueDate:dd/MM/yyyy}");
        invoiceDetailsFrame.Height = "3.0cm";
        invoiceDetailsFrame.Width = "7.0cm";
        invoiceDetailsFrame.Left = ShapePosition.Right;
        invoiceDetailsFrame.RelativeHorizontal = RelativeHorizontal.Margin;
        invoiceDetailsFrame.Top = "0cm";
        invoiceDetailsFrame.RelativeVertical = RelativeVertical.Page;

        var pdfDocument = new PdfDocument
        {
            PageLayout = PdfPageLayout.SinglePage,
            ViewerPreferences =
            {
                FitWindow = true
            }
        };
        // var page = pdfDocument.Pages[0];
        // page.Width = XUnit.FromMillimeter(210);
        // page.Height = XUnit.FromMillimeter(297);
        // var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Append);
        // var lineRed = new XPen(XColors.Red, 5);
        // gfx.DrawLine(lineRed, 0, page.Height / 2, page.Width, page.Height / 2);

        var pdfRenderer = new PdfDocumentRenderer
        {
            Document = document,
            PdfDocument = pdfDocument
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