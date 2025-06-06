﻿using MigraDocCore.DocumentObjectModel;
using MigraDocCore.DocumentObjectModel.Shapes;
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
        
        var header = section.AddParagraph("INVOICE");
        header.Format.Font.Bold = true;
        
        var issuedToHeader = section.AddParagraph("ISSUED TO:");
        issuedToHeader.Format.Font.Bold = true;
        section.AddParagraph(data.IssuedTo.Name);
        section.AddParagraph(data.IssuedTo.AddressLine1);
        section.AddParagraph(data.IssuedTo.AddressLine2);
        
        AddSpace(section, Unit.FromPoint(10));
        var payToHeader = section.AddParagraph("PAY TO:");
        payToHeader.Format.Font.Bold = true;
        section.AddParagraph(data.PayTo.BankName);
        section.AddParagraph($"Account Name: {data.PayTo.AccountName}");
        section.AddParagraph($"Account No.: {data.PayTo.AccountNumber}");

        var invoiceDetailsFrame = section.AddTextFrame();
        var invoiceDetails = invoiceDetailsFrame.AddParagraph($"INVOICE NO: {data.InvoiceNumber}");
        invoiceDetails.Format.Font.Bold = true;
        invoiceDetails.Format.Alignment = ParagraphAlignment.Right;
        var date = invoiceDetailsFrame.AddParagraph($"DATE: {data.InvoiceDate:dd/MM/yyyy}");
        date.Format.Alignment = ParagraphAlignment.Right;
        var dueDate = invoiceDetailsFrame.AddParagraph($"DUE DATE: {data.DueDate:dd/MM/yyyy}");
        dueDate.Format.Alignment = ParagraphAlignment.Right;
        invoiceDetailsFrame.Height = "3.0cm";
        invoiceDetailsFrame.Width = "7.0cm";
        invoiceDetailsFrame.Left = ShapePosition.Right;
        invoiceDetailsFrame.RelativeHorizontal = RelativeHorizontal.Margin;
        invoiceDetailsFrame.RelativeVertical = RelativeVertical.Margin;
        
        AddSpace(section, Unit.FromPoint(20));
        var itemsTable = section.AddTable();
        itemsTable.AddColumn(Unit.FromPoint(200));
        itemsTable.AddColumn(Unit.FromPoint(100));
        itemsTable.AddColumn(Unit.FromPoint(100));
        itemsTable.AddColumn(Unit.FromPoint(100));
        
        var titleRow = itemsTable.AddRow();
        titleRow[0].AddParagraph("DESCRIPTION");
        
        titleRow[1].AddParagraph("UNIT PRICE");
        titleRow[1].Format.Alignment = ParagraphAlignment.Right;
        
        titleRow[2].AddParagraph("QTY");
        titleRow[2].Format.Alignment = ParagraphAlignment.Right;
        
        titleRow[3].AddParagraph("TOTAL");
        titleRow[3].Format.Alignment = ParagraphAlignment.Right;
        
        decimal subtotal = 0;
        foreach (var item in data.Items)
        {
            var itemRow = itemsTable.AddRow();
            itemRow[0].AddParagraph(item.Description);
            
            itemRow[1].AddParagraph(item.UnitPrice.ToString("C"));
            itemRow[1].Format.Alignment = ParagraphAlignment.Right;
            
            itemRow[2].AddParagraph(item.Quantity.ToString("F"));
            itemRow[2].Format.Alignment = ParagraphAlignment.Right;
            
            itemRow[3].AddParagraph(item.Total.ToString("C"));
            itemRow[3].Format.Alignment = ParagraphAlignment.Right;
            
            subtotal += item.Total;
        }
        
        // Add empty row for spacing
        var spacingRow = itemsTable.AddRow();
        spacingRow[0].AddParagraph("");
        
        // Add subtotal row
        var subtotalRow = itemsTable.AddRow();
        subtotalRow[0].AddParagraph("");
        subtotalRow[1].AddParagraph("");
        subtotalRow[2].AddParagraph("SUBTOTAL:");
        subtotalRow[2].Format.Alignment = ParagraphAlignment.Right;
        subtotalRow[2].Format.Font.Bold = true;
        subtotalRow[3].AddParagraph(subtotal.ToString("C"));
        subtotalRow[3].Format.Alignment = ParagraphAlignment.Right;
        subtotalRow[3].Format.Font.Bold = true;
        
        // Calculate tax
        decimal taxAmount = subtotal * data.TaxRate;
        
        // Add tax row
        var taxRow = itemsTable.AddRow();
        taxRow[0].AddParagraph("");
        taxRow[1].AddParagraph("");
        taxRow[2].AddParagraph($"TAX ({data.TaxRate:P1}):");
        taxRow[2].Format.Alignment = ParagraphAlignment.Right;
        taxRow[2].Format.Font.Bold = true;
        taxRow[3].AddParagraph(taxAmount.ToString("C"));
        taxRow[3].Format.Alignment = ParagraphAlignment.Right;
        taxRow[3].Format.Font.Bold = true;
        
        // Calculate total
        decimal total = subtotal + taxAmount;
        
        // Add total row with border
        var totalRow = itemsTable.AddRow();
        totalRow[0].AddParagraph("");
        totalRow[1].AddParagraph("");
        totalRow[2].AddParagraph("TOTAL:");
        totalRow[2].Format.Alignment = ParagraphAlignment.Right;
        totalRow[2].Format.Font.Bold = true;
        totalRow[2].Format.Font.Size = Unit.FromPoint(12);
        totalRow[3].AddParagraph(total.ToString("C"));
        totalRow[3].Format.Alignment = ParagraphAlignment.Right;
        totalRow[3].Format.Font.Bold = true;
        totalRow[3].Format.Font.Size = Unit.FromPoint(12);
        
        // Add top border to total row
        totalRow[2].Borders.Top.Width = Unit.FromPoint(1);
        totalRow[2].Borders.Top.Color = Colors.Black;
        totalRow[3].Borders.Top.Width = Unit.FromPoint(1);
        totalRow[3].Borders.Top.Color = Colors.Black;
        
        var pdfDocument = new PdfDocument
        {
            PageLayout = PdfPageLayout.SinglePage,
            ViewerPreferences =
            {
                FitWindow = true
            }
        };

        var pdfRenderer = new PdfDocumentRenderer
        {
            Document = document,
            PdfDocument = pdfDocument
        };

        pdfRenderer.RenderDocument();
        
        AddWatermark(pdfDocument, "PAID");

        using var stream = new MemoryStream();
        pdfRenderer.Save(stream, false);
        return stream.ToArray();
    }

    private static void AddWatermark(PdfDocument document, string watermarkText)
    {
        foreach (var page in document.Pages)
        {
            var gfx = XGraphics.FromPdfPage(page, XGraphicsPdfPageOptions.Prepend);
            
            // Set up watermark font and brush
            var font = new XFont("Arial", 60, XFontStyle.Bold);
            XBrush brush = new XSolidBrush(XColor.FromArgb(50, 128, 128, 128)); // Semi-transparent gray
            
            // Calculate center position
            var textSize = gfx.MeasureString(watermarkText, font);
            
            // Save graphics state
            var state = gfx.Save();
            
            // Rotate and draw watermark
            gfx.TranslateTransform(page.Width / 2, page.Height / 2);
            gfx.RotateTransform(-45); // 45-degree rotation
            gfx.TranslateTransform(-textSize.Width / 2, -textSize.Height / 2);
            
            gfx.DrawString(watermarkText, font, brush, 0, 0);
            
            // Restore graphics state
            gfx.Restore(state);
            gfx.Dispose();
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