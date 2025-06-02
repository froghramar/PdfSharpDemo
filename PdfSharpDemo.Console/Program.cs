using PdfSharpDemo.Invoices.SimpleInvoice;

var data = new SimpleInvoiceData(
    "541223456",
    DateTime.UtcNow,
    DateTime.UtcNow,
    0.1m,
    new SimpleInvoiceData.IssuedToAddress("Richard Sanchez", "Thynk Unlimited", "123 Anywhere St., Any City"),
    new SimpleInvoiceData.PaymentAddress("Borcele Bank", "0123 4567 8901", "Adeline Palmerston"),
    [
        new SimpleInvoiceData.InvoiceItem("Brand consultation", 250, 1, 250),
        new SimpleInvoiceData.InvoiceItem("Logo design", 25, 2, 50),
        new SimpleInvoiceData.InvoiceItem("Website design", 33.5m, 1, 33.5m),
        new SimpleInvoiceData.InvoiceItem("Social media templates", 40, 2, 80),
        new SimpleInvoiceData.InvoiceItem("Brand photography", 10, 12, 120),
        new SimpleInvoiceData.InvoiceItem("Brand guide", 15, 1, 15),
    ]);
var bytes = SimpleInvoiceGenerator.GenerateReport(data);

var filePath = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "simple-invoice.pdf");
await File.WriteAllBytesAsync(filePath, bytes);
Console.WriteLine($"Invoice saved to {filePath}");