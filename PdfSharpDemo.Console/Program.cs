using PdfSharpDemo.Invoices.SimpleInvoice;

var data = new SimpleInvoiceData(
    "541223456",
    DateTime.UtcNow,
    DateTime.UtcNow,
    new SimpleInvoiceData.IssuedToAddress("Richard Sanchez", "Thynk Unlimited", "123 Anywhere St., Any City"),
    new SimpleInvoiceData.PaymentAddress("Borcele Bank", "0123 4567 8901", "Adeline Palmerston"));
var bytes = SimpleInvoiceGenerator.GenerateReport(data);

var filePath = Path.Join(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "simple-invoice.pdf");
await File.WriteAllBytesAsync(filePath, bytes);
Console.WriteLine($"Invoice saved to {filePath}");