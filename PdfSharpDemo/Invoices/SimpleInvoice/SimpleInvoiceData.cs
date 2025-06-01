namespace PdfSharpDemo.Invoices.SimpleInvoice;

public class SimpleInvoiceData(
    string invoiceNumber,
    DateTime invoiceDate,
    DateTime dueDate,
    SimpleInvoiceData.IssuedToAddress issuedTo,
    SimpleInvoiceData.PaymentAddress payTo,
    SimpleInvoiceData.InvoiceItem[] items)
{
    public string InvoiceNumber { get; set; } = invoiceNumber;
    public DateTime InvoiceDate { get; set; } = invoiceDate;
    public DateTime DueDate { get; set; } = dueDate;
    public IssuedToAddress IssuedTo { get; set; } = issuedTo;
    public PaymentAddress PayTo { get; set; } = payTo;
    public InvoiceItem[] Items { get; set; } = items;

    public class IssuedToAddress(string name, string addressLine1, string? addressLine2)
    {
        public string Name { get; set; } = name;
        public string AddressLine1 { get; set; } = addressLine1;
        public string? AddressLine2 { get; set; } = addressLine2;
    }

    public class PaymentAddress(string bankName, string accountNumber, string accountName)
    {
        public string BankName { get; set; } = bankName;
        public string AccountNumber { get; set; } = accountNumber;
        public string AccountName { get; set; } = accountName;
    }
    
    public class InvoiceItem(string description, decimal unitPrice, decimal quantity, decimal total)
    {
        public string Description { get; set; } = description;
        public decimal UnitPrice { get; set; } = unitPrice;
        public decimal Quantity { get; set; } = quantity;
        public decimal Total { get; set; } = total;
    }
}