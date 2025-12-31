using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System.Data;
using System.Diagnostics;

namespace SnackOfficeApp;

public static class PdfService
{
    private static string Money(double v) => v.ToString("N2");
    private static string Iso(DateTime dt) => dt.ToString("yyyy-MM-dd");

    public static void GenerateInvoicePdf(string invoiceNo, string filePath)
    {
        if (string.IsNullOrWhiteSpace(invoiceNo))
            throw new ArgumentException("InvoiceNo is required.");

        var lines = AppDb.Query("""
            SELECT Date, InvoiceNo, Customer, Product, QtyDozen, Rate, Amount, AddressHint, Remarks
            FROM SalesLines
            WHERE InvoiceNo=@inv
            ORDER BY Id ASC;
        """, ("@inv", invoiceNo.Trim()));

        if (lines.Rows.Count == 0)
            throw new InvalidOperationException("Invoice not found.");

        var first = lines.Rows[0];
        var date = first["Date"].ToString() ?? "";
        var customer = first["Customer"].ToString() ?? "";
        var addressHint = first["AddressHint"]?.ToString() ?? "";
        var remarks = first["Remarks"]?.ToString() ?? "";

        double totalDoz = 0;
        double totalAmt = 0;
        foreach (DataRow r in lines.Rows)
        {
            totalDoz += Convert.ToDouble(r["QtyDozen"]);
            totalAmt += Convert.ToDouble(r["Amount"]);
        }

        var companyName = AppDb.GetSetting("CompanyName", "Company");
        var companyAddress = AppDb.GetSetting("CompanyAddress", "");
        var companyPhone = AppDb.GetSetting("CompanyPhone", "");

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(10));

                page.Header().Column(col =>
                {
                    col.Item().Text(companyName).SemiBold().FontSize(16);
                    if (!string.IsNullOrWhiteSpace(companyAddress))
                        col.Item().Text(companyAddress);
                    if (!string.IsNullOrWhiteSpace(companyPhone))
                        col.Item().Text(companyPhone);
                    col.Item().PaddingTop(8).LineHorizontal(1);
                });

                page.Content().Column(col =>
                {
                    col.Item().Row(row =>
                    {
                        row.RelativeItem().Column(c =>
                        {
                            c.Item().Text($"Customer: {customer}").SemiBold();
                            if (!string.IsNullOrWhiteSpace(addressHint))
                                c.Item().Text($"Address Hint: {addressHint}");
                            if (!string.IsNullOrWhiteSpace(remarks))
                                c.Item().Text($"Remarks: {remarks}");
                        });

                        row.ConstantItem(220).Column(c =>
                        {
                            c.Item().Text($"Invoice No: {invoiceNo}").SemiBold();
                            c.Item().Text($"Date: {date}");
                            c.Item().Text($"Total Dozens: {totalDoz:N2}");
                        });
                    });

                    col.Item().PaddingTop(10);

                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.RelativeColumn(6); // Product
                            columns.RelativeColumn(2); // Dozens
                            columns.RelativeColumn(2); // Rate
                            columns.RelativeColumn(2); // Amount
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(CellStyleHeader).Text("Product");
                            header.Cell().Element(CellStyleHeader).AlignRight().Text("Dozens");
                            header.Cell().Element(CellStyleHeader).AlignRight().Text("Rate");
                            header.Cell().Element(CellStyleHeader).AlignRight().Text("Amount");
                        });

                        foreach (DataRow r in lines.Rows)
                        {
                            var product = r["Product"]?.ToString() ?? "";
                            var doz = Convert.ToDouble(r["QtyDozen"]);
                            var rate = Convert.ToDouble(r["Rate"]);
                            var amt = Convert.ToDouble(r["Amount"]);

                            table.Cell().Element(CellStyle).Text(product);
                            table.Cell().Element(CellStyle).AlignRight().Text($"{doz:N2}");
                            table.Cell().Element(CellStyle).AlignRight().Text(Money(rate));
                            table.Cell().Element(CellStyle).AlignRight().Text(Money(amt));
                        }

                        table.Cell().ColumnSpan(4).PaddingTop(6).LineHorizontal(1);

                        table.Cell().ColumnSpan(2).Element(CellStyle).AlignRight().Text("TOTAL").SemiBold();
                        table.Cell().Element(CellStyle).AlignRight().Text("");
                        table.Cell().Element(CellStyle).AlignRight().Text(Money(totalAmt)).SemiBold();
                    });

                    col.Item().PaddingTop(10).Text($"Grand Total Amount: {Money(totalAmt)}").SemiBold().FontSize(12);
                });

                page.Footer().AlignCenter().Text(text =>
                {
                    text.Span("Page ");
                    text.CurrentPageNumber();
                    text.Span(" / ");
                    text.TotalPages();
                });

                static IContainer CellStyle(IContainer c) =>
                    c.BorderBottom(1).BorderColor(Colors.Grey.Lighten2).PaddingVertical(4).PaddingHorizontal(2);

                static IContainer CellStyleHeader(IContainer c) =>
                    c.Background(Colors.Grey.Lighten3).PaddingVertical(6).PaddingHorizontal(2).BorderBottom(1).BorderColor(Colors.Grey.Darken1);
            });
        }).GeneratePdf(filePath);

        TryOpen(filePath);
    }

    public static void GenerateCustomerStatementPdf(string customerName, DateTime fromDate, DateTime toDate, string filePath)
    {
        if (string.IsNullOrWhiteSpace(customerName))
            throw new ArgumentException("Customer is required.");
        if (toDate.Date < fromDate.Date)
            throw new ArgumentException("ToDate must be >= FromDate.");

        string cust = customerName.Trim();
        string fromIso = Iso(fromDate.Date);
        string toIso = Iso(toDate.Date);

        // Opening = Sales before - Receipts before
        var openingSalesDt = AppDb.Query("""
            SELECT COALESCE(SUM(Amount),0) AS Amt
            FROM SalesLines
            WHERE Customer=@c AND date(Date) < date(@from);
        """, ("@c", cust), ("@from", fromIso));

        var openingRcptDt = AppDb.Query("""
            SELECT COALESCE(SUM(Amount),0) AS Amt
            FROM Receipts
            WHERE Customer=@c AND date(Date) < date(@from);
        """, ("@c", cust), ("@from", fromIso));

        double openingSales = Convert.ToDouble(openingSalesDt.Rows[0]["Amt"]);
        double openingRcpt = Convert.ToDouble(openingRcptDt.Rows[0]["Amt"]);
        double opening = openingSales - openingRcpt;

        // Invoice summary within range (group lines into invoice totals + address hint)
        var inv = AppDb.Query("""
            SELECT
                Date,
                InvoiceNo,
                MAX(COALESCE(AddressHint,'')) AS AddressHint,
                SUM(QtyDozen) AS Dozens,
                SUM(Amount) AS Amount
            FROM SalesLines
            WHERE Customer=@c AND date(Date) >= date(@from) AND date(Date) <= date(@to)
            GROUP BY Date, InvoiceNo
            ORDER BY date(Date) ASC, InvoiceNo ASC;
        """, ("@c", cust), ("@from", fromIso), ("@to", toIso));

        var rcpt = AppDb.Query("""
            SELECT Date, ReceiptNo, Amount, Mode, RefNo, Remarks
            FROM Receipts
            WHERE Customer=@c AND date(Date) >= date(@from) AND date(Date) <= date(@to)
            ORDER BY date(Date) ASC, ReceiptNo ASC;
        """, ("@c", cust), ("@from", fromIso), ("@to", toIso));

        // Combine into statement lines
        var lines = new List<StmtLine>();

        foreach (DataRow r in inv.Rows)
        {
            lines.Add(new StmtLine
            {
                Date = r["Date"]?.ToString() ?? "",
                DocNo = r["InvoiceNo"]?.ToString() ?? "",
                Type = "Invoice",
                Description = string.IsNullOrWhiteSpace(r["AddressHint"]?.ToString())
                    ? ""
                    : $"Address: {r["AddressHint"]}",
                Debit = Convert.ToDouble(r["Amount"]),
                Credit = 0
            });
        }

        foreach (DataRow r in rcpt.Rows)
        {
            var mode = r["Mode"]?.ToString() ?? "";
            var refNo = r["RefNo"]?.ToString() ?? "";
            var remarks = r["Remarks"]?.ToString() ?? "";
            var descParts = new List<string>();
            if (!string.IsNullOrWhiteSpace(mode)) descParts.Add(mode);
            if (!string.IsNullOrWhiteSpace(refNo)) descParts.Add($"Ref: {refNo}");
            if (!string.IsNullOrWhiteSpace(remarks)) descParts.Add(remarks);

            lines.Add(new StmtLine
            {
                Date = r["Date"]?.ToString() ?? "",
                DocNo = r["ReceiptNo"]?.ToString() ?? "",
                Type = "Receipt",
                Description = string.Join(" | ", descParts),
                Debit = 0,
                Credit = Convert.ToDouble(r["Amount"])
            });
        }

        // Sort by date then type/docno
        lines = lines
            .OrderBy(l => l.Date)
            .ThenBy(l => l.Type)
            .ThenBy(l => l.DocNo)
            .ToList();

        // Running balance
        double balance = opening;
        foreach (var l in lines)
        {
            balance += l.Debit;
            balance -= l.Credit;
            l.Balance = balance;
        }

        var companyName = AppDb.GetSetting("CompanyName", "Company");
        var companyAddress = AppDb.GetSetting("CompanyAddress", "");
        var companyPhone = AppDb.GetSetting("CompanyPhone", "");

        Document.Create(container =>
        {
            container.Page(page =>
            {
                page.Size(PageSizes.A4);
                page.Margin(20);
                page.DefaultTextStyle(x => x.FontSize(10));

                page.Header().Column(col =>
                {
                    col.Item().Text(companyName).SemiBold().FontSize(16);
                    if (!string.IsNullOrWhiteSpace(companyAddress))
                        col.Item().Text(companyAddress);
                    if (!string.IsNullOrWhiteSpace(companyPhone))
                        col.Item().Text(companyPhone);
                    col.Item().PaddingTop(8).LineHorizontal(1);
                });

                page.Content().Column(col =>
                {
                    col.Item().Text("Customer Statement").SemiBold().FontSize(14);
                    col.Item().Text($"Customer: {cust}");
                    col.Item().Text($"Period: {fromDate:yyyy-MM-dd} to {toDate:yyyy-MM-dd}");
                    col.Item().Text($"Opening Balance: {Money(opening)}").SemiBold();
                    col.Item().PaddingTop(8);

                    col.Item().Table(table =>
                    {
                        table.ColumnsDefinition(columns =>
                        {
                            columns.ConstantColumn(80);  // Date
                            columns.ConstantColumn(90);  // Doc
                            columns.ConstantColumn(70);  // Type
                            columns.RelativeColumn(5);   // Desc
                            columns.RelativeColumn(2);   // Debit
                            columns.RelativeColumn(2);   // Credit
                            columns.RelativeColumn(2);   // Balance
                        });

                        table.Header(header =>
                        {
                            header.Cell().Element(H).Text("Date");
                            header.Cell().Element(H).Text("Doc No");
                            header.Cell().Element(H).Text("Type");
                            header.Cell().Element(H).Text("Description");
                            header.Cell().Element(H).AlignRight().Text("Debit");
                            header.Cell().Element(H).AlignRight().Text("Credit");
                            header.Cell().Element(H).AlignRight().Text("Balance");
                        });

                        foreach (var l in lines)
                        {
                            table.Cell().Element(C).Text(l.Date);
                            table.Cell().Element(C).Text(l.DocNo);
                            table.Cell().Element(C).Text(l.Type);
                            table.Cell().Element(C).Text(l.Description);
                            table.Cell().Element(C).AlignRight().Text(l.Debit == 0 ? "" : Money(l.Debit));
                            table.Cell().Element(C).AlignRight().Text(l.Credit == 0 ? "" : Money(l.Credit));
                            table.Cell().Element(C).AlignRight().Text(Money(l.Balance));
                        }

                        table.Cell().ColumnSpan(7).PaddingTop(6).LineHorizontal(1);

                        double totalDebit = lines.Sum(x => x.Debit);
                        double totalCredit = lines.Sum(x => x.Credit);

                        table.Cell().ColumnSpan(4).Element(C).AlignRight().Text("TOTALS").SemiBold();
                        table.Cell().Element(C).AlignRight().Text(Money(totalDebit)).SemiBold();
                        table.Cell().Element(C).AlignRight().Text(Money(totalCredit)).SemiBold();
                        table.Cell().Element(C).AlignRight().Text(Money(balance)).SemiBold();

                        static IContainer H(IContainer c) =>
                            c.Background(Colors.Grey.Lighten3).PaddingVertical(6).PaddingHorizontal(2).BorderBottom(1).BorderColor(Colors.Grey.Darken1);

                        static IContainer C(IContainer c) =>
                            c.BorderBottom(1).BorderColor(Colors.Grey.Lighten2).PaddingVertical(4).PaddingHorizontal(2);
                    });
                });

                page.Footer().AlignCenter().Text(text =>
                {
                    text.Span("Page ");
                    text.CurrentPageNumber();
                    text.Span(" / ");
                    text.TotalPages();
                });
            });
        }).GeneratePdf(filePath);

        TryOpen(filePath);
    }

    private static void TryOpen(string filePath)
    {
        try
        {
            Process.Start(new ProcessStartInfo(filePath) { UseShellExecute = true });
        }
        catch
        {
            // ignore if system blocks auto-open
        }
    }

    private class StmtLine
    {
        public string Date { get; set; } = "";
        public string DocNo { get; set; } = "";
        public string Type { get; set; } = "";
        public string Description { get; set; } = "";
        public double Debit { get; set; }
        public double Credit { get; set; }
        public double Balance { get; set; }
    }
}
