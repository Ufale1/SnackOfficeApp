using Microsoft.Data.Sqlite;
using System.Data;

namespace SnackOfficeApp;

public static class AppDb
{
    private static string DbPath =>
        Path.Combine(AppContext.BaseDirectory, "SnackOffice.db");

    private static string ConnectionString =>
        new SqliteConnectionStringBuilder { DataSource = DbPath }.ToString();

    public static void Initialize()
    {
        using var con = new SqliteConnection(ConnectionString);
        con.Open();

        // Settings (so you can change company info without coding)
        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Settings(
            Key TEXT PRIMARY KEY,
            Value TEXT
        );
        """);

        EnsureSetting("CompanyName", "Your Company Name");
        EnsureSetting("CompanyAddress", "Your Company Address");
        EnsureSetting("CompanyPhone", "Phone: 0000-0000000");

        // Masters
        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Customers(
            Name TEXT PRIMARY KEY,
            Phone TEXT,
            Address TEXT
        );
        """);

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Vendors(
            Name TEXT PRIMARY KEY,
            Phone TEXT,
            Address TEXT
        );
        """);

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Products(
            Name TEXT PRIMARY KEY,
            Uom TEXT
        );
        """);

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS ExpenseHeads(
            Name TEXT PRIMARY KEY
        );
        """);

        // Transactions
        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS SalesLines(
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT NOT NULL,              -- yyyy-MM-dd
            InvoiceNo TEXT NOT NULL,
            Customer TEXT NOT NULL,
            Product TEXT NOT NULL,
            QtyDozen REAL NOT NULL DEFAULT 0,
            Rate REAL NOT NULL DEFAULT 0,
            Amount REAL NOT NULL DEFAULT 0,
            AddressHint TEXT,
            Remarks TEXT
        );
        """);

        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_SalesLines_Date ON SalesLines(Date);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_SalesLines_Customer ON SalesLines(Customer);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_SalesLines_Product ON SalesLines(Product);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_SalesLines_InvoiceNo ON SalesLines(InvoiceNo);");

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Receipts(
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT NOT NULL,
            ReceiptNo TEXT NOT NULL,
            Customer TEXT NOT NULL,
            Amount REAL NOT NULL DEFAULT 0,
            Mode TEXT,
            RefNo TEXT,
            Remarks TEXT
        );
        """);

        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_Receipts_Date ON Receipts(Date);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_Receipts_Customer ON Receipts(Customer);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_Receipts_ReceiptNo ON Receipts(ReceiptNo);");

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS VendorBills(
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT NOT NULL,
            BillNo TEXT NOT NULL,
            Vendor TEXT NOT NULL,
            Amount REAL NOT NULL DEFAULT 0,
            Remarks TEXT
        );
        """);

        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_VendorBills_Date ON VendorBills(Date);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_VendorBills_Vendor ON VendorBills(Vendor);");

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS VendorPayments(
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT NOT NULL,
            PaymentNo TEXT NOT NULL,
            Vendor TEXT NOT NULL,
            Amount REAL NOT NULL DEFAULT 0,
            Mode TEXT,
            RefNo TEXT,
            Remarks TEXT
        );
        """);

        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_VendorPayments_Date ON VendorPayments(Date);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_VendorPayments_Vendor ON VendorPayments(Vendor);");

        ExecNonQuery(con, """
        CREATE TABLE IF NOT EXISTS Expenses(
            Id INTEGER PRIMARY KEY AUTOINCREMENT,
            Date TEXT NOT NULL,
            VoucherNo TEXT NOT NULL,
            Head TEXT NOT NULL,
            Payee TEXT,
            Amount REAL NOT NULL DEFAULT 0,
            Mode TEXT,
            RefNo TEXT,
            Remarks TEXT
        );
        """);

        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_Expenses_Date ON Expenses(Date);");
        ExecNonQuery(con, "CREATE INDEX IF NOT EXISTS IX_Expenses_Head ON Expenses(Head);");

        // Seed basic expense heads if empty
        var count = ScalarLong(con, "SELECT COUNT(*) FROM ExpenseHeads;");
        if (count == 0)
        {
            string[] heads = { "Electricity", "Salaries", "Maintenance", "Fuel", "Office", "Misc" };
            foreach (var h in heads) InsertExpenseHead(h);
        }

        void EnsureSetting(string key, string value)
        {
            ExecNonQuery(con, """
            INSERT OR IGNORE INTO Settings(Key, Value)
            VALUES(@k, @v);
            """.Replace("@k", $"'{key.Replace("'", "''")}'")
               .Replace("@v", $"'{value.Replace("'", "''")}'"));
        }
    }

    private static void ExecNonQuery(SqliteConnection con, string sql)
    {
        using var cmd = con.CreateCommand();
        cmd.CommandText = sql;
        cmd.ExecuteNonQuery();
    }

    private static long ScalarLong(SqliteConnection con, string sql)
    {
        using var cmd = con.CreateCommand();
        cmd.CommandText = sql;
        return (long)(cmd.ExecuteScalar() ?? 0L);
    }

    public static DataTable Query(string sql, params (string key, object? value)[] parameters)
    {
        using var con = new SqliteConnection(ConnectionString);
        con.Open();
        using var cmd = con.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (key, value) in parameters)
            cmd.Parameters.AddWithValue(key, value ?? DBNull.Value);

        using var reader = cmd.ExecuteReader();
        var dt = new DataTable();
        dt.Load(reader);
        return dt;
    }

    public static int Execute(string sql, params (string key, object? value)[] parameters)
    {
        using var con = new SqliteConnection(ConnectionString);
        con.Open();
        using var cmd = con.CreateCommand();
        cmd.CommandText = sql;
        foreach (var (key, value) in parameters)
            cmd.Parameters.AddWithValue(key, value ?? DBNull.Value);

        return cmd.ExecuteNonQuery();
    }

    // Settings helpers
    public static string GetSetting(string key, string defaultValue = "")
    {
        var dt = Query("SELECT Value FROM Settings WHERE Key=@k;", ("@k", key));
        if (dt.Rows.Count == 0) return defaultValue;
        return dt.Rows[0]["Value"]?.ToString() ?? defaultValue;
    }

    public static void SetSetting(string key, string value)
    {
        Execute("""
        INSERT INTO Settings(Key, Value) VALUES(@k, @v)
        ON CONFLICT(Key) DO UPDATE SET Value=excluded.Value;
        """, ("@k", key), ("@v", value));
    }

    // Masters helpers
    public static List<string> GetCustomers() =>
        Query("SELECT Name FROM Customers ORDER BY Name;")
            .AsEnumerable().Select(r => r.Field<string>("Name")!).ToList();

    public static List<string> GetVendors() =>
        Query("SELECT Name FROM Vendors ORDER BY Name;")
            .AsEnumerable().Select(r => r.Field<string>("Name")!).ToList();

    public static List<string> GetProducts() =>
        Query("SELECT Name FROM Products ORDER BY Name;")
            .AsEnumerable().Select(r => r.Field<string>("Name")!).ToList();

    public static List<string> GetExpenseHeads() =>
        Query("SELECT Name FROM ExpenseHeads ORDER BY Name;")
            .AsEnumerable().Select(r => r.Field<string>("Name")!).ToList();

    public static List<string> GetInvoiceNos(int limit = 2000) =>
        Query($"""
            SELECT DISTINCT InvoiceNo
            FROM SalesLines
            ORDER BY InvoiceNo DESC
            LIMIT {limit};
        """).AsEnumerable().Select(r => r.Field<string>("InvoiceNo")!).ToList();

    public static void InsertCustomer(string name, string? phone, string? address)
    {
        Execute("""
        INSERT OR REPLACE INTO Customers(Name, Phone, Address)
        VALUES(@n, @p, @a);
        """, ("@n", name.Trim()), ("@p", phone), ("@a", address));
    }

    public static void InsertVendor(string name, string? phone, string? address)
    {
        Execute("""
        INSERT OR REPLACE INTO Vendors(Name, Phone, Address)
        VALUES(@n, @p, @a);
        """, ("@n", name.Trim()), ("@p", phone), ("@a", address));
    }

    public static void InsertProduct(string name, string? uom)
    {
        Execute("""
        INSERT OR REPLACE INTO Products(Name, Uom)
        VALUES(@n, @u);
        """, ("@n", name.Trim()), ("@u", uom));
    }

    public static void InsertExpenseHead(string name)
    {
        Execute("""
        INSERT OR IGNORE INTO ExpenseHeads(Name)
        VALUES(@n);
        """, ("@n", name.Trim()));
    }
}
