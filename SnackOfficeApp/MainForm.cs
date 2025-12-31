using System.Data;
using System.Text;

namespace SnackOfficeApp;

public class MainForm : Form
{
    private readonly TabControl _tabs = new() { Dock = DockStyle.Fill };

    // Reports grid
    private readonly DateTimePicker _reportDate = new() { Width = 150 };
    private readonly DataGridView _gridReport = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // Settings UI
    private TextBox _setCompany = new(), _setAddress = new(), _setPhone = new();

    // Masters UI
    private TextBox _custName = new(), _custPhone = new(), _custAddr = new();
    private DataGridView _gridCustomers = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    private TextBox _vendorName = new(), _vendorPhone = new(), _vendorAddr = new();
    private DataGridView _gridVendors = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    private TextBox _prodName = new(), _prodUom = new();
    private DataGridView _gridProducts = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // Sales UI
    private DateTimePicker _saleDate = new() { Width = 130 };
    private TextBox _invoiceNo = new() { Width = 120 };
    private ComboBox _saleCustomer = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private ComboBox _saleProduct = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private NumericUpDown _qtyDozen = new() { Width = 90, DecimalPlaces = 2, Maximum = 1000000 };
    private NumericUpDown _rate = new() { Width = 90, DecimalPlaces = 2, Maximum = 100000000 };
    private TextBox _addressHint = new() { Width = 300 };
    private TextBox _saleRemarks = new() { Width = 200 };
    private DataGridView _gridSales = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // Receipts UI
    private DateTimePicker _rcptDate = new() { Width = 130 };
    private TextBox _receiptNo = new() { Width = 120 };
    private ComboBox _rcptCustomer = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private NumericUpDown _rcptAmount = new() { Width = 120, DecimalPlaces = 2, Maximum = 1000000000 };
    private ComboBox _rcptMode = new() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
    private TextBox _rcptRef = new() { Width = 160 };
    private TextBox _rcptRemarks = new() { Width = 200 };
    private DataGridView _gridReceipts = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // Vendor bills/payments
    private DateTimePicker _vbDate = new() { Width = 130 };
    private TextBox _vbBillNo = new() { Width = 120 };
    private ComboBox _vbVendor = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private NumericUpDown _vbAmount = new() { Width = 120, DecimalPlaces = 2, Maximum = 1000000000 };
    private TextBox _vbRemarks = new() { Width = 240 };
    private DataGridView _gridVendorBills = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    private DateTimePicker _vpDate = new() { Width = 130 };
    private TextBox _vpPayNo = new() { Width = 120 };
    private ComboBox _vpVendor = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private NumericUpDown _vpAmount = new() { Width = 120, DecimalPlaces = 2, Maximum = 1000000000 };
    private ComboBox _vpMode = new() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
    private TextBox _vpRef = new() { Width = 160 };
    private TextBox _vpRemarks = new() { Width = 200 };
    private DataGridView _gridVendorPayments = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // Expenses
    private DateTimePicker _expDate = new() { Width = 130 };
    private TextBox _expVoucher = new() { Width = 120 };
    private ComboBox _expHead = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private TextBox _expPayee = new() { Width = 220 };
    private NumericUpDown _expAmount = new() { Width = 120, DecimalPlaces = 2, Maximum = 1000000000 };
    private ComboBox _expMode = new() { Width = 120, DropDownStyle = ComboBoxStyle.DropDownList };
    private TextBox _expRef = new() { Width = 160 };
    private TextBox _expRemarks = new() { Width = 200 };
    private DataGridView _gridExpenses = new() { Dock = DockStyle.Fill, ReadOnly = true, AllowUserToAddRows = false };

    // PDF controls
    private ComboBox _pdfInvoiceNo = new() { Width = 180, DropDownStyle = ComboBoxStyle.DropDown };
    private ComboBox _stmtCustomer = new() { Width = 220, DropDownStyle = ComboBoxStyle.DropDown };
    private DateTimePicker _stmtFrom = new() { Width = 130 };
    private DateTimePicker _stmtTo = new() { Width = 130 };

    public MainForm()
    {
        Text = "SnackOfficeApp (Debtors / Vendors / Expenses / Product Sales + PDFs)";
        Width = 1400;
        Height = 850;
        StartPosition = FormStartPosition.CenterScreen;

        Controls.Add(_tabs);

        BuildSettingsTab();
        BuildMastersTabs();
        BuildSalesTab();
        BuildReceiptsTab();
        BuildVendorsTabs();
        BuildExpensesTab();
        BuildReportsTab();

        LoadAllLists();
        RefreshAllGrids();
        LoadSettingsIntoUI();
    }

    private static string Iso(DateTime dt) => dt.ToString("yyyy-MM-dd");

    private void LoadAllLists()
    {
        var customers = AppDb.GetCustomers();
        SetupCombo(_saleCustomer, customers);
        SetupCombo(_rcptCustomer, customers);
        SetupCombo(_stmtCustomer, customers);

        var vendors = AppDb.GetVendors();
        SetupCombo(_vbVendor, vendors);
        SetupCombo(_vpVendor, vendors);

        var products = AppDb.GetProducts();
        SetupCombo(_saleProduct, products);

        var heads = AppDb.GetExpenseHeads();
        SetupCombo(_expHead, heads);

        var invoices = AppDb.GetInvoiceNos();
        SetupCombo(_pdfInvoiceNo, invoices);

        _rcptMode.Items.Clear();
        _rcptMode.Items.AddRange(["Cash", "Bank"]);
        _rcptMode.SelectedIndex = 0;

        _vpMode.Items.Clear();
        _vpMode.Items.AddRange(["Cash", "Bank"]);
        _vpMode.SelectedIndex = 0;

        _expMode.Items.Clear();
        _expMode.Items.AddRange(["Cash", "Bank"]);
        _expMode.SelectedIndex = 0;
    }

    private static void SetupCombo(ComboBox cb, List<string> values)
    {
        cb.Items.Clear();
        cb.Items.AddRange(values.Cast<object>().ToArray());
        cb.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
        cb.AutoCompleteSource = AutoCompleteSource.ListItems;
    }

    private void RefreshAllGrids()
    {
        _gridCustomers.DataSource = AppDb.Query("SELECT Name, Phone, Address FROM Customers ORDER BY Name;");
        _gridVendors.DataSource = AppDb.Query("SELECT Name, Phone, Address FROM Vendors ORDER BY Name;");
        _gridProducts.DataSource = AppDb.Query("SELECT Name, Uom FROM Products ORDER BY Name;");

        _gridSales.DataSource = AppDb.Query("""
            SELECT Id, Date, InvoiceNo, Customer, Product, QtyDozen, Rate, Amount, AddressHint, Remarks
            FROM SalesLines
            ORDER BY Date DESC, Id DESC
            LIMIT 500;
        """);

        _gridReceipts.DataSource = AppDb.Query("""
            SELECT Id, Date, ReceiptNo, Customer, Amount, Mode, RefNo, Remarks
            FROM Receipts
            ORDER BY Date DESC, Id DESC
            LIMIT 500;
        """);

        _gridVendorBills.DataSource = AppDb.Query("""
            SELECT Id, Date, BillNo, Vendor, Amount, Remarks
            FROM VendorBills
            ORDER BY Date DESC, Id DESC
            LIMIT 500;
        """);

        _gridVendorPayments.DataSource = AppDb.Query("""
            SELECT Id, Date, PaymentNo, Vendor, Amount, Mode, RefNo, Remarks
            FROM VendorPayments
            ORDER BY Date DESC, Id DESC
            LIMIT 500;
        """);

        _gridExpenses.DataSource = AppDb.Query("""
            SELECT Id, Date, VoucherNo, Head, Payee, Amount, Mode, RefNo, Remarks
            FROM Expenses
            ORDER BY Date DESC, Id DESC
            LIMIT 1000;
        """);
    }

    private void BuildSettingsTab()
    {
        var tab = new TabPage("Settings");

        var panel = new TableLayoutPanel
        {
            Dock = DockStyle.Top,
            Height = 200,
            Padding = new Padding(12),
            ColumnCount = 2,
            RowCount = 5
        };
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 140));
        panel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));

        _setCompany.Width = 600;
        _setAddress.Width = 600;
        _setPhone.Width = 600;

        panel.Controls.Add(new Label { Text = "Company Name", AutoSize = true }, 0, 0);
        panel.Controls.Add(_setCompany, 1, 0);

        panel.Controls.Add(new Label { Text = "Company Address", AutoSize = true }, 0, 1);
        panel.Controls.Add(_setAddress, 1, 1);

        panel.Controls.Add(new Label { Text = "Company Phone", AutoSize = true }, 0, 2);
        panel.Controls.Add(_setPhone, 1, 2);

        var btnSave = new Button { Text = "Save Settings", Width = 140 };
        btnSave.Click += (_, _) =>
        {
            AppDb.SetSetting("CompanyName", _setCompany.Text.Trim());
            AppDb.SetSetting("CompanyAddress", _setAddress.Text.Trim());
            AppDb.SetSetting("CompanyPhone", _setPhone.Text.Trim());
            MessageBox.Show("Saved. These will appear on Invoice/Statement PDFs.");
        };

        panel.Controls.Add(btnSave, 1, 3);

        tab.Controls.Add(panel);
        _tabs.TabPages.Add(tab);
    }

    private void LoadSettingsIntoUI()
    {
        _setCompany.Text = AppDb.GetSetting("CompanyName", "");
        _setAddress.Text = AppDb.GetSetting("CompanyAddress", "");
        _setPhone.Text = AppDb.GetSetting("CompanyPhone", "");
    }

    private void BuildMastersTabs()
    {
        _tabs.TabPages.Add(BuildCustomersTab());
        _tabs.TabPages.Add(BuildVendorsMasterTab());
        _tabs.TabPages.Add(BuildProductsTab());
    }

    private TabPage BuildCustomersTab()
    {
        var tab = new TabPage("Customers");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(8) };
        top.Controls.AddRange([
            new Label{ Text="Name", AutoSize=true }, _custName,
            new Label{ Text="Phone", AutoSize=true }, _custPhone,
            new Label{ Text="Address", AutoSize=true }, _custAddr
        ]);

        _custName.Width = 200;
        _custPhone.Width = 140;
        _custAddr.Width = 350;

        var btnAdd = new Button { Text = "Add/Update", Width = 120 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_custName.Text))
            {
                MessageBox.Show("Customer name is required."); return;
            }
            AppDb.InsertCustomer(_custName.Text, _custPhone.Text, _custAddr.Text);
            _custName.Clear(); _custPhone.Clear(); _custAddr.Clear();
            LoadAllLists();
            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridCustomers.CurrentRow?.Cells["Name"]?.Value is not string name) return;
            var ok = MessageBox.Show($"Delete customer '{name}'?", "Confirm", MessageBoxButtons.YesNo);
            if (ok != DialogResult.Yes) return;
            AppDb.Execute("DELETE FROM Customers WHERE Name=@n;", ("@n", name));
            LoadAllLists();
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridCustomers);
        tab.Controls.Add(top);
        return tab;
    }

    private TabPage BuildVendorsMasterTab()
    {
        var tab = new TabPage("Vendors");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(8) };
        _vendorName.Width = 220;
        _vendorPhone.Width = 140;
        _vendorAddr.Width = 350;

        top.Controls.AddRange([
            new Label{Text="Name", AutoSize=true}, _vendorName,
            new Label{Text="Phone", AutoSize=true}, _vendorPhone,
            new Label{Text="Address", AutoSize=true}, _vendorAddr
        ]);

        var btnAdd = new Button { Text = "Add/Update", Width = 120 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_vendorName.Text))
            {
                MessageBox.Show("Vendor name is required."); return;
            }
            AppDb.InsertVendor(_vendorName.Text, _vendorPhone.Text, _vendorAddr.Text);
            _vendorName.Clear(); _vendorPhone.Clear(); _vendorAddr.Clear();
            LoadAllLists();
            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridVendors.CurrentRow?.Cells["Name"]?.Value is not string name) return;
            var ok = MessageBox.Show($"Delete vendor '{name}'?", "Confirm", MessageBoxButtons.YesNo);
            if (ok != DialogResult.Yes) return;
            AppDb.Execute("DELETE FROM Vendors WHERE Name=@n;", ("@n", name));
            LoadAllLists();
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridVendors);
        tab.Controls.Add(top);
        return tab;
    }

    private TabPage BuildProductsTab()
    {
        var tab = new TabPage("Products");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 70, Padding = new Padding(8) };
        _prodName.Width = 260;
        _prodUom.Width = 120;

        top.Controls.AddRange([
            new Label{ Text="Product", AutoSize=true }, _prodName,
            new Label{ Text="UOM", AutoSize=true }, _prodUom
        ]);

        var btnAdd = new Button { Text = "Add/Update", Width = 120 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_prodName.Text))
            {
                MessageBox.Show("Product name is required."); return;
            }
            AppDb.InsertProduct(_prodName.Text, _prodUom.Text);
            _prodName.Clear(); _prodUom.Clear();
            LoadAllLists();
            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridProducts.CurrentRow?.Cells["Name"]?.Value is not string name) return;
            var ok = MessageBox.Show($"Delete product '{name}'?", "Confirm", MessageBoxButtons.YesNo);
            if (ok != DialogResult.Yes) return;
            AppDb.Execute("DELETE FROM Products WHERE Name=@n;", ("@n", name));
            LoadAllLists();
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridProducts);
        tab.Controls.Add(top);
        return tab;
    }

    private void BuildSalesTab()
    {
        var tab = new TabPage("Sales (Lines)");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoScroll = true };

        top.Controls.AddRange([
            new Label{Text="Date", AutoSize=true}, _saleDate,
            new Label{Text="Invoice#", AutoSize=true}, _invoiceNo,
            new Label{Text="Customer", AutoSize=true}, _saleCustomer,
            new Label{Text="Product", AutoSize=true}, _saleProduct,
            new Label{Text="Dozens", AutoSize=true}, _qtyDozen,
            new Label{Text="Rate", AutoSize=true}, _rate,
            new Label{Text="Address Hint", AutoSize=true}, _addressHint,
            new Label{Text="Remarks", AutoSize=true}, _saleRemarks
        ]);

        var btnAdd = new Button { Text = "Add Line", Width = 100 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_invoiceNo.Text)) { MessageBox.Show("Invoice# required."); return; }
            if (string.IsNullOrWhiteSpace(_saleCustomer.Text)) { MessageBox.Show("Customer required."); return; }
            if (string.IsNullOrWhiteSpace(_saleProduct.Text)) { MessageBox.Show("Product required."); return; }

            var qty = (decimal)_qtyDozen.Value;
            var rate = (decimal)_rate.Value;
            var amount = qty * rate;

            AppDb.Execute("""
                INSERT INTO SalesLines(Date, InvoiceNo, Customer, Product, QtyDozen, Rate, Amount, AddressHint, Remarks)
                VALUES(@d,@inv,@c,@p,@q,@r,@a,@h,@rmk);
            """,
            ("@d", Iso(_saleDate.Value)),
            ("@inv", _invoiceNo.Text.Trim()),
            ("@c", _saleCustomer.Text.Trim()),
            ("@p", _saleProduct.Text.Trim()),
            ("@q", (double)qty),
            ("@r", (double)rate),
            ("@a", (double)amount),
            ("@h", _addressHint.Text),
            ("@rmk", _saleRemarks.Text)
            );

            LoadAllLists();     // refresh invoice list also
            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridSales.CurrentRow?.Cells["Id"]?.Value == null) return;
            var id = Convert.ToInt64(_gridSales.CurrentRow.Cells["Id"].Value);
            AppDb.Execute("DELETE FROM SalesLines WHERE Id=@id;", ("@id", id));
            LoadAllLists();
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridSales);
        tab.Controls.Add(top);
        _tabs.TabPages.Add(tab);
    }

    private void BuildReceiptsTab()
    {
        var tab = new TabPage("Receipts");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoScroll = true };

        top.Controls.AddRange([
            new Label{Text="Date", AutoSize=true}, _rcptDate,
            new Label{Text="Receipt#", AutoSize=true}, _receiptNo,
            new Label{Text="Customer", AutoSize=true}, _rcptCustomer,
            new Label{Text="Amount", AutoSize=true}, _rcptAmount,
            new Label{Text="Mode", AutoSize=true}, _rcptMode,
            new Label{Text="Ref", AutoSize=true}, _rcptRef,
            new Label{Text="Remarks", AutoSize=true}, _rcptRemarks
        ]);

        var btnAdd = new Button { Text = "Add Receipt", Width = 110 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_receiptNo.Text)) { MessageBox.Show("Receipt# required."); return; }
            if (string.IsNullOrWhiteSpace(_rcptCustomer.Text)) { MessageBox.Show("Customer required."); return; }

            AppDb.Execute("""
                INSERT INTO Receipts(Date, ReceiptNo, Customer, Amount, Mode, RefNo, Remarks)
                VALUES(@d,@no,@c,@a,@m,@r,@rmk);
            """,
            ("@d", Iso(_rcptDate.Value)),
            ("@no", _receiptNo.Text.Trim()),
            ("@c", _rcptCustomer.Text.Trim()),
            ("@a", (double)_rcptAmount.Value),
            ("@m", _rcptMode.Text),
            ("@r", _rcptRef.Text),
            ("@rmk", _rcptRemarks.Text)
            );

            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridReceipts.CurrentRow?.Cells["Id"]?.Value == null) return;
            var id = Convert.ToInt64(_gridReceipts.CurrentRow.Cells["Id"].Value);
            AppDb.Execute("DELETE FROM Receipts WHERE Id=@id;", ("@id", id));
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridReceipts);
        tab.Controls.Add(top);
        _tabs.TabPages.Add(tab);
    }

    private void BuildVendorsTabs()
    {
        _tabs.TabPages.Add(BuildVendorBillsTab());
        _tabs.TabPages.Add(BuildVendorPaymentsTab());
    }

    private TabPage BuildVendorBillsTab()
    {
        var tab = new TabPage("Vendor Bills");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoScroll = true };
        top.Controls.AddRange([
            new Label{Text="Date", AutoSize=true}, _vbDate,
            new Label{Text="Bill#", AutoSize=true}, _vbBillNo,
            new Label{Text="Vendor", AutoSize=true}, _vbVendor,
            new Label{Text="Amount", AutoSize=true}, _vbAmount,
            new Label{Text="Remarks", AutoSize=true}, _vbRemarks
        ]);

        var btnAdd = new Button { Text = "Add Bill", Width = 100 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_vbBillNo.Text)) { MessageBox.Show("Bill# required."); return; }
            if (string.IsNullOrWhiteSpace(_vbVendor.Text)) { MessageBox.Show("Vendor required."); return; }

            AppDb.Execute("""
                INSERT INTO VendorBills(Date, BillNo, Vendor, Amount, Remarks)
                VALUES(@d,@no,@v,@a,@rmk);
            """,
            ("@d", Iso(_vbDate.Value)),
            ("@no", _vbBillNo.Text.Trim()),
            ("@v", _vbVendor.Text.Trim()),
            ("@a", (double)_vbAmount.Value),
            ("@rmk", _vbRemarks.Text)
            );

            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridVendorBills.CurrentRow?.Cells["Id"]?.Value == null) return;
            var id = Convert.ToInt64(_gridVendorBills.CurrentRow.Cells["Id"].Value);
            AppDb.Execute("DELETE FROM VendorBills WHERE Id=@id;", ("@id", id));
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridVendorBills);
        tab.Controls.Add(top);
        return tab;
    }

    private TabPage BuildVendorPaymentsTab()
    {
        var tab = new TabPage("Vendor Payments");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoScroll = true };
        top.Controls.AddRange([
            new Label{Text="Date", AutoSize=true}, _vpDate,
            new Label{Text="Payment#", AutoSize=true}, _vpPayNo,
            new Label{Text="Vendor", AutoSize=true}, _vpVendor,
            new Label{Text="Amount", AutoSize=true}, _vpAmount,
            new Label{Text="Mode", AutoSize=true}, _vpMode,
            new Label{Text="Ref", AutoSize=true}, _vpRef,
            new Label{Text="Remarks", AutoSize=true}, _vpRemarks
        ]);

        var btnAdd = new Button { Text = "Add Payment", Width = 110 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_vpPayNo.Text)) { MessageBox.Show("Payment# required."); return; }
            if (string.IsNullOrWhiteSpace(_vpVendor.Text)) { MessageBox.Show("Vendor required."); return; }

            AppDb.Execute("""
                INSERT INTO VendorPayments(Date, PaymentNo, Vendor, Amount, Mode, RefNo, Remarks)
                VALUES(@d,@no,@v,@a,@m,@r,@rmk);
            """,
            ("@d", Iso(_vpDate.Value)),
            ("@no", _vpPayNo.Text.Trim()),
            ("@v", _vpVendor.Text.Trim()),
            ("@a", (double)_vpAmount.Value),
            ("@m", _vpMode.Text),
            ("@r", _vpRef.Text),
            ("@rmk", _vpRemarks.Text)
            );

            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridVendorPayments.CurrentRow?.Cells["Id"]?.Value == null) return;
            var id = Convert.ToInt64(_gridVendorPayments.CurrentRow.Cells["Id"].Value);
            AppDb.Execute("DELETE FROM VendorPayments WHERE Id=@id;", ("@id", id));
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridVendorPayments);
        tab.Controls.Add(top);
        return tab;
    }

    private void BuildExpensesTab()
    {
        var tab = new TabPage("Expenses");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 80, Padding = new Padding(8), AutoScroll = true };
        top.Controls.AddRange([
            new Label{Text="Date", AutoSize=true}, _expDate,
            new Label{Text="Voucher#", AutoSize=true}, _expVoucher,
            new Label{Text="Head", AutoSize=true}, _expHead,
            new Label{Text="Payee", AutoSize=true}, _expPayee,
            new Label{Text="Amount", AutoSize=true}, _expAmount,
            new Label{Text="Mode", AutoSize=true}, _expMode,
            new Label{Text="Ref", AutoSize=true}, _expRef,
            new Label{Text="Remarks", AutoSize=true}, _expRemarks
        ]);

        var btnAdd = new Button { Text = "Add Expense", Width = 110 };
        btnAdd.Click += (_, _) =>
        {
            if (string.IsNullOrWhiteSpace(_expVoucher.Text)) { MessageBox.Show("Voucher# required."); return; }
            if (string.IsNullOrWhiteSpace(_expHead.Text)) { MessageBox.Show("Expense head required."); return; }

            AppDb.InsertExpenseHead(_expHead.Text.Trim());
            LoadAllLists();

            AppDb.Execute("""
                INSERT INTO Expenses(Date, VoucherNo, Head, Payee, Amount, Mode, RefNo, Remarks)
                VALUES(@d,@no,@h,@p,@a,@m,@r,@rmk);
            """,
            ("@d", Iso(_expDate.Value)),
            ("@no", _expVoucher.Text.Trim()),
            ("@h", _expHead.Text.Trim()),
            ("@p", _expPayee.Text),
            ("@a", (double)_expAmount.Value),
            ("@m", _expMode.Text),
            ("@r", _expRef.Text),
            ("@rmk", _expRemarks.Text)
            );

            RefreshAllGrids();
        };

        var btnDelete = new Button { Text = "Delete Selected", Width = 120 };
        btnDelete.Click += (_, _) =>
        {
            if (_gridExpenses.CurrentRow?.Cells["Id"]?.Value == null) return;
            var id = Convert.ToInt64(_gridExpenses.CurrentRow.Cells["Id"].Value);
            AppDb.Execute("DELETE FROM Expenses WHERE Id=@id;", ("@id", id));
            RefreshAllGrids();
        };

        top.Controls.Add(btnAdd);
        top.Controls.Add(btnDelete);

        tab.Controls.Add(_gridExpenses);
        tab.Controls.Add(top);
        _tabs.TabPages.Add(tab);
    }

    private void BuildReportsTab()
    {
        var tab = new TabPage("Reports + PDFs");

        var top = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 120, Padding = new Padding(8), AutoScroll = true };

        _reportDate.Value = DateTime.Today;
        _stmtFrom.Value = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
        _stmtTo.Value = DateTime.Today;

        top.Controls.Add(new Label { Text = "As of Date", AutoSize = true });
        top.Controls.Add(_reportDate);

        var btnDebtors = new Button { Text = "Debtors (MTD + Closing)", Width = 180 };
        btnDebtors.Click += (_, _) => RunDebtorsReport();

        var btnVendors = new Button { Text = "Vendors (MTD + Closing)", Width = 180 };
        btnVendors.Click += (_, _) => RunVendorsReport();

        var btnExpenses = new Button { Text = "Expenses MTD (Detail)", Width = 170 };
        btnExpenses.Click += (_, _) => RunExpensesMtd();

        var btnSalesMatrix = new Button { Text = "Sales Matrix MTD (Amount)", Width = 210 };
        btnSalesMatrix.Click += (_, _) => RunSalesMatrixMtd();

        var btnExport = new Button { Text = "Export Grid to CSV", Width = 140 };
        btnExport.Click += (_, _) => ExportGridToCsv(_gridReport);

        // PDF section
        var btnRefreshInvoices = new Button { Text = "Refresh Invoice List", Width = 150 };
        btnRefreshInvoices.Click += (_, _) => LoadAllLists();

        var btnInvoicePdf = new Button { Text = "Create Invoice PDF", Width = 150 };
        btnInvoicePdf.Click += (_, _) => CreateInvoicePdf();

        var btnStmtPdf = new Button { Text = "Create Statement PDF", Width = 160 };
        btnStmtPdf.Click += (_, _) => CreateStatementPdf();

        top.Controls.Add(btnDebtors);
        top.Controls.Add(btnVendors);
        top.Controls.Add(btnExpenses);
        top.Controls.Add(btnSalesMatrix);
        top.Controls.Add(btnExport);

        top.SetFlowBreak(btnExport, true);

        top.Controls.Add(new Label { Text = "Invoice No", AutoSize = true });
        top.Controls.Add(_pdfInvoiceNo);
        top.Controls.Add(btnRefreshInvoices);
        top.Controls.Add(btnInvoicePdf);

        top.SetFlowBreak(btnInvoicePdf, true);

        top.Controls.Add(new Label { Text = "Statement Customer", AutoSize = true });
        top.Controls.Add(_stmtCustomer);
        top.Controls.Add(new Label { Text = "From", AutoSize = true });
        top.Controls.Add(_stmtFrom);
        top.Controls.Add(new Label { Text = "To", AutoSize = true });
        top.Controls.Add(_stmtTo);
        top.Controls.Add(btnStmtPdf);

        tab.Controls.Add(_gridReport);
        tab.Controls.Add(top);
        _tabs.TabPages.Add(tab);
    }

    private void CreateInvoicePdf()
    {
        if (string.IsNullOrWhiteSpace(_pdfInvoiceNo.Text))
        {
            MessageBox.Show("Select/enter Invoice No first.");
            return;
        }

        using var sfd = new SaveFileDialog
        {
            Filter = "PDF (*.pdf)|*.pdf",
            FileName = $"Invoice_{_pdfInvoiceNo.Text.Trim()}.pdf"
        };
        if (sfd.ShowDialog() != DialogResult.OK) return;

        try
        {
            PdfService.GenerateInvoicePdf(_pdfInvoiceNo.Text.Trim(), sfd.FileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"PDF failed: {ex.Message}");
        }
    }

    private void CreateStatementPdf()
    {
        if (string.IsNullOrWhiteSpace(_stmtCustomer.Text))
        {
            MessageBox.Show("Select customer first.");
            return;
        }

        using var sfd = new SaveFileDialog
        {
            Filter = "PDF (*.pdf)|*.pdf",
            FileName = $"Statement_{_stmtCustomer.Text.Trim()}_{_stmtFrom.Value:yyyyMMdd}_{_stmtTo.Value:yyyyMMdd}.pdf"
        };
        if (sfd.ShowDialog() != DialogResult.OK) return;

        try
        {
            PdfService.GenerateCustomerStatementPdf(_stmtCustomer.Text.Trim(), _stmtFrom.Value.Date, _stmtTo.Value.Date, sfd.FileName);
        }
        catch (Exception ex)
        {
            MessageBox.Show($"PDF failed: {ex.Message}");
        }
    }

    private void RunDebtorsReport()
    {
        var asOf = _reportDate.Value.Date;
        var ms = new DateTime(asOf.Year, asOf.Month, 1);

        var dt = AppDb.Query("""
        WITH
        SalesBefore AS (
            SELECT Customer, SUM(Amount) AS Amt
            FROM SalesLines
            WHERE date(Date) < date(@ms)
            GROUP BY Customer
        ),
        RcptBefore AS (
            SELECT Customer, SUM(Amount) AS Amt
            FROM Receipts
            WHERE date(Date) < date(@ms)
            GROUP BY Customer
        ),
        SalesMtd AS (
            SELECT Customer, SUM(Amount) AS Amt, SUM(QtyDozen) AS Doz
            FROM SalesLines
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            GROUP BY Customer
        ),
        RcptMtd AS (
            SELECT Customer, SUM(Amount) AS Amt
            FROM Receipts
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            GROUP BY Customer
        )
        SELECT
            c.Name AS Customer,
            COALESCE(sb.Amt,0) - COALESCE(rb.Amt,0) AS Opening,
            COALESCE(sm.Amt,0) AS SalesMTD,
            COALESCE(sm.Doz,0) AS DozensMTD,
            COALESCE(rm.Amt,0) AS ReceiptsMTD,
            (COALESCE(sb.Amt,0) - COALESCE(rb.Amt,0)) + COALESCE(sm.Amt,0) - COALESCE(rm.Amt,0) AS Closing
        FROM Customers c
        LEFT JOIN SalesBefore sb ON sb.Customer = c.Name
        LEFT JOIN RcptBefore rb ON rb.Customer = c.Name
        LEFT JOIN SalesMtd sm ON sm.Customer = c.Name
        LEFT JOIN RcptMtd rm ON rm.Customer = c.Name
        ORDER BY c.Name;
        """, ("@ms", Iso(ms)), ("@asof", Iso(asOf)));

        _gridReport.DataSource = dt;
    }

    private void RunVendorsReport()
    {
        var asOf = _reportDate.Value.Date;
        var ms = new DateTime(asOf.Year, asOf.Month, 1);

        var dt = AppDb.Query("""
        WITH
        BillsBefore AS (
            SELECT Vendor, SUM(Amount) AS Amt
            FROM VendorBills
            WHERE date(Date) < date(@ms)
            GROUP BY Vendor
        ),
        PaidBefore AS (
            SELECT Vendor, SUM(Amount) AS Amt
            FROM VendorPayments
            WHERE date(Date) < date(@ms)
            GROUP BY Vendor
        ),
        BillsMtd AS (
            SELECT Vendor, SUM(Amount) AS Amt
            FROM VendorBills
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            GROUP BY Vendor
        ),
        PaidMtd AS (
            SELECT Vendor, SUM(Amount) AS Amt
            FROM VendorPayments
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            GROUP BY Vendor
        )
        SELECT
            v.Name AS Vendor,
            COALESCE(bb.Amt,0) - COALESCE(pb.Amt,0) AS Opening,
            COALESCE(bm.Amt,0) AS BillsMTD,
            COALESCE(pm.Amt,0) AS PaidMTD,
            (COALESCE(bb.Amt,0) - COALESCE(pb.Amt,0)) + COALESCE(bm.Amt,0) - COALESCE(pm.Amt,0) AS Closing
        FROM Vendors v
        LEFT JOIN BillsBefore bb ON bb.Vendor = v.Name
        LEFT JOIN PaidBefore pb ON pb.Vendor = v.Name
        LEFT JOIN BillsMtd bm ON bm.Vendor = v.Name
        LEFT JOIN PaidMtd pm ON pm.Vendor = v.Name
        ORDER BY v.Name;
        """, ("@ms", Iso(ms)), ("@asof", Iso(asOf)));

        _gridReport.DataSource = dt;
    }

    private void RunExpensesMtd()
    {
        var asOf = _reportDate.Value.Date;
        var ms = new DateTime(asOf.Year, asOf.Month, 1);

        var dt = AppDb.Query("""
            SELECT Date, VoucherNo, Head, Payee, Amount, Mode, RefNo, Remarks
            FROM Expenses
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            ORDER BY Date ASC, Id ASC;
        """, ("@ms", Iso(ms)), ("@asof", Iso(asOf)));

        _gridReport.DataSource = dt;
    }

    private void RunSalesMatrixMtd()
    {
        var asOf = _reportDate.Value.Date;
        var ms = new DateTime(asOf.Year, asOf.Month, 1);

        var raw = AppDb.Query("""
            SELECT Customer, Product, SUM(Amount) AS Amt
            FROM SalesLines
            WHERE date(Date) >= date(@ms) AND date(Date) <= date(@asof)
            GROUP BY Customer, Product
            ORDER BY Customer, Product;
        """, ("@ms", Iso(ms)), ("@asof", Iso(asOf)));

        var products = AppDb.GetProducts();
        var customers = raw.AsEnumerable()
            .Select(r => r.Field<string>("Customer")!)
            .Distinct()
            .OrderBy(x => x)
            .ToList();

        var dt = new DataTable();
        dt.Columns.Add("Customer", typeof(string));
        foreach (var p in products)
            dt.Columns.Add(p, typeof(double));
        dt.Columns.Add("Total", typeof(double));

        var dict = new Dictionary<(string c, string p), double>();
        foreach (DataRow r in raw.Rows)
        {
            var c = (string)r["Customer"];
            var p = (string)r["Product"];
            var a = Convert.ToDouble(r["Amt"]);
            dict[(c, p)] = a;
        }

        foreach (var c in customers)
        {
            var row = dt.NewRow();
            row["Customer"] = c;
            double total = 0;

            foreach (var p in products)
            {
                var val = dict.TryGetValue((c, p), out var a) ? a : 0;
                row[p] = val;
                total += val;
            }

            row["Total"] = total;
            dt.Rows.Add(row);
        }

        _gridReport.DataSource = dt;
    }

    private void ExportGridToCsv(DataGridView grid)
    {
        if (grid.DataSource is not DataTable dt || dt.Columns.Count == 0)
        {
            MessageBox.Show("Nothing to export."); return;
        }

        using var sfd = new SaveFileDialog
        {
            Filter = "CSV (*.csv)|*.csv",
            FileName = $"Export_{DateTime.Now:yyyyMMdd_HHmm}.csv"
        };

        if (sfd.ShowDialog() != DialogResult.OK) return;

        var sb = new StringBuilder();

        sb.AppendLine(string.Join(",", dt.Columns.Cast<DataColumn>().Select(c => Csv(c.ColumnName))));

        foreach (DataRow r in dt.Rows)
        {
            var cells = dt.Columns.Cast<DataColumn>().Select(c => Csv(r[c]?.ToString() ?? ""));
            sb.AppendLine(string.Join(",", cells));
        }

        File.WriteAllText(sfd.FileName, sb.ToString(), Encoding.UTF8);
        MessageBox.Show("Exported.");
    }

    private static string Csv(string s)
    {
        if (s.Contains('"') || s.Contains(',') || s.Contains('\n') || s.Contains('\r'))
            return "\"" + s.Replace("\"", "\"\"") + "\"";
        return s;
    }
}
