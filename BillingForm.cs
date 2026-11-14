using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.Globalization;
using System.Reflection;
using System.Windows.Forms;

namespace BarberShopApp
{
    public partial class  BillingForm : Form
    {
        private int _transactionId;
        private string _customerName, _customerPhone, _barberName, _servicesBooked;
        private decimal _baseSubtotal, _discountAmount, _tipAmount, _totalAmountDue, _amountTendered, _changeDue;
        private bool _isPaymentCompleted = false;

        // Left Column Controls
        private Panel cardInput;
        private ComboBox cmbSelectAppointment, cmbDiscount, cmbPaymentMethod;
        private Label lblSelectAppointment, lblCustomerNameDisplay, lblBarberTag, lblPhoneTag, lblServicesDisplay;
        private Label lblSubtotalValue, lblTotalDueValue, lblChangeValue, lblRefNo;
        private TextBox txtCustomDiscount, txtTip, txtReferenceNo, txtTendered;
        private Button btnConfirmPayment, btnPrintReceipt, btnClose, btnToggleHistory;

        // Right Column Controls
        private Panel cardRight, pnlReceiptView, pnlHistoryView;
        private RichTextBox rtbReceiptPreview;
        private DataGridView dgvPaidHistory;
        private Label lblRightHeader, lblHistoryTotalRevenue;
        private Button btnBackToHistory;
        private PrintDocument printDoc;

        public BillingForm() : this(0, "Walk-in Guest", "Walk-in", "All-Around Staff", "Standard Haircut", 250m) { }

        public BillingForm(int transactionId, string customerName, string customerPhone, string barberName, string servicesBooked, decimal baseSubtotal)
        {
            InitializeComponent();
            _transactionId = transactionId;
            _customerName = string.IsNullOrWhiteSpace(customerName) ? "Walk-in Customer" : customerName;
            _customerPhone = string.IsNullOrWhiteSpace(customerPhone) ? "Walk-in" : customerPhone;
            _barberName = string.IsNullOrWhiteSpace(barberName) ? "All-Around Staff" : barberName;
            _servicesBooked = string.IsNullOrWhiteSpace(servicesBooked) ? "Standard Haircut" : servicesBooked;
            _baseSubtotal = baseSubtotal;

            InitializeCustomDesign();
            LoadPendingAppointments();
            UpdateClientDisplay();
            CalculateBill();
            ResetButtonStates();
        }

        protected override void OnShown(EventArgs e)
        {
            base.OnShown(e);
            if (cardInput != null)
            {
                cardInput.AutoScrollPosition = new Point(0, 0);
            }
        }

        private void InitializeCustomDesign()
        {
            this.Text = "Checkout & Billing — The Classic Cut";
            this.Size = new Size(1160, 840);
            this.StartPosition = FormStartPosition.CenterScreen;
            this.BackColor = UIStyle.BgMain;
            this.Font = UIStyle.InputFont;

            TableLayoutPanel mainLayout = new TableLayoutPanel { Dock = DockStyle.Fill, ColumnCount = 2, RowCount = 1, Padding = new Padding(16) };
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 54F));
            mainLayout.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 46F));

            // ==========================================
            // LEFT COLUMN: Payment Inputs Card
            // ==========================================
            cardInput = UIStyle.CreateCard();
            cardInput.Dock = DockStyle.Fill;
            cardInput.AutoScroll = true;

            const int X = 22;
            const int W = 540;
            int y = 16;

            // Header line: Title on Left, Paid button ending exactly in line with W
            Label lblTitle = new Label { Text = "Checkout & Payment", UseMnemonic = false, Font = new Font("Segoe UI", 16F, FontStyle.Bold), ForeColor = UIStyle.TextDark, Location = new Point(X, y + 2), AutoSize = true };

            btnToggleHistory = new Button
            {
                Text = "📜 Paid",
                Location = new Point(X + W - 120, y),
                Size = new Size(120, 34),
                Font = new Font("Segoe UI", 9.5F, FontStyle.Bold),
                Cursor = Cursors.Hand
            };
            UIStyle.StyleButton(btnToggleHistory, Color.FromArgb(238, 242, 255), Color.FromArgb(67, 56, 202), 34);
            btnToggleHistory.Click += (s, e) => ToggleRightPanel();

            cardInput.Controls.Add(lblTitle);
            cardInput.Controls.Add(btnToggleHistory);
            y += 48;

            // Unpaid Appointments Dropdown
            lblSelectAppointment = new Label { Text = "Select Active Client to Bill (Unpaid Queue):", Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(71, 85, 105), Location = new Point(X, y), AutoSize = true };
            cmbSelectAppointment = new ComboBox { Location = new Point(X, y + 22), Size = new Size(W, 32), DropDownStyle = ComboBoxStyle.DropDownList, Font = new Font("Segoe UI", 10F) };
            cmbSelectAppointment.SelectedIndexChanged += CmbSelectAppointment_SelectedIndexChanged;
            cardInput.Controls.Add(lblSelectAppointment);
            cardInput.Controls.Add(cmbSelectAppointment);
            y += 66;

            // Customer Booking Details Card
            Panel pnlCard = new Panel { Location = new Point(X, y), Size = new Size(W, 160), BackColor = Color.FromArgb(248, 250, 252), BorderStyle = BorderStyle.FixedSingle };
            Label lblH = new Label { Text = "CUSTOMER BOOKING DETAILS", Font = new Font("Segoe UI", 8F, FontStyle.Bold), ForeColor = Color.FromArgb(100, 116, 139), Location = new Point(14, 10), AutoSize = true };

            lblCustomerNameDisplay = new Label
            {
                Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_customerName),
                Font = new Font("Segoe UI", 19F, FontStyle.Bold),
                ForeColor = Color.FromArgb(15, 23, 42),
                Location = new Point(12, 26),
                Size = new Size(W - 24, 48),
                AutoEllipsis = true
            };

            lblBarberTag = new Label { Text = $"✂ Barber: {_barberName}", Font = new Font("Segoe UI", 9F, FontStyle.Bold), BackColor = Color.FromArgb(238, 242, 255), ForeColor = Color.FromArgb(67, 56, 202), Location = new Point(14, 82), AutoSize = true, Padding = new Padding(6, 4, 6, 4) };
            lblPhoneTag = new Label { Text = $"📞 {_customerPhone}", Font = new Font("Segoe UI", 9F, FontStyle.Regular), BackColor = Color.FromArgb(241, 245, 249), ForeColor = Color.FromArgb(71, 85, 105), Location = new Point(230, 82), AutoSize = true, Padding = new Padding(6, 4, 6, 4) };
            lblServicesDisplay = new Label { Text = $"Services: {_servicesBooked}", Font = new Font("Segoe UI", 9F, FontStyle.Regular), ForeColor = Color.FromArgb(51, 65, 85), Location = new Point(14, 120), Size = new Size(W - 28, 30), AutoEllipsis = true };

            pnlCard.Controls.AddRange(new Control[] { lblH, lblCustomerNameDisplay, lblBarberTag, lblPhoneTag, lblServicesDisplay });
            cardInput.Controls.Add(pnlCard);
            y += 175;

            // Subtotal
            cardInput.Controls.Add(new Label { Text = "Base Subtotal:", Font = new Font("Segoe UI", 11F), Location = new Point(X, y), AutoSize = true });
            lblSubtotalValue = new Label { Text = $"₱{_baseSubtotal:N2}", Font = new Font("Segoe UI", 14F, FontStyle.Bold), Location = new Point(160, y - 2), AutoSize = true };
            y += 42;

            // Discount
            cardInput.Controls.Add(new Label { Text = "Discount:", Font = new Font("Segoe UI", 11F), Location = new Point(X, y), AutoSize = true });
            cmbDiscount = new ComboBox { Location = new Point(160, y), Size = new Size(230, 30), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbDiscount.Items.AddRange(new object[] { "None (0%)", "Senior Citizen / PWD (20%)", "Student Discount (10%)", "Custom Amount (₱)" });
            cmbDiscount.SelectedIndex = 0;
            txtCustomDiscount = new TextBox { Location = new Point(400, y), Size = new Size(130, 28), Visible = false, Text = "0" };
            UIStyle.StyleInput(txtCustomDiscount);
            cmbDiscount.SelectedIndexChanged += (s, e) => { txtCustomDiscount.Visible = cmbDiscount.SelectedIndex == 3; OnInputChanged(); };
            txtCustomDiscount.TextChanged += (s, e) => OnInputChanged();
            y += 44;

            // Barber Tip
            cardInput.Controls.Add(new Label { Text = "Barber Tip:", Font = new Font("Segoe UI", 11F), Location = new Point(X, y), AutoSize = true });
            txtTip = new TextBox { Location = new Point(160, y), Size = new Size(160, 28), Text = "0" };
            UIStyle.StyleInput(txtTip);
            txtTip.TextChanged += (s, e) => OnInputChanged();
            y += 48;

            // Total Due Banner
            Panel pnlTotal = new Panel { Location = new Point(X, y), Size = new Size(W, 64), BackColor = Color.FromArgb(17, 24, 39) };
            pnlTotal.Controls.Add(new Label { Text = "TOTAL AMOUNT DUE", Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(203, 213, 225), Location = new Point(16, 22), AutoSize = true });
            lblTotalDueValue = new Label { Text = "₱0.00", Font = new Font("Segoe UI", 20F, FontStyle.Bold), ForeColor = UIStyle.AccentGoldDark, Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 16, 0), AutoSize = true };
            pnlTotal.Controls.Add(lblTotalDueValue);
            cardInput.Controls.Add(pnlTotal);
            y += 74;

            // Payment Mode (Cash and GCash ONLY)
            cardInput.Controls.Add(new Label { Text = "Payment Mode:", Font = new Font("Segoe UI", 11F), Location = new Point(X, y), AutoSize = true });
            cmbPaymentMethod = new ComboBox { Location = new Point(160, y), Size = new Size(150, 30), DropDownStyle = ComboBoxStyle.DropDownList };
            cmbPaymentMethod.Items.AddRange(new object[] { "Cash", "GCash" });
            cmbPaymentMethod.SelectedIndex = 0;
            lblRefNo = new Label { Text = "Ref #:", Location = new Point(320, y + 4), AutoSize = true, Visible = false };
            txtReferenceNo = new TextBox { Location = new Point(370, y), Size = new Size(160, 28), Visible = false };
            UIStyle.StyleInput(txtReferenceNo);
            cmbPaymentMethod.SelectedIndexChanged += (s, e) => {
                bool isDig = cmbPaymentMethod.SelectedIndex > 0;
                lblRefNo.Visible = txtReferenceNo.Visible = isDig;
                if (isDig) txtTendered.Text = _totalAmountDue.ToString("N2");
                OnInputChanged();
            };
            txtReferenceNo.TextChanged += (s, e) => UpdateReceiptPreview();
            y += 44;

            // Cash Received
            cardInput.Controls.Add(new Label { Text = "Cash Received:", Font = new Font("Segoe UI", 11F), Location = new Point(X, y + 4), AutoSize = true });
            txtTendered = new TextBox { Location = new Point(160, y), Size = new Size(180, 32), Font = new Font("Segoe UI", 13F, FontStyle.Bold), Text = "" };
            UIStyle.StyleInput(txtTendered);
            txtTendered.TextChanged += (s, e) => OnInputChanged();

            // Quick Cash Buttons
            FlowLayoutPanel pnlQuick = new FlowLayoutPanel { Location = new Point(160, y + 38), Size = new Size(380, 42), WrapContents = false };
            Button btnEx = new Button { Text = "Exact", Size = new Size(68, 32), Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 0, 5, 0) };
            UIStyle.StyleButton(btnEx, Color.FromArgb(241, 245, 249), UIStyle.TextDark, 32);
            btnEx.Click += (s, e) => txtTendered.Text = _totalAmountDue.ToString("N2");
            pnlQuick.Controls.Add(btnEx);

            foreach (int amt in new int[] { 100, 200, 500, 1000 })
            {
                Button b = new Button { Text = $"₱{amt:N0}", Size = new Size(amt == 1000 ? 76 : 64, 32), Font = new Font("Segoe UI", 9F, FontStyle.Bold), Cursor = Cursors.Hand, Margin = new Padding(0, 0, 5, 0) };
                UIStyle.StyleButton(b, Color.FromArgb(241, 245, 249), UIStyle.TextDark, 32);
                b.Click += (s, e) => txtTendered.Text = amt.ToString();
                pnlQuick.Controls.Add(b);
            }
            y += 84;

            // Change Due Banner
            Panel pnlChg = new Panel { Location = new Point(X, y), Size = new Size(W, 54), BackColor = Color.FromArgb(240, 253, 244), BorderStyle = BorderStyle.FixedSingle };
            pnlChg.Controls.Add(new Label { Text = "CHANGE DUE:", Font = new Font("Segoe UI", 10.5F, FontStyle.Bold), ForeColor = Color.FromArgb(22, 101, 52), Location = new Point(14, 16), AutoSize = true });
            lblChangeValue = new Label { Text = "₱0.00", Font = new Font("Segoe UI", 16.5F, FontStyle.Bold), ForeColor = UIStyle.ActionSuccess, Dock = DockStyle.Right, TextAlign = ContentAlignment.MiddleRight, Padding = new Padding(0, 0, 14, 0), AutoSize = true };
            pnlChg.Controls.Add(lblChangeValue);
            cardInput.Controls.Add(pnlChg);
            y += 70;

            // Action Buttons
            btnConfirmPayment = new Button { Text = "💳 Confirm Payment", Location = new Point(X, y), Size = new Size(260, 48), Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand };
            UIStyle.StyleButton(btnConfirmPayment, UIStyle.ActionSuccess, Color.White, 48);
            btnConfirmPayment.Click += BtnConfirmPayment_Click;

            btnPrintReceipt = new Button { Text = "🖨️ Print Slip", Location = new Point(X + 280, y), Size = new Size(260, 48), Font = new Font("Segoe UI", 11F, FontStyle.Bold), Cursor = Cursors.Hand, Enabled = false };
            UIStyle.StyleButton(btnPrintReceipt, Color.FromArgb(203, 213, 225), Color.FromArgb(100, 116, 139), 48);
            btnPrintReceipt.Click += BtnPrintReceipt_Click;
            y += 56;

            btnClose = new Button { Text = "Close", Location = new Point(X, y), Size = new Size(W, 38), Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), Cursor = Cursors.Hand };
            UIStyle.StyleButton(btnClose, Color.FromArgb(241, 245, 249), UIStyle.TextDark, 38);
            btnClose.Click += (s, e) => this.Close();

            cardInput.Controls.AddRange(new Control[] {
                lblSubtotalValue, cmbDiscount, txtCustomDiscount, txtTip,
                cmbPaymentMethod, lblRefNo, txtReferenceNo, txtTendered, pnlQuick,
                btnConfirmPayment, btnPrintReceipt, btnClose
            });

            // ==========================================
            // RIGHT COLUMN: Toggleable Receipt / History
            // ==========================================
            cardRight = UIStyle.CreateCard();
            cardRight.Dock = DockStyle.Fill;
            cardRight.Padding = new Padding(14);

            Panel pnlRightHeaderBar = new Panel { Dock = DockStyle.Top, Height = 36 };
            lblRightHeader = new Label { Text = "RECEIPT PREVIEW", Font = new Font("Segoe UI", 10F, FontStyle.Bold), ForeColor = Color.FromArgb(100, 116, 139), Location = new Point(0, 6), AutoSize = true };

            btnBackToHistory = new Button
            {
                Text = "📜 Back to History List",
                Size = new Size(160, 32),
                Font = new Font("Segoe UI", 9F, FontStyle.Bold),
                Cursor = Cursors.Hand,
                Dock = DockStyle.Right,
                Visible = false
            };
            UIStyle.StyleButton(btnBackToHistory, Color.FromArgb(238, 242, 255), Color.FromArgb(67, 56, 202), 32);
            btnBackToHistory.Click += (s, e) => ToggleRightPanel();

            pnlRightHeaderBar.Controls.Add(lblRightHeader);
            pnlRightHeaderBar.Controls.Add(btnBackToHistory);

            // Panel 1: Receipt Preview
            pnlReceiptView = new Panel { Dock = DockStyle.Fill };
            rtbReceiptPreview = new RichTextBox { Dock = DockStyle.Fill, Font = new Font("Consolas", 9.5F), ReadOnly = true, BackColor = Color.White };
            pnlReceiptView.Controls.Add(rtbReceiptPreview);

            // Panel 2: In-Card Paid History
            pnlHistoryView = new Panel { Dock = DockStyle.Fill, Visible = false };
            dgvPaidHistory = new DataGridView
            {
                Dock = DockStyle.Fill,
                ReadOnly = true,
                AllowUserToAddRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect,
                BackgroundColor = Color.White,
                BorderStyle = BorderStyle.None,
                RowHeadersVisible = false,
                AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill,
                EnableHeadersVisualStyles = false,
                Cursor = Cursors.Hand
            };
            dgvPaidHistory.ColumnHeadersDefaultCellStyle.BackColor = Color.FromArgb(30, 41, 59);
            dgvPaidHistory.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dgvPaidHistory.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 9F, FontStyle.Bold);
            dgvPaidHistory.ColumnHeadersHeight = 34;
            dgvPaidHistory.RowTemplate.Height = 32;
            dgvPaidHistory.AlternatingRowsDefaultCellStyle.BackColor = Color.FromArgb(248, 250, 252);

            // Clicking ANY row loads and displays that customer's receipt
            dgvPaidHistory.CellClick += DgvPaidHistory_CellClick;

            lblHistoryTotalRevenue = new Label { Dock = DockStyle.Bottom, Height = 30, Font = new Font("Segoe UI", 9.5F, FontStyle.Bold), ForeColor = Color.FromArgb(22, 101, 52), TextAlign = ContentAlignment.MiddleLeft };
            pnlHistoryView.Controls.Add(dgvPaidHistory);
            pnlHistoryView.Controls.Add(lblHistoryTotalRevenue);

            cardRight.Controls.Add(pnlReceiptView);
            cardRight.Controls.Add(pnlHistoryView);
            cardRight.Controls.Add(pnlRightHeaderBar);

            mainLayout.Controls.Add(cardInput, 0, 0);
            mainLayout.Controls.Add(cardRight, 1, 0);
            this.Controls.Add(mainLayout);

            printDoc = new PrintDocument();
            printDoc.PrintPage += (s, e) => e.Graphics.DrawString(rtbReceiptPreview.Text, new Font("Consolas", 9.5F), Brushes.Black, new PointF(10, 10));
        }

        private void ToggleRightPanel()
        {
            bool showHistory = !pnlHistoryView.Visible;
            pnlHistoryView.Visible = showHistory;
            pnlReceiptView.Visible = !showHistory;
            btnBackToHistory.Visible = false;

            if (showHistory)
            {
                lblRightHeader.Text = "PAID RECEIPTS & SALES (Click customer to view receipt)";
                btnToggleHistory.Text = "🧾 Current Slip";
                LoadPaidHistoryData();
            }
            else
            {
                lblRightHeader.Text = _isPaymentCompleted ? "OFFICIAL RECEIPT PREVIEW" : "RECEIPT PREVIEW";
                btnToggleHistory.Text = "📜 Paid";
                UpdateReceiptPreview();
            }
        }

        private void DgvPaidHistory_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            var row = dgvPaidHistory.Rows[e.RowIndex];
            string recStr = row.Cells["Receipt #"]?.Value?.ToString() ?? "";

            if (int.TryParse(recStr.Replace("REC-", "").Trim(), out int transId))
            {
                LoadHistoricalReceipt(transId);
            }
        }

        private void LoadHistoricalReceipt(int transId)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(GlobalConfig.ConnectionString))
                {
                    conn.Open();
                    string q = "SELECT * FROM Transactions WHERE TransactionID = @id";

                    using (SqlCommand cmd = new SqlCommand(q, conn))
                    {
                        cmd.Parameters.AddWithValue("@id", transId);
                        using (SqlDataReader rdr = cmd.ExecuteReader())
                        {
                            if (rdr.Read())
                            {
                                bool HasCol(string colName)
                                {
                                    for (int i = 0; i < rdr.FieldCount; i++)
                                    {
                                        if (rdr.GetName(i).Equals(colName, StringComparison.OrdinalIgnoreCase))
                                            return true;
                                    }
                                    return false;
                                }

                                string cName = HasCol("CustomerName") && rdr["CustomerName"] != DBNull.Value ? rdr["CustomerName"].ToString() : "Customer";
                                string bName = HasCol("BarberName") && rdr["BarberName"] != DBNull.Value ? rdr["BarberName"].ToString() : "Assigned Staff";
                                string sName = HasCol("ServiceName") && rdr["ServiceName"] != DBNull.Value ? rdr["ServiceName"].ToString() : "Haircut Service";

                                decimal total = HasCol("TotalAmount") && rdr["TotalAmount"] != DBNull.Value ? Convert.ToDecimal(rdr["TotalAmount"]) : 0m;
                                decimal disc = HasCol("Discount") && rdr["Discount"] != DBNull.Value ? Convert.ToDecimal(rdr["Discount"]) : 0m;
                                decimal tip = HasCol("Tip") && rdr["Tip"] != DBNull.Value ? Convert.ToDecimal(rdr["Tip"]) : 0m;
                                decimal tendered = HasCol("AmountTendered") && rdr["AmountTendered"] != DBNull.Value ? Convert.ToDecimal(rdr["AmountTendered"]) : total;
                                decimal change = HasCol("ChangeDue") && rdr["ChangeDue"] != DBNull.Value ? Convert.ToDecimal(rdr["ChangeDue"]) : 0m;

                                string payMode = HasCol("PaymentMethod") && rdr["PaymentMethod"] != DBNull.Value ? rdr["PaymentMethod"].ToString() : "Cash";
                                string refNo = HasCol("ReferenceNo") && rdr["ReferenceNo"] != DBNull.Value ? rdr["ReferenceNo"].ToString() : "";
                                DateTime date = HasCol("TransactionDate") && rdr["TransactionDate"] != DBNull.Value ? Convert.ToDateTime(rdr["TransactionDate"]) : DateTime.Now;

                                decimal subtotal = (total + disc) - tip;
                                if (subtotal <= 0) subtotal = total;

                                string receipt =
                                    "========================================\n" +
                                    "            THE CLASSIC CUT             \n" +
                                    "         BARBERSHOP MANAGEMENT          \n" +
                                    "========================================\n" +
                                    $"Date: {date:MMM dd, yyyy  hh:mm tt}\n" +
                                    $"Receipt #: REC-{transId:D5} (PAID ARCHIVE)\n" +
                                    $"Customer: {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(cName)}\n" +
                                    $"Barber:   {bName}\n" +
                                    "----------------------------------------\n" +
                                    $"Services:\n{sName}\n" +
                                    "----------------------------------------\n" +
                                    $"Subtotal:                  ₱{subtotal,10:N2}\n";

                                if (disc > 0) receipt += $"Discount:                 -₱{disc,10:N2}\n";
                                if (tip > 0) receipt += $"Barber Tip:               +₱{tip,10:N2}\n";

                                receipt += "----------------------------------------\n" +
                                           $"TOTAL AMOUNT PAID:         ₱{total,10:N2}\n" +
                                           $"Payment Mode: {payMode}\n";

                                if (!string.IsNullOrEmpty(refNo)) receipt += $"Ref No: {refNo}\n";

                                receipt += $"Cash Received:             ₱{tendered,10:N2}\n" +
                                           $"Change Given:              ₱{change,10:N2}\n" +
                                           "========================================\n" +
                                           "        Thank you for your visit!       \n" +
                                           "========================================\n";

                                rtbReceiptPreview.Text = receipt;

                                pnlReceiptView.Visible = true;
                                pnlHistoryView.Visible = false;
                                btnBackToHistory.Visible = true;
                                lblRightHeader.Text = $"RECEIPT FOR: {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(cName).ToUpper()} (REC-{transId:D5})";
                                btnToggleHistory.Text = "📜 Paid";
                                btnPrintReceipt.Enabled = true;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Unable to load receipt: " + ex.Message, "Notice", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void LoadPaidHistoryData()
        {
            DataTable dt = new DataTable();
            dt.Columns.Add("Receipt #");
            dt.Columns.Add("Customer");
            dt.Columns.Add("Barber");
            dt.Columns.Add("Paid");
            dt.Columns.Add("Mode");
            dt.Columns.Add("Time");

            decimal totalRevenue = 0m;
            try
            {
                using (SqlConnection conn = new SqlConnection(GlobalConfig.ConnectionString))
                {
                    conn.Open();
                    string q = @"
                        SELECT TOP 50 TransactionID, CustomerName, BarberName, TotalAmount,
                               ISNULL(PaymentMethod, 'Cash') AS PaymentMethod, TransactionDate
                        FROM Transactions WHERE Status = 'Completed' ORDER BY TransactionDate DESC";

                    using (SqlCommand cmd = new SqlCommand(q, conn))
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            int tId = Convert.ToInt32(rdr["TransactionID"]);
                            decimal amt = Convert.ToDecimal(rdr["TotalAmount"]);
                            totalRevenue += amt;
                            dt.Rows.Add(
                                $"REC-{tId:D5}",
                                CultureInfo.CurrentCulture.TextInfo.ToTitleCase(rdr["CustomerName"].ToString()),
                                rdr["BarberName"].ToString(),
                                $"₱{amt:N2}",
                                rdr["PaymentMethod"].ToString(),
                                Convert.ToDateTime(rdr["TransactionDate"]).ToString("hh:mm tt")
                            );
                        }
                    }
                }
            }
            catch { }

            dgvPaidHistory.DataSource = dt;
            lblHistoryTotalRevenue.Text = $"Total Revenue Today: ₱{totalRevenue:N2}";
        }

        private void OnInputChanged()
        {
            if (_isPaymentCompleted) ResetButtonStates();
            CalculateBill();
        }

        private void ResetButtonStates()
        {
            _isPaymentCompleted = false;
            btnConfirmPayment.Enabled = true;
            btnConfirmPayment.Text = "💳 Confirm Payment";
            UIStyle.StyleButton(btnConfirmPayment, UIStyle.ActionSuccess, Color.White, 48);
            btnPrintReceipt.Enabled = false;
            UIStyle.StyleButton(btnPrintReceipt, Color.FromArgb(203, 213, 225), Color.FromArgb(100, 116, 139), 48);
        }

        private void SetPaymentCompletedState()
        {
            _isPaymentCompleted = true;
            btnConfirmPayment.Enabled = false;
            btnConfirmPayment.Text = "✔ Payment Completed";
            UIStyle.StyleButton(btnConfirmPayment, Color.FromArgb(220, 252, 231), Color.FromArgb(22, 101, 52), 48);
            btnPrintReceipt.Enabled = true;
            UIStyle.StyleButton(btnPrintReceipt, UIStyle.AccentGoldDark, Color.White, 48);
            btnPrintReceipt.Focus();
        }

        private void UpdateClientDisplay()
        {
            if (lblCustomerNameDisplay != null)
                lblCustomerNameDisplay.Text = CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_customerName);
            if (lblBarberTag != null) lblBarberTag.Text = $"✂ Barber: {_barberName}";
            if (lblPhoneTag != null) lblPhoneTag.Text = $"📞 {_customerPhone}";
            if (lblServicesDisplay != null) lblServicesDisplay.Text = $"Services: {_servicesBooked}";
            if (lblSubtotalValue != null) lblSubtotalValue.Text = $"₱{_baseSubtotal:N2}";
        }

        private class PendingAppointmentItem
        {
            public int TransactionID { get; set; }
            public string CustomerName { get; set; }
            public string CustomerPhone { get; set; }
            public string BarberName { get; set; }
            public string ServiceName { get; set; }
            public decimal TotalAmount { get; set; }
            public override string ToString() => $"{CustomerName} — {ServiceName} [₱{TotalAmount:N2}]";
        }

        private void LoadPendingAppointments()
        {
            if (cmbSelectAppointment == null) return;
            cmbSelectAppointment.Items.Clear();

            try
            {
                using (SqlConnection conn = new SqlConnection(GlobalConfig.ConnectionString))
                {
                    conn.Open();
                    string query = @"
                        SELECT TransactionID, CustomerName, CustomerPhone, BarberName, ServiceName, TotalAmount 
                        FROM Transactions 
                        WHERE ISNULL(Status, 'Scheduled') IN ('In Chair', 'Scheduled') 
                        ORDER BY TransactionDate DESC";

                    using (SqlCommand cmd = new SqlCommand(query, conn))
                    using (SqlDataReader rdr = cmd.ExecuteReader())
                    {
                        while (rdr.Read())
                        {
                            cmbSelectAppointment.Items.Add(new PendingAppointmentItem
                            {
                                TransactionID = Convert.ToInt32(rdr["TransactionID"]),
                                CustomerName = rdr["CustomerName"].ToString(),
                                CustomerPhone = rdr["CustomerPhone"].ToString(),
                                BarberName = rdr["BarberName"].ToString(),
                                ServiceName = rdr["ServiceName"].ToString(),
                                TotalAmount = Convert.ToDecimal(rdr["TotalAmount"])
                            });
                        }
                    }
                }

                if (cmbSelectAppointment.Items.Count > 0)
                {
                    int matchIdx = -1;
                    if (_transactionId > 0)
                    {
                        for (int i = 0; i < cmbSelectAppointment.Items.Count; i++)
                        {
                            if (((PendingAppointmentItem)cmbSelectAppointment.Items[i]).TransactionID == _transactionId)
                            {
                                matchIdx = i;
                                break;
                            }
                        }
                    }
                    cmbSelectAppointment.SelectedIndex = (matchIdx >= 0) ? matchIdx : 0;
                }
            }
            catch { }
        }

        private void CmbSelectAppointment_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSelectAppointment.SelectedItem is PendingAppointmentItem item)
            {
                _transactionId = item.TransactionID;
                _customerName = item.CustomerName;
                _customerPhone = item.CustomerPhone;
                _barberName = item.BarberName;
                _servicesBooked = item.ServiceName;
                _baseSubtotal = item.TotalAmount;

                UpdateClientDisplay();
                ResetButtonStates();
                CalculateBill();
            }
        }

        private void CalculateBill()
        {
            _discountAmount = 0m;
            int discIndex = cmbDiscount?.SelectedIndex ?? 0;
            if (discIndex == 1) _discountAmount = _baseSubtotal * 0.20m;
            else if (discIndex == 2) _discountAmount = _baseSubtotal * 0.10m;
            else if (discIndex == 3) decimal.TryParse(txtCustomDiscount.Text.Replace("₱", "").Replace(",", "").Trim(), out _discountAmount);
            if (_discountAmount > _baseSubtotal) _discountAmount = _baseSubtotal;

            _tipAmount = 0m;
            decimal.TryParse(txtTip.Text.Replace("₱", "").Replace(",", "").Trim(), out _tipAmount);
            _totalAmountDue = Math.Max(0m, (_baseSubtotal - _discountAmount) + _tipAmount);
            if (lblTotalDueValue != null) lblTotalDueValue.Text = $"₱{_totalAmountDue:N2}";

            decimal.TryParse(txtTendered.Text.Replace("₱", "").Replace(",", "").Trim(), out _amountTendered);
            _changeDue = Math.Max(0m, _amountTendered - _totalAmountDue);
            if (lblChangeValue != null) lblChangeValue.Text = $"₱{_changeDue:N2}";

            UpdateReceiptPreview();
        }

        // Only generates the official receipt when payment has been confirmed!
        private void UpdateReceiptPreview()
        {
            if (rtbReceiptPreview == null) return;

            // Before payment is completed: show a clean pending notice (NOT the official receipt)
            if (!_isPaymentCompleted)
            {
                lblRightHeader.Text = "RECEIPT PREVIEW (PAYMENT PENDING)";
                string pendingNotice =
                    "========================================\n" +
                    "            THE CLASSIC CUT             \n" +
                    "         BARBERSHOP MANAGEMENT          \n" +
                    "========================================\n" +
                    $"Date: {DateTime.Now:MMM dd, yyyy  hh:mm tt}\n" +
                    $"Customer: {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_customerName)}\n" +
                    $"Barber:   {_barberName}\n" +
                    "----------------------------------------\n" +
                    $"Services:\n{_servicesBooked}\n" +
                    "----------------------------------------\n" +
                    $"TOTAL DUE:                 ₱{_totalAmountDue,10:N2}\n" +
                    "========================================\n\n" +
                    "          *** PAYMENT PENDING ***        \n\n" +
                    "  The official receipt will be generated \n" +
                    "  and ready to print once payment is     \n" +
                    "  confirmed.                             \n\n" +
                    "========================================\n";

                rtbReceiptPreview.Text = pendingNotice;
                return;
            }

            // Once confirmed: generate the full official receipt slip
            lblRightHeader.Text = "OFFICIAL RECEIPT PREVIEW";
            string payMethod = cmbPaymentMethod.SelectedItem?.ToString() ?? "Cash";
            string refNo = txtReferenceNo.Text.Trim();

            string receipt =
                "========================================\n" +
                "            THE CLASSIC CUT             \n" +
                "         BARBERSHOP MANAGEMENT          \n" +
                "========================================\n" +
                $"Date: {DateTime.Now:MMM dd, yyyy  hh:mm tt}\n" +
                $"Receipt #: REC-{_transactionId:D5}\n" +
                $"Customer: {CultureInfo.CurrentCulture.TextInfo.ToTitleCase(_customerName)}\n" +
                $"Barber:   {_barberName}\n" +
                "----------------------------------------\n" +
                $"Services:\n{_servicesBooked}\n" +
                "----------------------------------------\n" +
                $"Subtotal:                  ₱{_baseSubtotal,10:N2}\n";

            if (_discountAmount > 0) receipt += $"Discount:                 -₱{_discountAmount,10:N2}\n";
            if (_tipAmount > 0) receipt += $"Barber Tip:               +₱{_tipAmount,10:N2}\n";

            receipt += "----------------------------------------\n" +
                       $"TOTAL AMOUNT PAID:         ₱{_totalAmountDue,10:N2}\n" +
                       $"Payment Mode: {payMethod}\n";

            if (!string.IsNullOrEmpty(refNo)) receipt += $"Ref No: {refNo}\n";

            receipt += $"Cash Received:             ₱{_amountTendered,10:N2}\n" +
                       $"Change Given:              ₱{_changeDue,10:N2}\n" +
                       "========================================\n" +
                       "        Thank you for your visit!       \n" +
                       "========================================\n";

            rtbReceiptPreview.Text = receipt;
        }

        private void BtnConfirmPayment_Click(object sender, EventArgs e)
        {
            if (_amountTendered < _totalAmountDue)
            {
                MessageBox.Show($"Cash received (₱{_amountTendered:N2}) is less than total due (₱{_totalAmountDue:N2}).", "Insufficient Payment", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                txtTendered.Focus();
                return;
            }

            try
            {
                using (SqlConnection conn = new SqlConnection(GlobalConfig.ConnectionString))
                {
                    conn.Open();

                    HashSet<string> existingTables = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    using (SqlCommand cmdTbls = new SqlCommand("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES", conn))
                    using (SqlDataReader r = cmdTbls.ExecuteReader())
                    {
                        while (r.Read()) existingTables.Add(r.GetString(0));
                    }

                    HashSet<string> transCols = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    if (existingTables.Contains("Transactions"))
                    {
                        using (SqlCommand cmdCols = new SqlCommand("SELECT COLUMN_NAME FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = 'Transactions'", conn))
                        using (SqlDataReader r = cmdCols.ExecuteReader())
                        {
                            while (r.Read()) transCols.Add(r.GetString(0));
                        }
                    }

                    string payMethod = cmbPaymentMethod.SelectedItem?.ToString() ?? "Cash";
                    string refNo = txtReferenceNo.Text.Trim();

                    // 1. Update / Insert Transactions Table
                    if (_transactionId > 0 && existingTables.Contains("Transactions"))
                    {
                        List<string> setClauses = new List<string> { "Status = 'Completed'" };
                        SqlCommand cmd = new SqlCommand { Connection = conn };
                        if (transCols.Contains("TotalAmount")) { setClauses.Add("TotalAmount = @total"); cmd.Parameters.AddWithValue("@total", _totalAmountDue); }
                        if (transCols.Contains("Discount")) { setClauses.Add("Discount = @disc"); cmd.Parameters.AddWithValue("@disc", _discountAmount); }
                        if (transCols.Contains("PaymentMethod")) { setClauses.Add("PaymentMethod = @payMethod"); cmd.Parameters.AddWithValue("@payMethod", payMethod); }
                        if (transCols.Contains("Tip")) { setClauses.Add("Tip = @tip"); cmd.Parameters.AddWithValue("@tip", _tipAmount); }
                        if (transCols.Contains("AmountTendered")) { setClauses.Add("AmountTendered = @tendered"); cmd.Parameters.AddWithValue("@tendered", _amountTendered); }
                        if (transCols.Contains("ChangeDue")) { setClauses.Add("ChangeDue = @change"); cmd.Parameters.AddWithValue("@change", _changeDue); }
                        if (transCols.Contains("ReferenceNo") && !string.IsNullOrEmpty(refNo)) { setClauses.Add("ReferenceNo = @refNo"); cmd.Parameters.AddWithValue("@refNo", refNo); }

                        cmd.CommandText = $"UPDATE Transactions SET {string.Join(", ", setClauses)} WHERE TransactionID = @id";
                        cmd.Parameters.AddWithValue("@id", _transactionId);
                        cmd.ExecuteNonQuery();
                    }
                    else if (existingTables.Contains("Transactions"))
                    {
                        SqlCommand cmdInsert = new SqlCommand(@"
                            INSERT INTO Transactions (CustomerName, BarberName, ServiceName, TotalAmount, Status, TransactionDate)
                            VALUES (@cust, @barber, @service, @total, 'Completed', GETDATE());
                            SELECT SCOPE_IDENTITY();", conn);
                        cmdInsert.Parameters.AddWithValue("@cust", _customerName);
                        cmdInsert.Parameters.AddWithValue("@barber", _barberName);
                        cmdInsert.Parameters.AddWithValue("@service", _servicesBooked);
                        cmdInsert.Parameters.AddWithValue("@total", _totalAmountDue);
                        object newId = cmdInsert.ExecuteScalar();
                        if (newId != null && int.TryParse(newId.ToString(), out int parsedId)) _transactionId = parsedId;
                    }

                    // 2. Also Update Appointments Table (Sets Status to Completed)
                    if (existingTables.Contains("Appointments"))
                    {
                        using (SqlCommand cmdApp = new SqlCommand(@"
                            UPDATE Appointments 
                            SET Status = 'Completed' 
                            WHERE CustomerName = @custName AND Status IN ('In Chair', 'Scheduled')", conn))
                        {
                            cmdApp.Parameters.AddWithValue("@custName", _customerName);
                            cmdApp.ExecuteNonQuery();
                        }
                    }

                    // 3. Free Barber in Roster
                    if (!string.IsNullOrEmpty(_barberName) && _barberName != "Assigned Barber" && _barberName != "All-Around Staff" && existingTables.Contains("Barbers"))
                    {
                        using (SqlCommand cmdB = new SqlCommand("UPDATE Barbers SET Status = 'Available' WHERE BarberName = @bName", conn))
                        {
                            cmdB.Parameters.AddWithValue("@bName", _barberName);
                            cmdB.ExecuteNonQuery();
                        }
                    }

                    GlobalConfig.LogAudit(GlobalConfig.CurrentUserID, "PAYMENT", "Transactions", $"Completed Payment #{_transactionId} for {_customerName} — Paid ₱{_totalAmountDue:N2} via {payMethod}");
                }

                RefreshDashboardAppointmentsQueue();

                this.DialogResult = DialogResult.OK;
                SetPaymentCompletedState(); // Sets completed to true first!
                UpdateReceiptPreview();     // Now renders the official receipt

                LoadPaidHistoryData();
                LoadPendingAppointments();

                MessageBox.Show($"Payment completed successfully!\n\nReceipt #: REC-{_transactionId:D5}\nChange Due: ₱{_changeDue:N2}", "Payment Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error processing payment: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RefreshDashboardAppointmentsQueue()
        {
            try
            {
                foreach (Form openForm in Application.OpenForms)
                {
                    if (openForm == this) continue;
                    string[] refreshMethods = new string[] {
                        "LoadAppointments", "LoadTodayAppointments", "LoadAppointmentsQueue",
                        "LoadQueue", "LoadData", "RefreshData", "RefreshAppointments", "LoadAllAppointments"
                    };
                    foreach (string mName in refreshMethods)
                    {
                        MethodInfo method = openForm.GetType().GetMethod(mName, BindingFlags.Instance | BindingFlags.Public | BindingFlags.NonPublic);
                        if (method != null && method.GetParameters().Length == 0)
                        {
                            try { method.Invoke(openForm, null); } catch { }
                        }
                    }
                }
            }
            catch { }
        }

        private void BtnPrintReceipt_Click(object sender, EventArgs e)
        {
            if (!_isPaymentCompleted && !rtbReceiptPreview.Text.Contains("REC-")) return;
            try
            {
                using (PrintPreviewDialog previewDialog = new PrintPreviewDialog())
                {
                    previewDialog.Document = printDoc;
                    previewDialog.Width = 600;
                    previewDialog.Height = 700;
                    previewDialog.StartPosition = FormStartPosition.CenterParent;
                    previewDialog.ShowDialog();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Printing error: " + ex.Message, "Printer Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }
    }
}
