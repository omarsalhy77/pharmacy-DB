using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace database_project
{
    public partial class Form1 : Form
    {
        SqlConnection con = new SqlConnection(
@"Data Source=VIVO\SQLEXPRESS;Initial Catalog=PharmacyDB;Integrated Security=True");
        int selectedCustomerID = 0;
        int selectedMedicineID = 0;
        int selectedSupplierID = 0;
        int selectedStaffID = 0;
        DataTable cart = new DataTable();

        public Form1()
        {
            
            InitializeComponent();
        }
        private bool CheckStock(int medicineID, int qty)
        {
            string q = "select Quantity from Medicines where MedicineID=" + medicineID;

            SqlDataAdapter da = new SqlDataAdapter(q, con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            int stock = Convert.ToInt32(dt.Rows[0][0]);

            return qty <= stock;
        }

        void LoadCustomers()
        {
            string query = "select * from Customers";

            SqlDataAdapter da = new SqlDataAdapter(query, con);
            DataTable dt = new DataTable();
            da.Fill(dt);
           // MessageBox.Show(dt.Rows.Count.ToString());

            dgvCustomers.DataSource = null;
            dgvCustomers.Columns.Clear();
            dgvCustomers.AutoGenerateColumns = true;
            dgvCustomers.DataSource = dt;
        }
        void LoadSuppliers()
        {
            SqlDataAdapter da =
            new SqlDataAdapter("select SupplierID, SupplierName from Suppliers", con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbSupplier.DataSource = dt;
            cmbSupplier.DisplayMember = "SupplierName";
            cmbSupplier.ValueMember = "SupplierID";
        }
        void LoadCategories()
        {
            SqlDataAdapter da =
            new SqlDataAdapter("select CategoryID, CategoryName from Categories", con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbCategory.DataSource = dt;
            cmbCategory.DisplayMember = "CategoryName";
            cmbCategory.ValueMember = "CategoryID";
        }
        void LoadMedicines()
        {
            string query = @"select MedicineID, MedicineName, Price, Quantity,
                    ExpiryDate
                    from Medicines";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            dgvMedicines.DataSource = null;
            dgvMedicines.Columns.Clear();
            dgvMedicines.AutoGenerateColumns = true;
            dgvMedicines.DataSource = dt;
        }
        void LoadSuppliersGrid()
        {
            string query = "select * from Suppliers";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvSuppliers.DataSource = null;
            dgvSuppliers.Columns.Clear();
            dgvSuppliers.AutoGenerateColumns = true;
            dgvSuppliers.DataSource = dt;
        }
        void LoadStaff()
        {
            string query = "select * from Staff";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvStaff.DataSource = null;
            dgvStaff.Columns.Clear();
            dgvStaff.AutoGenerateColumns = true;
            dgvStaff.DataSource = dt;
        }
        void CalculateTotalOrder()
        {
            double sum = 0;

            foreach (DataRow row in purchaseCart.Rows)
            {
                sum += Convert.ToDouble(row["Total"]);
            }

            txtTotalOrder.Text = sum.ToString();
        }
        void CalculateTotal()
        {
            double sum = 0;

            foreach (DataRow row in cart.Rows)
                sum += Convert.ToDouble(row["Total"]);

            txtTotal.Text = sum.ToString();

            CalculateFinalTotal();
        }

        void CalculateFinalTotal()
        {
            double total = 0, discount = 0;

            double.TryParse(txtTotal.Text, out total);
            double.TryParse(txtDiscount.Text, out discount);

            double final = total * (discount/100);

            if (final < 0) final = 0;

            txtFinalTotal.Text = final.ToString();
        }
        DataTable purchaseCart = new DataTable();

        void CreatePurchaseOrderCart()
        {
            purchaseCart.Columns.Add("MedicineID");   // مخفي
            purchaseCart.Columns.Add("MedicineName");
            purchaseCart.Columns.Add("UnitPrice");
            purchaseCart.Columns.Add("Quantity");
            purchaseCart.Columns.Add("Total");

            dgvPurchaseOrder.DataSource = purchaseCart;
        }

        void CreateCartTable()
        {
            cart.Columns.Add("MedicineID");     // مخفي
            cart.Columns.Add("MedicineName");
            cart.Columns.Add("UnitPrice");
            cart.Columns.Add("Quantity");
            cart.Columns.Add("Total");

            dgvCart.DataSource = cart;
        }
        void LoadCustomersCombo()
        {
            SqlDataAdapter da =
            new SqlDataAdapter("select CustomerID, FirstName + ' ' + LastName as FullName from Customers", con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbCustomer.DataSource = dt;
            cmbCustomer.DisplayMember = "FullName";
            cmbCustomer.ValueMember = "CustomerID";
        }
        void LoadStaffCombo()
        {
            SqlDataAdapter da =
            new SqlDataAdapter("select StaffID, FirstName + ' ' + LastName as FullName from Staff", con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbStaff.DataSource = dt;
            cmbStaff.DisplayMember = "FullName";
            cmbStaff.ValueMember = "StaffID";
        }
        void LoadMedicinesCombo()
        {
            SqlDataAdapter da =
            new SqlDataAdapter("select MedicineID, MedicineName from Medicines", con);

            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbMedicine.DataSource = dt;
            cmbMedicine.DisplayMember = "MedicineName";
            cmbMedicine.ValueMember = "MedicineID";
        }
        void LoadSuppliersOrder()
        {
            string q = "select SupplierID, SupplierName from Suppliers";

            SqlDataAdapter da = new SqlDataAdapter(q, con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbSupplierOrder.DataSource = dt;
            cmbSupplierOrder.DisplayMember = "SupplierName";
            cmbSupplierOrder.ValueMember = "SupplierID";
        }
        void LoadMedicinesOrder()
        {
            string q = "select MedicineID, MedicineName from Medicines";

            SqlDataAdapter da = new SqlDataAdapter(q, con);
            DataTable dt = new DataTable();
            da.Fill(dt);

            cmbMedicineOrder.DataSource = dt;
            cmbMedicineOrder.DisplayMember = "MedicineName";
            cmbMedicineOrder.ValueMember = "MedicineID";
        }
        private void Form1_Load(object sender, EventArgs e)
        {
            LoadSuppliersOrder();
            LoadMedicinesOrder();
            CreatePurchaseOrderCart();
            LoadCustomersCombo();
            LoadStaffCombo();
            LoadMedicinesCombo();
            LoadCustomers();
            LoadSuppliers();
            LoadCategories();
            LoadMedicines();
            LoadSuppliersGrid();
            LoadStaff();
            CreateCartTable();
            //CreatePurchaseOrderCart();
            dgvCustomers.AutoGenerateColumns = true;
            dgvCustomers.AllowUserToAddRows = false;
            cmbRole.DropDownStyle = ComboBoxStyle.DropDownList;
            dgvCart.Columns["MedicineID"].Visible = false;
            dgvPurchaseOrder.Columns["MedicineID"].Visible = false;
            cmbRole.Items.AddRange(new string[]
{
    "Pharmacist",
    "Cashier",
    "Manager"
});
            //cmbMedicine.DisplayMember = "MedicineName"; // الاسم اللي اليوزر بيشوفه
            //cmbMedicine.ValueMember = "MedicineID";     // الرقم اللي الكود بيستخدمه
            //cmbMedicine.DataSource = dtMedicines;       // مصدر البيانات
        }
        
        private void button4_Click(object sender, EventArgs e)
        {
            
        }

        private void btnadd_Click(object sender, EventArgs e)
        {
            string query = "insert into Medicines " +
    "(MedicineName, SupplierID, CategoryID, Price, Quantity, ExpiryDate) values ('"
    + txtMedicineName.Text + "',"
    + cmbSupplier.SelectedValue + ","
    + cmbCategory.SelectedValue + ","
    + txtPrice.Text + ","
    + txtQuantity.Text + ",'"
    + dtpExpiryDate.Value.ToString("yyyy-MM-dd") + "')";

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Medicine Added");

            LoadMedicines();
            LoadMedicinesCombo();
             LoadMedicinesOrder();
            LoadMedicinesCombo();

        }

        private void btnaddcus_Click(object sender, EventArgs e)
        {
            string query = "insert into Customers (FirstName, LastName, Phone) values ('"
               + txtFirstName.Text + "','"
               + txtLastName.Text + "','"
               + txtPhone.Text + "')";

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Customer Added");
            LoadCustomers();
            dgvCustomers.Refresh();
            LoadCustomersCombo();

        }

        private void btnUpdateCustomer_Click(object sender, EventArgs e)
        {
            string query = "update Customers set " +
   "FirstName='" + txtFirstName.Text + "'," +
   "LastName='" + txtLastName.Text + "'," +
   "Phone='" + txtPhone.Text + "'" +
   " where CustomerID=" + selectedCustomerID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            LoadCustomers();
            LoadCustomersCombo();

            MessageBox.Show("Updated");
        }

        private void dgvCustomers_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedCustomerID =
       Convert.ToInt32(dgvCustomers.CurrentRow.Cells[0].Value);

            txtFirstName.Text =
                dgvCustomers.CurrentRow.Cells[1].Value.ToString();

            txtLastName.Text =
                dgvCustomers.CurrentRow.Cells[2].Value.ToString();

            txtPhone.Text =
                dgvCustomers.CurrentRow.Cells[3].Value.ToString();
        }

        private void btndeletecustomer_Click(object sender, EventArgs e)
        {
            string query = "delete from Customers where CustomerID=" + selectedCustomerID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Customer Deleted");

            LoadCustomers();
            LoadCustomersCombo();
        }

        private void btnSearchCustomer_Click(object sender, EventArgs e)
        {
            if (txtSearchCustomer.Text == "")
            {
                LoadCustomers();
                return;
            }

            string query = "select * from Customers where " +
                           "FirstName like '%" + txtSearchCustomer.Text + "%' " +
                           "or LastName like '%" + txtSearchCustomer.Text + "%'";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvCustomers.DataSource = dt;
        }

        private void btnUpdateMedicine_Click(object sender, EventArgs e)
        {
            string query = "update Medicines set " +
    "MedicineName='" + txtMedicineName.Text + "'," +
    "SupplierID=" + cmbSupplier.SelectedValue + "," +
    "CategoryID=" + cmbCategory.SelectedValue + "," +
    "Price=" + txtPrice.Text + "," +
    "Quantity=" + txtQuantity.Text + "," +
    "ExpiryDate='" + dtpExpiryDate.Value.ToString("yyyy-MM-dd") + "'" +
    " where MedicineID=" + selectedMedicineID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Updated");

            LoadMedicines();
            LoadMedicinesCombo();
        }

        private void btnDeleteMedicine_Click(object sender, EventArgs e)
        {
            string query =
    "delete from Medicines where MedicineID=" + selectedMedicineID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Deleted");

            LoadMedicines();
            LoadMedicinesCombo();
        }

        private void btnClearMedicine_Click(object sender, EventArgs e)
        {
            txtMedicineName.Clear();
            txtPrice.Clear();
            txtQuantity.Clear();
        }

        private void dgvMedicines_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedMedicineID =
   Convert.ToInt32(dgvMedicines.CurrentRow.Cells[0].Value);

            txtMedicineName.Text =
            dgvMedicines.CurrentRow.Cells[1].Value.ToString();

            txtPrice.Text =
            dgvMedicines.CurrentRow.Cells[2].Value.ToString();

            txtQuantity.Text =
            dgvMedicines.CurrentRow.Cells[3].Value.ToString();

            dtpExpiryDate.Value =
            Convert.ToDateTime(dgvMedicines.CurrentRow.Cells[4].Value);
        }

        private void btnsearch_Click(object sender, EventArgs e)
        {
            if (txtSearchMedicine.Text == "")
            {
                LoadMedicines();
                return;
            }

            string query = "select MedicineID, MedicineName, Price, Quantity, ExpiryDate " +
                           "from Medicines " +
                           "where MedicineName like '%" + txtSearchMedicine.Text + "%' " +
                           "or MedicineID like '%" + txtSearchMedicine.Text + "%'";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvMedicines.DataSource = dt;
        }

        private void btnAddSupplier_Click(object sender, EventArgs e)
        {
            string query = "insert into Suppliers (SupplierName, Phone, Address) values ('"
        + txtSupplierName.Text + "','"
        + txtSupplierPhone.Text + "','"
        + txtSupplierAddress.Text + "')";

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Supplier Added");

            LoadSuppliersGrid();
        }

        private void btnUpdateSupplier_Click(object sender, EventArgs e)
        {
            string query = "update Suppliers set " +
        "SupplierName='" + txtSupplierName.Text + "'," +
        "Phone='" + txtSupplierPhone.Text + "'," +
        "Address='" + txtSupplierAddress.Text + "'" +
        " where SupplierID=" + selectedSupplierID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Updated");

            LoadSuppliersGrid();
        }

        private void btnDeleteSupplier_Click(object sender, EventArgs e)
        {
            string query =
        "delete from Suppliers where SupplierID=" + selectedSupplierID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Deleted");

            LoadSuppliersGrid();
        }

        private void btnSearchSupplier_Click(object sender, EventArgs e)
        {
            if (txtSearchSupplier.Text == "")
            {
                LoadSuppliersGrid();
                return;
            }

            string query = "select * from Suppliers where " +
                "SupplierName like '%" + txtSearchSupplier.Text + "%' " +
                "or Phone like '%" + txtSearchSupplier.Text + "%'";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvSuppliers.DataSource = dt;

        }

        private void dgvSuppliers_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedSupplierID =
        Convert.ToInt32(dgvSuppliers.CurrentRow.Cells[0].Value);

            txtSupplierName.Text =
                dgvSuppliers.CurrentRow.Cells[1].Value.ToString();

            txtSupplierPhone.Text =
                dgvSuppliers.CurrentRow.Cells[2].Value.ToString();

            txtSupplierAddress.Text =
                dgvSuppliers.CurrentRow.Cells[3].Value.ToString();
        }

        private void btnAddStaff_Click(object sender, EventArgs e)
        {
            string query = "insert into Staff " +
    "(FirstName, LastName, Email, Phone, Role, Address, Salary) values ('"
    + txtFirstNamest.Text + "','"
    + txtLastNamest.Text + "','"
    + txtEmail.Text + "','"
    + txtPhonest.Text + "','"
    + cmbRole.Text + "','"
    + txtAddressst.Text + "',"
    + txtSalary.Text + ")";

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Staff Added");

            LoadStaff();
            LoadStaffCombo();
        }

        private void btnUpdateStaff_Click(object sender, EventArgs e)
        {
            string query = "update Staff set " +
    "FirstName='" + txtFirstNamest.Text + "'," +
    "LastName='" + txtLastNamest.Text + "'," +
    "Email='" + txtEmail.Text + "'," +
    "Phone='" + txtPhonest.Text + "'," +
    "Role='" + cmbRole.Text + "'," +
    "Address='" + txtAddressst.Text + "'," +
    "Salary=" + txtSalary.Text +
    " where StaffID=" + selectedStaffID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Updated");

            LoadStaff();
            LoadStaffCombo();
        }

        private void btnDeleteStaff_Click(object sender, EventArgs e)
        {
            string query =
        "delete from Staff where StaffID=" + selectedStaffID;

            SqlCommand cmd = new SqlCommand(query, con);

            con.Open();
            cmd.ExecuteNonQuery();
            con.Close();

            MessageBox.Show("Deleted");

            LoadStaff();
            LoadStaffCombo();
        }

        private void btnSearchStaff_Click(object sender, EventArgs e)
        {
            if (txtSearchStaff.Text == "")
            {
                LoadStaff();
                return;
            }

            string query = "select * from Staff where " +
                "FirstName like '%" + txtSearchStaff.Text + "%' " +
                "or LastName like '%" + txtSearchStaff.Text + "%' " +
                "or Role like '%" + txtSearchStaff.Text + "%'";

            SqlDataAdapter da = new SqlDataAdapter(query, con);

            DataTable dt = new DataTable();

            da.Fill(dt);

            dgvStaff.DataSource = dt;
        }

        private void dgvStaff_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            selectedStaffID =
      Convert.ToInt32(dgvStaff.CurrentRow.Cells[0].Value);

            txtFirstNamest.Text =
                dgvStaff.CurrentRow.Cells[1].Value.ToString();

            txtLastNamest.Text =
                dgvStaff.CurrentRow.Cells[2].Value.ToString();

            txtEmail.Text =
                dgvStaff.CurrentRow.Cells[3].Value.ToString();

            txtPhonest.Text =
                dgvStaff.CurrentRow.Cells[4].Value.ToString();

            cmbRole.Text =
                dgvStaff.CurrentRow.Cells[5].Value.ToString();

            txtAddressst.Text =
                dgvStaff.CurrentRow.Cells[6].Value.ToString();

            txtSalary.Text =
                dgvStaff.CurrentRow.Cells[7].Value.ToString();
        }

        private void btnclear_Click(object sender, EventArgs e)
        {
            txtFirstNamest.Clear();
            txtLastNamest.Clear();
            txtEmail.Clear();
            txtPhonest.Clear();
            txtAddressst.Clear();
            txtSalary.Clear();
            txtSearchStaff.Clear();

            cmbRole.SelectedIndex = -1;

            selectedStaffID = 0;
        }

        private void txtDiscount_TextChanged(object sender, EventArgs e)
        {
            CalculateFinalTotal();
        }

        private void btnAddToCart_Click(object sender, EventArgs e)
        {
            if (!CheckStock(
     Convert.ToInt32(cmbMedicine.SelectedValue),
     Convert.ToInt32(numQuantity.Value)
 ))
            {
                MessageBox.Show("Not enough stock!");
                return;
            }

            int qty = Convert.ToInt32(numQuantity.Value);

            double price = Convert.ToDouble(txtPricesal.Text);

            double total = qty * price;

            cart.Rows.Add(
                cmbMedicine.SelectedValue,
                cmbMedicine.Text,
                price,
                qty,
                total
            );

            CalculateTotal();
        }

        private void btnSaveSale_Click(object sender, EventArgs e)
        {
            if(txtDiscount.Text == "")
            {
                txtDiscount.Text = "0";
               // MessageBox.Show("Please enter a discount value");
               // return;
            }

            if (cart.Rows.Count == 0)
            {
                MessageBox.Show("Cart is empty");
                return;
            }

            con.Open();

            // Insert Sales
            string q1 = "insert into Sales (CustomerID, StaffID, SaleDate, TotalAmount, Discount) values ("
                + cmbCustomer.SelectedValue + ","
                + cmbStaff.SelectedValue + ", GETDATE(), "
                + txtFinalTotal.Text + ","
                + txtDiscount.Text + ")";

            SqlCommand cmd = new SqlCommand(q1, con);
            cmd.ExecuteNonQuery();

            // Get SaleID
            SqlCommand cmdId = new SqlCommand("select max(SaleID) from Sales", con);
            int saleID = Convert.ToInt32(cmdId.ExecuteScalar());

            // Insert Details + Update Stock
            foreach (DataRow row in cart.Rows)
            {
                string q2 = "insert into SaleDetails values ("
                    + saleID + ","
                    + row["MedicineID"] + ","
                    + row["Quantity"] + ","
                    + row["UnitPrice"] + ")";

                new SqlCommand(q2, con).ExecuteNonQuery();

                string q3 = "update Medicines set Quantity = Quantity - "
                    + row["Quantity"] +
                    " where MedicineID=" + row["MedicineID"];

                new SqlCommand(q3, con).ExecuteNonQuery();
            }

            con.Close();

            MessageBox.Show("Sale Saved Successfully");

            cart.Rows.Clear();
            txtTotal.Clear();
            txtDiscount.Clear();
            txtFinalTotal.Clear();
            LoadMedicines();
        }

        private void cmbMedicine_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (cmbMedicine.SelectedValue != null && cmbMedicine.SelectedIndex != -1)
                    {
                        try
                        {
                            // نستخدم .ToString() لضمان تحويل القيمة لنص يمكن استخدامه في الاستعلام
                            string q = "select Price from Medicines where MedicineID = " + cmbMedicine.SelectedValue.ToString();

                            SqlDataAdapter da = new SqlDataAdapter(q, con);
                            DataTable dt = new DataTable();
                            da.Fill(dt);

                            if (dt.Rows.Count > 0)
                            {
                                txtPrice.Text = dt.Rows[0]["Price"].ToString();
                            }
                    txtPricesal.Text = txtPrice.Text; // عشان السعر يظهر في خانة السعر في صفحة البيع كمان
                }
                        catch (Exception ex)
                        {
                            // عشان لو حصل أي خطأ تاني تعرف سببه إيه
                            // MessageBox.Show(ex.Message); 
                        }
                    }
           

        }

        private void btnClearSale_Click(object sender, EventArgs e)
        {
            cart.Rows.Clear();
            txtTotal.Clear();
            txtDiscount.Clear();
            txtFinalTotal.Clear();
        }

        private void btnRemoveItem_Click(object sender, EventArgs e)
        {
            if (dgvCart.CurrentRow != null)
            {
                dgvCart.Rows.RemoveAt(dgvCart.CurrentRow.Index);
                CalculateTotal();
            }
        }

        private void btnadd_Click_1(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtPriceOrder.Text))
            {
                MessageBox.Show("Select medicine first");
                return;
            }

            double price;
            double.TryParse(txtPriceOrder.Text, out price);

            int qty = Convert.ToInt32(numQuantityOrder.Value);

            double total = price * qty;

            purchaseCart.Rows.Add(
                cmbMedicineOrder.SelectedValue,
                cmbMedicineOrder.Text,
                price,
                qty,
                total
            );

            CalculateTotalOrder();
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbMedicineOrder.SelectedIndex != -1 && cmbMedicineOrder.SelectedValue != null && !(cmbMedicineOrder.SelectedValue is DataRowView))
            {
                try
                {
                    
                    int medId = Convert.ToInt32(cmbMedicineOrder.SelectedValue);

                    
                    string q = "select Price from Medicines where MedicineID = " + medId;

                    SqlDataAdapter da = new SqlDataAdapter(q, con);
                    DataTable dt = new DataTable();
                    da.Fill(dt);

                    if (dt.Rows.Count > 0)
                    {
                        // اسحب سعر التكلفة من الداتابيز
                        double costPrice = Convert.ToDouble(dt.Rows[0][0]);

                        // طبق هامش الربح بتاعك (25%)
                        double salePrice = costPrice * 0.75;

                        // اعرض السعر الجديد في التكست بوكس
                        txtPriceOrder.Text = salePrice.ToString();
                    }
                }
                catch (Exception ex)
                {
                    // لو حصل أي حاجة غلط متخليش البرنامج يقفل
                    // MessageBox.Show(ex.Message);
                }
            }
        }

        private void btnsave_Click(object sender, EventArgs e)
        {
            if (purchaseCart.Rows.Count == 0)
            {
                MessageBox.Show("Cart is empty");
                return;
            }

            con.Open();

            // 1. Insert Purchase Order (Header)
            string q1 = "insert into PurchaseOrders (SupplierID, OrderDate) values ("
                + cmbSupplier.SelectedValue + ", GETDATE())";

            SqlCommand cmd = new SqlCommand(q1, con);
            cmd.ExecuteNonQuery();

            // Get OrderID
            SqlCommand cmdId = new SqlCommand("select max(OrderID) from PurchaseOrders", con);
            int orderID = Convert.ToInt32(cmdId.ExecuteScalar());

            // 2. Loop Items
            foreach (DataRow row in purchaseCart.Rows)
            {
                // Insert Details
                string q2 = "insert into PurchaseDetails values ("
                    + orderID + ","
                    + row["MedicineID"] + ","
                    + row["Quantity"] + ","
                    + row["UnitPrice"] + ")";

                new SqlCommand(q2, con).ExecuteNonQuery();

                // 3. Update Stock (+ زيادة)
                string q3 = "update Medicines set Quantity = Quantity + "
                    + row["Quantity"] +
                    " where MedicineID=" + row["MedicineID"];

                new SqlCommand(q3, con).ExecuteNonQuery();
            }

            con.Close();

            MessageBox.Show("Purchase Order Saved");

            purchaseCart.Rows.Clear();
            txtTotal.Clear();
            LoadMedicines();
        }

        private void btnclearcart_Click(object sender, EventArgs e)
        {
            purchaseCart.Rows.Clear();
            txtTotalOrder.Clear();
        }

        private void cmbSupplierOrder_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void tabPage5_Click(object sender, EventArgs e)
        {

        }

        private void numQuantityOrder_ValueChanged(object sender, EventArgs e)
        {

        }

        private void txtPriceOrder_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTotalOrder_TextChanged(object sender, EventArgs e)
        {

        }

        private void cmbCustomer_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void cmbStaff_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void txtPricesal_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtTotal_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtFinalTotal_TextChanged(object sender, EventArgs e)
        {

        }

        private void numQuantity_ValueChanged(object sender, EventArgs e)
        {

        }

        private void txtMedicineName_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtPrice_TextChanged(object sender, EventArgs e)
        {

        }

        private void txtQuantity_TextChanged(object sender, EventArgs e)
        {

        }
    }
    }

