using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Printing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using static System.Windows.Forms.VisualStyles.VisualStyleElement;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.Menu;

namespace 銷貨管理系統
{

    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

        }
        SqlConnection conn = new SqlConnection("Server=.;Database=ERPDB;Trusted_Connection=True;");
        DataSet ds = new DataSet();

        class CartItem
        {
            public string Customer { get; set; } // 新增：客戶名稱
            public string Payment { get; set; }
            public int ProductID { get; set; }
            public string Name { get; set; }
            public int Price { get; set; }
            public int Qty { get; set; }
            public override string ToString()
            {
                // 列表顯示：[顧客名](付款方式) 商品 x 數量
                return $"[{Customer}]({Payment}) {Name} x{Qty} = {Price * Qty}";
            }
        }

        List<CartItem> cart = new List<CartItem>();
        private void Form1_Load(object sender, EventArgs e)
        {
            // 初始化載入
            LoadData();
            LoadCustomerList();
            LoadSupplierData();
            LoadInventoryData();
            LoadSupplierList();
            LoadPurchaseData();
            // 設定 DataGridView 樣式
            SetupDataGridView(dataGridView1);
            SetupDataGridView(dataGridView2);
        }
        void LoadCustomerList()
        {
            try
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                // 假設你的客戶資料表叫 Customer，欄位有 CustomerName
                string sql = "SELECT CustomerName FROM Customer";
                SqlCommand cmd = new SqlCommand(sql, conn);
                SqlDataReader reader = cmd.ExecuteReader();

                comboBox2.Items.Clear(); // 假設紅框是 comboBox2
                while (reader.Read())
                {
                    comboBox2.Items.Add(reader["CustomerName"].ToString());
                }
                conn.Close();
            }
            catch (Exception ex) { MessageBox.Show("載入客戶名單失敗：" + ex.Message); }
        }
        void SetupDataGridView(DataGridView dgv)
        {
            if (dgv.Columns.Contains("ImagePath")) dgv.Columns["ImagePath"].Visible = false;
            if (dgv.Columns.Contains("Image"))
            {
                dgv.Columns["Image"].HeaderText = "圖片";
                dgv.RowTemplate.Height = 80;
                ((DataGridViewImageColumn)dgv.Columns["Image"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
            }
        }
        void LoadData()
        {
            if (this.DesignMode) return;
            try
            {
                // 1. 從資料庫撈取資料
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                if (ds.Tables.Contains("Product")) ds.Tables["Product"].Clear();

                SqlDataAdapter adapter = new SqlDataAdapter("SELECT * FROM Product", conn);
                adapter.Fill(ds, "Product");
                conn.Close();

                // 2. 確保有 Image 欄位
                if (!ds.Tables["Product"].Columns.Contains("Image"))
                {
                    ds.Tables["Product"].Columns.Add("Image", typeof(Image));
                }

                // 3. 轉換圖片路徑為圖片物件
                foreach (DataRow row in ds.Tables["Product"].Rows)
                {
                    try
                    {
                        string fileName = row["ImagePath"].ToString();
                        string fullPath = Path.Combine(Application.StartupPath, "images", fileName);

                        if (!string.IsNullOrEmpty(fileName) && File.Exists(fullPath))
                        {
                            using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
                            {
                                using (Image tempImg = Image.FromStream(fs))
                                {
                                    row["Image"] = new Bitmap(tempImg);
                                }
                            }
                        }
                    }
                    catch (Exception)
                    {
                        // 如果這張圖讀取失敗 (參數無效)，我們就跳過它，或給它一張預設圖
                        // 這樣程式才不會整台卡死
                        row["Image"] = null;
                    }
                }

                // 4. 重新綁定來源
                dataGridView1.DataSource = null; // 先斷開再連上，強制刷新
                dataGridView1.DataSource = ds.Tables["Product"];
                dataGridView2.DataSource = null;
                dataGridView2.DataSource = ds.Tables["Product"];

                SetupDataGridView(dataGridView1);
                SetupDataGridView(dataGridView2);
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料載入失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }
        void RefreshCart()
        {
            listBox1.Items.Clear();

            foreach (var item in cart)
            {
                listBox1.Items.Add(item);
            }
        }
        private void dataGridView1_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            // 避免點到標題列
            if (e.RowIndex < 0) return;

            // 判斷是不是點到圖片欄位
            if (dataGridView1.Columns[e.ColumnIndex].Name == "Image")
            {
                var cellValue = dataGridView1.Rows[e.RowIndex].Cells["Image"].Value;

                if (cellValue != null)
                {
                    Image img = (Image)cellValue;

                    FormPreview preview = new FormPreview(img);
                    preview.ShowDialog(); // 彈出視窗
                }
            }
        }
        private void button1_Click_1(object sender, EventArgs e)
        {
            if (comboBox2.SelectedIndex == -1)
            {
                MessageBox.Show("請先選擇客戶！");
                return;
            }

            // 假設付款方式存在 comboBox3
            string payMethod = comboBox3.Text;
            if (string.IsNullOrEmpty(payMethod)) payMethod = "未設定";

            cart.Add(new CartItem
            {
                Customer = comboBox2.Text,
                Payment = payMethod, // 存入付款方式
                Name = dataGridView1.CurrentRow.Cells["Name"].Value.ToString(),
                Price = Convert.ToInt32(dataGridView1.CurrentRow.Cells["Price"].Value),
                Qty = 1
            });
            RefreshCart();
        }

        private void tabControl1_SelectedIndexChanged(object sender, EventArgs e)
        {
            switch (tabControl1.SelectedIndex)
            {
                case 1: // 當切換到「產品」頁面 (通常是第 2 個分頁)
                    LoadData();
                    break;

                case 0: // 當切換回「購物車與結帳」頁面 (通常是第 1 個分頁)
                    RefreshCart();
                    break;

                case 2: // 如果你有第三個分頁「客戶」
                        // 處理客戶資料載入
                    break;
            }
            // 假設你的「客戶」分頁是第三個（索引值從 0 開始，所以是 2）
            if (tabControl1.SelectedIndex == 2)
            {
                LoadCustomerData();
            }
            // 假設「銷貨主檔」是在第 3 個分頁 (索引從 0 開始)
            if (tabControl1.SelectedTab.Text == "銷貨主檔、明細")
            {
                LoadSalesOrders();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (listBox1.SelectedIndex == -1) return;

            cart.RemoveAt(listBox1.SelectedIndex);
            RefreshCart();
        }
        void LoadCustomerData()
        {
            try
            {
                // 1. 確保連線是關閉的再開啟，避免衝突
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                // 2. 撰寫 SQL 指令，設定中文欄位名稱讓 DataGridView 顯示更漂亮
                string sql = "SELECT CustomerID as '編號', CustomerName as '名字', Region as '地區', PaymentMethod as '付款方式' FROM Customer";

                // 3. 使用 SqlDataAdapter 撈取資料
                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                // 4. 將結果餵給 DataGridView
                dataGridView3.DataSource = dt;

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取客戶資料失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }
        private string GetSelectedPayments()
        {
            List<string> selected = new List<string>();
            foreach (var item in checkedListBox1.CheckedItems)
            {
                selected.Add(item.ToString());
            }
            return string.Join(", ", selected); // 用逗號隔開
        }
        private void button3_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                string sql = "INSERT INTO Customer (CustomerName, Region, PaymentMethod) VALUES (@name, @region, @pay)";
                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@name", textBox2.Text.Trim());
                cmd.Parameters.AddWithValue("@region", comboBox1.Text);

                // 這裡調用上面的方法取得複選字串
                cmd.Parameters.AddWithValue("@pay", GetSelectedPayments());

                cmd.ExecuteNonQuery();
                conn.Close();
                LoadCustomerData();
                MessageBox.Show("新增成功");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); conn.Close(); }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(textBox1.Text)) { MessageBox.Show("請先選擇客戶"); return; }

                conn.Open();
                string sql = "UPDATE Customer SET CustomerName=@name, Region=@region, PaymentMethod=@pay WHERE CustomerID=@id";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@id", textBox1.Text);
                cmd.Parameters.AddWithValue("@name", textBox2.Text.Trim());
                cmd.Parameters.AddWithValue("@region", comboBox1.Text);
                // 使用複選字串
                cmd.Parameters.AddWithValue("@pay", GetSelectedPayments());

                cmd.ExecuteNonQuery();
                conn.Close();
                LoadCustomerData();
                MessageBox.Show("資料更新成功");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); conn.Close(); }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox1.Text)) return;
            if (MessageBox.Show("確定要刪除此客戶嗎？", "警告", MessageBoxButtons.YesNo) == DialogResult.No) return;

            try
            {
                conn.Open();
                string sql = "DELETE FROM Customer WHERE CustomerID=@id";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@id", textBox1.Text);
                cmd.ExecuteNonQuery();
                conn.Close();
                LoadCustomerData();
                textBox1.Clear(); textBox2.Clear(); // 刪除後清空
                MessageBox.Show("資料已刪除");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); if (conn.State == ConnectionState.Open) conn.Close(); }
        }

        private void dataGridView3_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView3.Rows[e.RowIndex];
                textBox1.Text = row.Cells["編號"].Value?.ToString();
                textBox2.Text = row.Cells["名字"].Value?.ToString();
                comboBox1.Text = row.Cells["地區"].Value?.ToString();

                // --- 處理複選還原 ---
                string payValue = row.Cells["付款方式"].Value?.ToString() ?? "";
                // 將字串拆解回陣列 (去除空白)
                string[] payments = payValue.Split(new[] { ", " }, StringSplitOptions.RemoveEmptyEntries);

                // 先把所有選項取消勾選
                for (int i = 0; i < checkedListBox1.Items.Count; i++)
                {
                    checkedListBox1.SetItemChecked(i, false);
                }

                // 根據拆解後的名稱進行勾選
                foreach (string p in payments)
                {
                    int index = checkedListBox1.Items.IndexOf(p.Trim());
                    if (index != -1)
                    {
                        checkedListBox1.SetItemChecked(index, true);
                    }
                }
            }
        }

        private void button6_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                // 基本 SQL 語法
                string sql = "SELECT CustomerID as '編號', CustomerName as '名字', Region as '地區', PaymentMethod as '付款方式' FROM Customer WHERE 1=1";

                // 動態增加查詢條件
                if (!string.IsNullOrWhiteSpace(textBox2.Text))
                {
                    sql += " AND CustomerName LIKE @name";
                }
                if (comboBox1.SelectedIndex != -1 && comboBox1.Text != "全部")
                {
                    sql += " AND Region = @region";
                }

                SqlCommand cmd = new SqlCommand(sql, conn);

                // 帶入參數避免 SQL 注入
                if (!string.IsNullOrWhiteSpace(textBox2.Text))
                    cmd.Parameters.AddWithValue("@name", "%" + textBox2.Text.Trim() + "%");
                if (comboBox1.SelectedIndex != -1 && comboBox1.Text != "全部")
                    cmd.Parameters.AddWithValue("@region", comboBox1.Text);

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                dataGridView3.DataSource = dt;
                conn.Close();
            }
            catch (Exception ex) { MessageBox.Show("查詢失敗：" + ex.Message); conn.Close(); }
        }
        PrintDocument printDoc = new PrintDocument();
        private void button7_Click(object sender, EventArgs e)
        {
            // 綁定列印事件
            printDoc.PrintPage += new PrintPageEventHandler(printDocument1_PrintPage);

            // 顯示列印預覽視窗 (選用)
            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = printDoc;
            preview.ShowDialog();
        }

        private void printDocument1_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font fontTitle = new Font("微軟正黑體", 16, FontStyle.Bold);
            Font fontBody = new Font("微軟正黑體", 10);
            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float lineHeight = fontBody.GetHeight();

            // 標題
            e.Graphics.DrawString("客戶資料報表", fontTitle, Brushes.Black, x, y);
            y += 40;

            // 畫欄位表頭
            e.Graphics.DrawString("編號\t姓名\t地區\t付款方式", fontBody, Brushes.Blue, x, y);
            e.Graphics.DrawLine(Pens.Black, x, y + lineHeight, e.MarginBounds.Right, y + lineHeight);
            y += lineHeight + 10;

            // 逐筆繪製 DataGridView 的內容
            foreach (DataGridViewRow row in dataGridView3.Rows)
            {
                if (row.IsNewRow) continue;

                string id = row.Cells["編號"].Value?.ToString() ?? "";
                string name = row.Cells["名字"].Value?.ToString() ?? "";
                string region = row.Cells["地區"].Value?.ToString() ?? "";
                string pay = row.Cells["付款方式"].Value?.ToString() ?? "";

                string line = $"{id}\t{name}\t{region}\t{pay}";
                e.Graphics.DrawString(line, fontBody, Brushes.Black, x, y);
                y += lineHeight;

                // 如果頁面滿了（簡單分頁判斷）
                if (y > e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }
            e.HasMorePages = false;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Filter = "圖片檔案 (*.jpg;*.png)|*.jpg;*.png";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                // 取得檔名並存入 Tag 或特定的 Label/TextBox
                string fileName = Path.GetFileName(ofd.FileName);
                textBox6.Text = fileName;

                string imgFolder = Path.Combine(Application.StartupPath, "images");
                if (!Directory.Exists(imgFolder)) Directory.CreateDirectory(imgFolder);

                string destPath = Path.Combine(imgFolder, fileName);
                if (!File.Exists(destPath)) File.Copy(ofd.FileName, destPath);

                // 預覽圖片 (假設你有一個 pictureBox1)
                pictureBox1.Image = Image.FromFile(ofd.FileName);
            }
        }

        private void button13_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(textBox4.Text) || string.IsNullOrWhiteSpace(textBox5.Text))
                {
                    MessageBox.Show("請輸入名稱與價格");
                    return;
                }

                conn.Open();
                string sql = "INSERT INTO Product (Name, Price, ImagePath) VALUES (@name, @price, @path)";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@name", textBox4.Text.Trim());
                cmd.Parameters.AddWithValue("@price", textBox5.Text.Trim());
                // 確保 textBox6 裡有檔名（例如：apple.jpg）
                cmd.Parameters.AddWithValue("@path", textBox6.Text.Trim());

                cmd.ExecuteNonQuery();
                conn.Close();

                // 清空 UI 輸入框
                textBox4.Clear(); textBox5.Clear(); textBox6.Clear();
                pictureBox1.Image = null;

                // 重要：重新載入資料（包含圖片轉換邏輯）
                LoadData();
                MessageBox.Show("商品新增成功");
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private void button12_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(textBox3.Text)) return;

                conn.Open();
                string sql = "UPDATE Product SET Name=@name, Price=@price, ImagePath=@path WHERE ProductID=@id";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@id", textBox3.Text);
                cmd.Parameters.AddWithValue("@name", textBox4.Text);
                cmd.Parameters.AddWithValue("@price", textBox5.Text);
                cmd.Parameters.AddWithValue("@path", textBox6.Text);

                cmd.ExecuteNonQuery();
                conn.Close();
                LoadData();
                MessageBox.Show("更新完成");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); conn.Close(); }
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox3.Text)) return;
            if (MessageBox.Show("確定要刪除此商品？", "警告", MessageBoxButtons.YesNo) == DialogResult.No) return;

            try
            {
                conn.Open();
                string sql = "DELETE FROM Product WHERE ProductID=@id";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@id", textBox3.Text);
                cmd.ExecuteNonQuery();
                conn.Close();
                LoadData();
                MessageBox.Show("商品已刪除");
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); conn.Close(); }
        }

        private void dataGridView2_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
                DataGridViewRow row = dataGridView2.Rows[e.RowIndex];

                // 使用資料庫原始欄位名稱 (ProductID, Name, Price) 較保險
                textBox3.Text = row.Cells["ProductID"].Value?.ToString(); // 編號
                textBox4.Text = row.Cells["Name"].Value?.ToString();      // 名稱
                textBox5.Text = row.Cells["Price"].Value?.ToString();     // 價格

                // 如果有圖片路徑文字框，也要一併帶入
                if (dataGridView2.Columns.Contains("ImagePath"))
                {
                    string imgName = row.Cells["ImagePath"].Value?.ToString() ?? "";
                    // txtImagePath.Text = imgName; // 假設你有這個框
                }
            }
        }

        private void dataGridView2_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            // 避免點到標題列
            if (e.RowIndex < 0) return;

            // 判斷是不是點到圖片欄位
            if (dataGridView2.Columns[e.ColumnIndex].Name == "Image")
            {
                var cellValue = dataGridView2.Rows[e.RowIndex].Cells["Image"].Value;

                if (cellValue != null)
                {
                    Image img = (Image)cellValue;

                    FormPreview preview = new FormPreview(img);
                    preview.ShowDialog(); // 彈出視窗
                }
            }
        }

        private void button10_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                // 1. 修改 SQL 語法：改為查詢 Product 資料表
                // 這裡將欄位名稱改為商品的 ProductID, Name, Price, ImagePath
                string sql = "SELECT ProductID as '商品編號', Name as '商品名稱', Price as '價格', ImagePath as '圖片路徑' FROM Product WHERE 1=1";

                // 2. 動態增加查詢條件：假設使用 textBox4 來搜尋商品名稱
                if (!string.IsNullOrWhiteSpace(textBox4.Text))
                {
                    sql += " AND Name LIKE @name";
                }

                SqlCommand cmd = new SqlCommand(sql, conn);

                // 3. 帶入參數
                if (!string.IsNullOrWhiteSpace(textBox4.Text))
                    cmd.Parameters.AddWithValue("@name", "%" + textBox4.Text.Trim() + "%");

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);
                DataTable dt = new DataTable();
                adapter.Fill(dt);

                // 4. 加入 Image 欄位並處理圖片顯示 (如同 LoadData 的邏輯)
                if (!dt.Columns.Contains("Image"))
                {
                    dt.Columns.Add("Image", typeof(Image));
                }

                foreach (DataRow row in dt.Rows)
                {
                    try
                    {
                        string fileName = row["圖片路徑"].ToString();
                        string fullPath = Path.Combine(Application.StartupPath, "images", fileName);

                        if (!string.IsNullOrEmpty(fileName) && File.Exists(fullPath))
                        {
                            using (FileStream fs = new FileStream(fullPath, FileMode.Open, FileAccess.Read))
                            {
                                using (Image tempImg = Image.FromStream(fs))
                                {
                                    row["Image"] = new Bitmap(tempImg);
                                }
                            }
                        }
                    }
                    catch { /* 忽略損壞的圖片 */ }
                }

                // 5. 顯示到商品頁面的 DataGridView (假設是 dataGridView2)
                dataGridView2.DataSource = dt;

                // 6. 隱藏文字路徑，設定圖片欄位
                if (dataGridView2.Columns.Contains("圖片路徑")) dataGridView2.Columns["圖片路徑"].Visible = false;
                if (dataGridView2.Columns.Contains("Image"))
                {
                    dataGridView2.Columns["Image"].HeaderText = "圖片";
                    dataGridView2.RowTemplate.Height = 80;
                    ((DataGridViewImageColumn)dataGridView2.Columns["Image"]).ImageLayout = DataGridViewImageCellLayout.Zoom;
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("商品查詢失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private void button14_Click(object sender, EventArgs e)
        {
            // 1. 檢查購物車是否有選中要被換掉的項目
            if (listBox1.SelectedItem == null)
            {
                MessageBox.Show("請先在右邊購物車選中要「被換掉」的項目");
                return;
            }

            // 2. 檢查商品表是否有選中新的商品
            if (dataGridView1.CurrentRow == null)
            {
                MessageBox.Show("請先在左邊商品表選中「新商品」");
                return;
            }

            // 3. 獲取新商品的資訊
            string newName = dataGridView1.CurrentRow.Cells["Name"].Value.ToString();
            int newPrice = Convert.ToInt32(dataGridView1.CurrentRow.Cells["Price"].Value);

            // 4. 獲取購物車中被選中的物件參考
            CartItem selectedItem = (CartItem)listBox1.SelectedItem;

            // 提示確認
            DialogResult result = MessageBox.Show($"確定要將 [{selectedItem.Name}] 更換為 [{newName}] 嗎？",
                                                  "更換確認", MessageBoxButtons.YesNo);

            if (result == DialogResult.Yes)
            {
                // 5. 直接修改該物件的屬性
                selectedItem.Name = newName;
                selectedItem.Price = newPrice;
                // 數量 (Qty) 通常保持不變，若要重設為 1 可取消註解下面這行
                // selectedItem.Qty = 1; 

                // 6. 刷新顯示
                RefreshCart();
                MessageBox.Show("商品已成功更換");
            }
        }

        private void button15_Click(object sender, EventArgs e)
        {
            if (dataGridView1.CurrentRow == null) return;

            // 1. 獲取 DataGridView 目前選中的商品名稱或 ID
            string selectedProductName = dataGridView1.CurrentRow.Cells["Name"].Value.ToString();

            // 2. 在購物車清單中尋找
            bool found = false;
            for (int i = 0; i < listBox1.Items.Count; i++)
            {
                CartItem item = (CartItem)listBox1.Items[i];
                if (item.Name == selectedProductName)
                {
                    listBox1.SelectedIndex = i; // 自動選中該列
                    found = true;
                    break;
                }
            }

            if (found)
            {
                MessageBox.Show($"已在購物車中找到：{selectedProductName}");
            }
            else
            {
                MessageBox.Show("該商品目前不在購物車中");
            }
        }

        private void button16_Click(object sender, EventArgs e)
        {
            if (cart.Count <= 1) return; // 沒東西或只有一筆就不需要排

            // 這裡示範「對話框選擇」或是「循環切換排序」
            // 做法：先依照客戶名稱排，如果客戶一樣，就依照商品名稱排
            cart = cart.OrderBy(x => x.Customer).ThenBy(x => x.Name).ToList();

            RefreshCart();
            MessageBox.Show("已完成排序");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            // 綁定列印事件
            printDoc.PrintPage += new PrintPageEventHandler(printDocument2_PrintPage);

            // 顯示列印預覽視窗 (選用)
            PrintPreviewDialog preview = new PrintPreviewDialog();
            preview.Document = printDoc;
            preview.ShowDialog();
        }

        private void printDocument2_PrintPage(object sender, PrintPageEventArgs e)
        {
            Font fontTitle = new Font("微軟正黑體", 16, FontStyle.Bold);
            Font fontBody = new Font("微軟正黑體", 10);
            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float lineHeight = fontBody.GetHeight();

            // 1. 修改標題
            e.Graphics.DrawString("商品資料報表", fontTitle, Brushes.Black, x, y);
            y += 40;

            // 2. 修改欄位表頭 (對齊商品資訊)
            e.Graphics.DrawString("編號\t商品名稱\t\t價格\t圖片路徑", fontBody, Brushes.Blue, x, y);
            e.Graphics.DrawLine(Pens.Black, x, y + lineHeight, e.MarginBounds.Right, y + lineHeight);
            y += lineHeight + 10;

            // 3. 逐筆繪製 DataGridView1 (商品) 的內容
            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                if (row.IsNewRow) continue;

                // 根據你 DataGridView 的欄位名稱抓取資料
                string id = row.Cells["ProductID"].Value?.ToString() ?? "";
                string name = row.Cells["Name"].Value?.ToString() ?? "";
                string price = row.Cells["Price"].Value?.ToString() ?? "";
                string path = row.Cells["ImagePath"].Value?.ToString() ?? "";

                // 4. 組合字串 (加上 \t 分隔)
                // 註：商品名稱後多加一個 \t 是為了排版對齊
                string line = $"{id}\t{name}\t\t{price}\t{path}";

                e.Graphics.DrawString(line, fontBody, Brushes.Black, x, y);
                y += lineHeight;

                // 5. 頁面分頁判斷
                if (y > e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }
            e.HasMorePages = false;
        }

        private void button17_Click(object sender, EventArgs e)
        {
            if (cart.Count == 0) return;

            StringBuilder sb = new StringBuilder();
            sb.AppendLine("========== 結帳收據 ==========");

            int totalQty = 0;
            int totalPrice = 0;

            // 按顧客群組顯示（如果購物車有多人）
            var groupedCart = cart.GroupBy(x => new { x.Customer, x.Payment });

            foreach (var group in groupedCart)
            {
                sb.AppendLine($"顧客：{group.Key.Customer}");
                sb.AppendLine($"付款方式：{group.Key.Payment}");
                sb.AppendLine("----------------------------");

                foreach (var item in group)
                {
                    sb.AppendLine($"- {item.Name} x{item.Qty}  價錢：${item.Price * item.Qty}");
                    totalQty += item.Qty;
                    totalPrice += (item.Price * item.Qty);
                }
                sb.AppendLine(" "); // 空行隔開不同顧客
            }

            sb.AppendLine("============================");
            sb.AppendLine($"總品項數：{cart.Count}");
            sb.AppendLine($"總數量：{totalQty}");
            sb.AppendLine($"應付總金額：${totalPrice}");
            sb.AppendLine("============================");

            // 顯示收據
            MessageBox.Show(sb.ToString(), "結帳明細");
            try
            {
                if (conn.State == ConnectionState.Closed) conn.Open();

                // 寫入主檔並取得 OrderID
                string sqlMaster = @"INSERT INTO SalesOrder (CustomerName, PaymentMethod, TotalAmount) 
                             OUTPUT INSERTED.OrderID 
                             VALUES (@cust, @pay, @total)";

                SqlCommand cmdMaster = new SqlCommand(sqlMaster, conn);
                // 這裡抓取第一筆項目的顧客資訊作為主檔依據
                cmdMaster.Parameters.AddWithValue("@cust", cart[0].Customer);
                cmdMaster.Parameters.AddWithValue("@pay", cart[0].Payment);
                cmdMaster.Parameters.AddWithValue("@total", totalPrice);

                int newOrderID = (int)cmdMaster.ExecuteScalar();

                // 逐筆寫入明細檔
                foreach (var item in cart)
                {
                    string sqlDetail = @"INSERT INTO SalesOrderDetail (OrderID, ProductID, ProductName, UnitPrice, Quantity) 
                                 VALUES (@oid, @pid, @pname, @price, @qty)";
                    SqlCommand cmdDetail = new SqlCommand(sqlDetail, conn);
                    cmdDetail.Parameters.AddWithValue("@oid", newOrderID);
                    cmdDetail.Parameters.AddWithValue("@pid", 0); // 若你有 ProductID 欄位請替換 item.ProductID
                    cmdDetail.Parameters.AddWithValue("@pname", item.Name);
                    cmdDetail.Parameters.AddWithValue("@price", item.Price);
                    cmdDetail.Parameters.AddWithValue("@qty", item.Qty);
                    cmdDetail.ExecuteNonQuery();
                }

                conn.Close();

                // 3. 寫入成功後，重新整理銷貨主檔的 DataGridView
                LoadSalesOrders();

                MessageBox.Show("資料已成功存入銷貨主檔！");
            }
            catch (Exception ex)
            {
                MessageBox.Show("資料庫存檔失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }

            // 4. 最後詢問是否清空購物車
            if (MessageBox.Show("交易完成，是否清空購物車？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                cart.Clear();
                RefreshCart();
            }
        }

        private void printDocument3_PrintPage(object sender, PrintPageEventArgs e)
        {
            // 設定字體
            Font fontTitle = new Font("微軟正黑體", 18, FontStyle.Bold);
            Font fontHeader = new Font("微軟正黑體", 12, FontStyle.Bold);
            Font fontBody = new Font("微軟正黑體", 12);

            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float lineHeight = fontBody.GetHeight() + 5;

            // 1. 報表標題
            e.Graphics.DrawString("購物車結帳收據", fontTitle, Brushes.Black, x, y);
            y += 50;

            // 2. 畫欄位表頭
            // 格式：顧客 | 付款方式 | 品項 | 數量 | 小計
            e.Graphics.DrawString("顧客\t付款\t品項\t\t數量\t小計", fontHeader, Brushes.Blue, x, y);
            e.Graphics.DrawLine(Pens.Black, x, y + lineHeight, e.MarginBounds.Right, y + lineHeight);
            y += lineHeight + 10;

            int totalQty = 0;
            int totalPrice = 0;

            // 3. 逐筆繪製購物車 (cart) 的內容
            foreach (var item in cart)
            {
                // 組合資料列字串
                // \t 是為了對齊，品項後加兩個 \t 是因為品項名稱通常較長
                string line = $"{item.Customer}\t{item.Payment}\t{item.Name}\t\t{item.Qty}\t${item.Price * item.Qty}";

                e.Graphics.DrawString(line, fontBody, Brushes.Black, x, y);
                y += lineHeight;

                // 累加統計資料
                totalQty += item.Qty;
                totalPrice += (item.Price * item.Qty);

                // 簡單分頁檢查
                if (y > e.MarginBounds.Bottom)
                {
                    e.HasMorePages = true;
                    return;
                }
            }

            // 4. 畫底線並顯示總計
            y += 20;
            e.Graphics.DrawLine(Pens.Black, x, y, e.MarginBounds.Right, y);
            y += 10;

            e.Graphics.DrawString($"總品項數：{cart.Count} 項", fontBody, Brushes.Black, x, y);
            y += lineHeight;
            e.Graphics.DrawString($"總計數量：{totalQty} 件", fontBody, Brushes.Black, x, y);
            y += lineHeight;
            e.Graphics.DrawString($"應付總額：NT$ {totalPrice}", fontTitle, Brushes.Red, x, y);

            e.HasMorePages = false;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            if (cart.Count == 0)
            {
                MessageBox.Show("購物車內無商品，無法列印收據。");
                return;
            }

            // 顯示列印預覽對話框
            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = printDocument3; // 連結剛才寫好邏輯的 PrintDocument
            ppd.ShowDialog();
        }
        void LoadSalesOrders()
        {
            string sql = "SELECT * FROM SalesOrder ORDER BY OrderDate DESC";
            SqlDataAdapter da = new SqlDataAdapter(sql, conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView4.DataSource = dt;
        }
       
        private void dataGridView4_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;
            try
            {
                // 2. 取得選中那一列的 OrderID
                // 確保你的 dataGridView4 裡確實有一個欄位叫 "OrderID"
                string orderId = dataGridView4.Rows[e.RowIndex].Cells["OrderID"].Value.ToString();

                // 3. 組合完整的 SQL 指令（記得在等號後加上 orderId）
                string sql = "SELECT * FROM SalesOrderDetail WHERE OrderID = " + orderId;

                // 4. 執行查詢
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                da.Fill(dt); // 這一次就不會報錯了，因為 SQL 指令現在是完整的

                // 5. 僅將結果顯示在明細表 (dataGridView5)
                // 不要去動 dataGridView4.DataSource，否則主檔清單會不見
                dataGridView5.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取明細失敗：" + ex.Message);
            }
            // 1. 確保點擊的是有效資料列
            if (e.RowIndex < 0) return;

            // 2. 從選中的 Row 抓取資料並放入 TextBox
            // 這裡的 Cells["名稱"] 必須對應你 SQL 查詢出的欄位名
            textBox7.Text = dataGridView4.Rows[e.RowIndex].Cells["OrderID"].Value.ToString();
            textBox8.Text = dataGridView4.Rows[e.RowIndex].Cells["CustomerName"].Value.ToString();
            textBox9.Text = dataGridView4.Rows[e.RowIndex].Cells["PaymentMethod"].Value.ToString();
            textBox10.Text = dataGridView4.Rows[e.RowIndex].Cells["TotalAmount"].Value.ToString();

            // 提示：如果你希望點擊主檔「順便」查詢明細，可以在這裡呼叫 LoadDetail
            // LoadSalesOrderDetails(txtOrderID.Text);
        }
        void LoadSalesOrderDetails(string orderId)
        {
            try
            {
                string sql = $"SELECT * FROM SalesOrderDetail WHERE OrderID = {orderId}";
                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);

                // 假設 dataGridView5 是您的明細表
                dataGridView5.DataSource = dt;
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入明細失敗：" + ex.Message);
            }
        }

        private void comboBox2_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Open) conn.Close();
                conn.Open();

                string sql = "SELECT PaymentMethod FROM Customer WHERE CustomerName = @name";
                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@name", comboBox2.Text);

                object result = cmd.ExecuteScalar();

                if (result != null)
                {
                    // 自動帶入付款方式（顯示在 comboBox3）
                    comboBox3.Text = result.ToString();
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取付款方式失敗：" + ex.Message);
            }
        }

        private void dataGridView5_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex < 0) return;

            // 將明細資料填入對應的 TextBox
            textBox11.Text = dataGridView5.Rows[e.RowIndex].Cells["DetailID"].Value.ToString();
            textBox12.Text = dataGridView5.Rows[e.RowIndex].Cells["ProductName"].Value.ToString();
            textBox13.Text = dataGridView5.Rows[e.RowIndex].Cells["UnitPrice"].Value.ToString();
            textBox14.Text = dataGridView5.Rows[e.RowIndex].Cells["Quantity"].Value.ToString();
        }

        private void button19_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox7.Text)) return;

            try
            {
                conn.Open();
                string sql = @"UPDATE SalesOrder 
                       SET CustomerName = @name, PaymentMethod = @pay 
                       WHERE OrderID = @id";

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@name", textBox8.Text);
                cmd.Parameters.AddWithValue("@pay", textBox9.Text);
                cmd.Parameters.AddWithValue("@id", textBox7.Text);

                cmd.ExecuteNonQuery();
                conn.Close();

                MessageBox.Show("主檔資料更新成功！");
                LoadSalesOrders(); // 重新整理 dataGridView4
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox7.Text)) return;
            if (MessageBox.Show("確定刪除此訂單及其所有明細？", "警告", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                conn.Open();
                SqlTransaction trans = conn.BeginTransaction();
                try
                {
                    // 必須先刪明細，再刪主檔
                    new SqlCommand($"DELETE FROM SalesOrderDetail WHERE OrderID={textBox7.Text}", conn, trans).ExecuteNonQuery();
                    new SqlCommand($"DELETE FROM SalesOrder WHERE OrderID={textBox7.Text}", conn, trans).ExecuteNonQuery();
                    trans.Commit();
                    MessageBox.Show("訂單已完全移除");
                }
                catch { trans.Rollback(); }
                conn.Close();
                LoadSalesOrders();
                dataGridView5.DataSource = null;
            }
        }

        private void button23_Click(object sender, EventArgs e)
        {
            try
            {
                conn.Open();
                // 1. 更新明細
                string sqlDetail = "UPDATE SalesOrderDetail SET Quantity=@q, UnitPrice=@p WHERE DetailID=@did";
                SqlCommand cmd1 = new SqlCommand(sqlDetail, conn);
                cmd1.Parameters.AddWithValue("@q", textBox14.Text);
                cmd1.Parameters.AddWithValue("@p", textBox13.Text);
                cmd1.Parameters.AddWithValue("@did", textBox11.Text);
                cmd1.ExecuteNonQuery();

                // 2. 重新計算該訂單總額並更新回主檔
                string sqlTotal = @"UPDATE SalesOrder SET TotalAmount = 
                            (SELECT SUM(UnitPrice * Quantity) FROM SalesOrderDetail WHERE OrderID=@oid) 
                            WHERE OrderID=@oid";
                SqlCommand cmd2 = new SqlCommand(sqlTotal, conn);
                cmd2.Parameters.AddWithValue("@oid", textBox11.Text);
                cmd2.ExecuteNonQuery();

                conn.Close();
                MessageBox.Show("明細已更新，總金額已同步計算");
                LoadSalesOrders(); // 刷新主表看新金額
                LoadSalesOrderDetails(textBox11.Text); // 刷新明細表
            }
            catch (Exception ex) { MessageBox.Show(ex.Message); conn.Close(); }
        }

        private void button24_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox11.Text)) return;
            conn.Open();
            // 1. 刪除該品項
            new SqlCommand($"DELETE FROM SalesOrderDetail WHERE DetailID={textBox11.Text}", conn).ExecuteNonQuery();

            // 2. 重新計算主檔總額 (處理刪除最後一項後變成 NULL 的情況)
            string sqlTotal = $@"UPDATE SalesOrder SET TotalAmount = 
                         ISNULL((SELECT SUM(UnitPrice * Quantity) FROM SalesOrderDetail WHERE OrderID={textBox11.Text}), 0) 
                         WHERE OrderID={textBox11.Text}";
            new SqlCommand(sqlTotal, conn).ExecuteNonQuery();

            conn.Close();
            LoadSalesOrders();
            LoadSalesOrderDetails(textBox11.Text);
        }

        private void button21_Click(object sender, EventArgs e)
        {
            try
            {
                // 使用 LIKE 進行模糊查詢
                string sql = @"SELECT * FROM SalesOrder 
                       WHERE CustomerName LIKE @search 
                       OR OrderID LIKE @search 
                       ORDER BY OrderID DESC";

                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                // 加上 % 符號代表只要包含該文字即可
                da.SelectCommand.Parameters.AddWithValue("@search", "%" + txtSearch.Text + "%");

                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView4.DataSource = dt;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("查無相關訂單資料。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢發生錯誤：" + ex.Message);
            }
        }

        private void printDocument4_PrintPage(object sender, PrintPageEventArgs e)
        {
            // 1. 基本設定
            Font fontTitle = new Font("微軟正黑體", 18, FontStyle.Bold);
            Font fontBody = new Font("微軟正黑體", 10);
            float x = e.MarginBounds.Left;
            float y = e.MarginBounds.Top;
            float lineHeight = fontBody.GetHeight() + 5;

            // 2. 標題與主檔資訊 (從 TextBox 抓取目前選中的資料)
            e.Graphics.DrawString("銷貨結帳單 (歷史紀錄)", fontTitle, Brushes.Black, x, y);
            y += 40;
            e.Graphics.DrawString($"訂單編號：{textBox7.Text}", fontBody, Brushes.Black, x, y);
            y += lineHeight;
            e.Graphics.DrawString($"顧客名稱：{textBox8.Text}", fontBody, Brushes.Black, x, y);
            y += lineHeight;
            e.Graphics.DrawString($"付款方式：{textBox9.Text}", fontBody, Brushes.Black, x, y);
            y += 20;

            // 3. 欄位標題
            e.Graphics.DrawLine(Pens.Black, x, y, e.MarginBounds.Right, y);
            y += 5;
            e.Graphics.DrawString("品項名稱", fontBody, Brushes.Black, x, y);
            e.Graphics.DrawString("單價", fontBody, Brushes.Black, x + 200, y);
            e.Graphics.DrawString("數量", fontBody, Brushes.Black, x + 300, y);
            e.Graphics.DrawString("小計", fontBody, Brushes.Black, x + 400, y);
            y += lineHeight;
            e.Graphics.DrawLine(Pens.Black, x, y, e.MarginBounds.Right, y);
            y += 10;

            // 4. 逐筆列印明細內容 (從 dataGridView5 抓取)
            foreach (DataGridViewRow row in dataGridView5.Rows)
            {
                if (row.IsNewRow) continue;

                string name = row.Cells["ProductName"].Value.ToString();
                string price = row.Cells["UnitPrice"].Value.ToString();
                string qty = row.Cells["Quantity"].Value.ToString();
                int subtotal = Convert.ToInt32(price) * Convert.ToInt32(qty);

                e.Graphics.DrawString(name, fontBody, Brushes.Black, x, y);
                e.Graphics.DrawString(price, fontBody, Brushes.Black, x + 200, y);
                e.Graphics.DrawString(qty, fontBody, Brushes.Black, x + 300, y);
                e.Graphics.DrawString(subtotal.ToString(), fontBody, Brushes.Black, x + 400, y);

                y += lineHeight;
            }

            // 5. 結尾總計
            y += 20;
            e.Graphics.DrawLine(Pens.Black, x, y, e.MarginBounds.Right, y);
            y += 10;
            e.Graphics.DrawString($"總計金額：NT$ {textBox10.Text}", fontTitle, Brushes.Red, x, y);
        }

        private void button22_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(textBox7.Text))
            {
                MessageBox.Show("請先選擇一筆訂單再列印。");
                return;
            }

            PrintPreviewDialog ppd = new PrintPreviewDialog();
            ppd.Document = printDocument4;
            ppd.ShowDialog();
        }

        private void button25_Click(object sender, EventArgs e)
        {
            try
            {
                // 使用 LIKE 進行模糊查詢
                string sql = @"SELECT * FROM SalesOrderDetail 
                       WHERE ProductName LIKE @search 
                       OR DetailID LIKE @search 
                       ORDER BY DetailID DESC";

                SqlDataAdapter da = new SqlDataAdapter(sql, conn);
                // 加上 % 符號代表只要包含該文字即可
                da.SelectCommand.Parameters.AddWithValue("@search", "%" + textBox15.Text + "%");

                DataTable dt = new DataTable();
                da.Fill(dt);
                dataGridView5.DataSource = dt;

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("查無相關訂單資料。");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢發生錯誤：" + ex.Message);
            }
        }

        private void button26_Click(object sender, EventArgs e)
        {
            // 1. 檢查必要欄位（例如客戶名稱不能為空）
            if (string.IsNullOrEmpty(textBox8.Text))
            {
                MessageBox.Show("請輸入客戶名稱");
                return;
            }

            try
            {
                conn.Open();
                // 2. 撰寫 SQL 新增指令
                string sql = @"INSERT INTO SalesOrder (CustomerName, PaymentMethod, TotalAmount, OrderDate) 
                       OUTPUT INSERTED.OrderID 
                       VALUES (@name, @pay, 0, GETDATE())"; // 新建單據總額先設為 0

                SqlCommand cmd = new SqlCommand(sql, conn);
                cmd.Parameters.AddWithValue("@name", textBox9.Text);
                cmd.Parameters.AddWithValue("@pay", textBox10.Text);

                // 3. 執行並取得新產生的 ID，直接顯示在 TextBox 供後續新增明細使用
                int newID = (int)cmd.ExecuteScalar();
                textBox7.Text = newID.ToString();

                conn.Close();

                MessageBox.Show($"主檔已建立！單號：{newID}。接下來請新增商品明細。");
                LoadSalesOrders(); // 重新整理 dataGridView4
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }

        private void button27_Click(object sender, EventArgs e)
        {
            // 1. 確保已經選中或建立了一筆主檔 ID
            if (string.IsNullOrEmpty(textBox11.Text))
            {
                MessageBox.Show("請先選擇或建立一筆銷貨主檔！");
                return;
            }

            try
            {
                conn.Open();
                // 2. 新增明細
                string sqlDetail = @"INSERT INTO SalesOrderDetail (OrderID, ProductID, ProductName, UnitPrice, Quantity) 
                             VALUES (@oid, 0, @pname, @price, @qty)";

                SqlCommand cmd1 = new SqlCommand(sqlDetail, conn);
                cmd1.Parameters.AddWithValue("@oid", textBox11.Text);
                cmd1.Parameters.AddWithValue("@pname", textBox12.Text);
                cmd1.Parameters.AddWithValue("@price", Convert.ToInt32(textBox13.Text));
                cmd1.Parameters.AddWithValue("@qty", Convert.ToInt32(textBox14.Text));
                cmd1.ExecuteNonQuery();

                // 3. 關鍵動作：同步更新主檔的總金額
                string sqlUpdateTotal = @"UPDATE SalesOrder SET TotalAmount = 
                                 (SELECT SUM(UnitPrice * Quantity) FROM SalesOrderDetail WHERE OrderID=@oid) 
                                 WHERE OrderID=@oid";
                SqlCommand cmd2 = new SqlCommand(sqlUpdateTotal, conn);
                cmd2.Parameters.AddWithValue("@oid", textBox11.Text);
                cmd2.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("品項新增完成，總金額已自動更新。");

                // 4. 重新整理畫面
                LoadSalesOrders(); // 刷新主檔 (看到新的總額)
                LoadSalesOrderDetails(textBox11.Text); // 刷新明細 (看到剛加的品項)
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增明細失敗：" + ex.Message);
                if (conn.State == ConnectionState.Open) conn.Close();
            }
        }
        void LoadSupplierData()
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"SELECT 
                       SupplierID as '供應商編號',
                       SupplierName as '供應商名稱',
                       ContactPerson as '聯絡人',
                       Phone as '電話',
                       Address as '地址'
                       FROM Supplier";

                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);

                DataTable dt = new DataTable();

                adapter.Fill(dt);

                dataGridView6.DataSource = dt;

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取供應商失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void dataGridView6_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 避免點到標題列
            if (e.RowIndex < 0) return;

            DataGridViewRow row = dataGridView6.Rows[e.RowIndex];

            textBox16.Text = row.Cells["供應商編號"].Value?.ToString();
            textBox17.Text = row.Cells["供應商名稱"].Value?.ToString();
            textBox18.Text = row.Cells["聯絡人"].Value?.ToString();
            textBox19.Text = row.Cells["電話"].Value?.ToString();
            textBox20.Text = row.Cells["地址"].Value?.ToString();
        }

        private void button28_Click(object sender, EventArgs e)
        {
            // 基本檢查
            if (string.IsNullOrWhiteSpace(textBox17.Text))
            {
                MessageBox.Show("請輸入供應商名稱");
                return;
            }

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"INSERT INTO Supplier
                       (SupplierName, ContactPerson, Phone, Address)
                       VALUES
                       (@name, @contact, @phone, @address)";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@name", textBox17.Text.Trim());
                cmd.Parameters.AddWithValue("@contact", textBox18.Text.Trim());
                cmd.Parameters.AddWithValue("@phone", textBox19.Text.Trim());
                cmd.Parameters.AddWithValue("@address", textBox20.Text.Trim());

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("供應商新增成功！");

                // 重新整理 DataGridView
                LoadSupplierData();

                // 清空欄位
                textBox16.Clear();
                textBox17.Clear();
                textBox18.Clear();
                textBox19.Clear();
                textBox20.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button29_Click(object sender, EventArgs e)
        {
            // 必須先選到供應商
            if (string.IsNullOrWhiteSpace(textBox16.Text))
            {
                MessageBox.Show("請先選擇供應商");
                return;
            }

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"UPDATE Supplier
                       SET SupplierName = @name,
                           ContactPerson = @contact,
                           Phone = @phone,
                           Address = @address
                       WHERE SupplierID = @id";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@id", textBox16.Text);
                cmd.Parameters.AddWithValue("@name", textBox17.Text.Trim());
                cmd.Parameters.AddWithValue("@contact", textBox18.Text.Trim());
                cmd.Parameters.AddWithValue("@phone", textBox19.Text.Trim());
                cmd.Parameters.AddWithValue("@address", textBox20.Text.Trim());

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("供應商資料更新成功！");

                // 重新整理畫面
                LoadSupplierData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button30_Click(object sender, EventArgs e)
        {
            // 必須先選擇供應商
            if (string.IsNullOrWhiteSpace(textBox16.Text))
            {
                MessageBox.Show("請先選擇供應商");
                return;
            }

            // 刪除確認
            DialogResult result = MessageBox.Show(
                "確定要刪除此供應商嗎？",
                "刪除確認",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Warning);

            if (result == DialogResult.No)
                return;

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = "DELETE FROM Supplier WHERE SupplierID = @id";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@id", textBox16.Text);

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("供應商已刪除");

                // 重新整理
                LoadSupplierData();

                // 清空欄位
                textBox16.Clear();
                textBox17.Clear();
                textBox18.Clear();
                textBox19.Clear();
                textBox20.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button31_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"SELECT
                       SupplierID as '供應商編號',
                       SupplierName as '供應商名稱',
                       ContactPerson as '聯絡人',
                       Phone as '電話',
                       Address as '地址'
                       FROM Supplier
                       WHERE SupplierName LIKE @search
                       OR ContactPerson LIKE @search
                       OR Phone LIKE @search";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue(
                    "@search",
                    "%" + textBox21.Text.Trim() + "%"
                );

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();

                adapter.Fill(dt);

                dataGridView6.DataSource = dt;

                conn.Close();

                // 查無資料提示
                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("查無資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        void LoadInventoryData()
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"SELECT
                       InventoryID as '庫存編號',
                       ProductName as '商品名稱',
                       StockQty as '庫存數量'
                       FROM Inventory";

                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);

                DataTable dt = new DataTable();

                adapter.Fill(dt);

                dataGridView7.DataSource = dt;

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取庫存失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void dataGridView7_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 避免點到標題列
            if (e.RowIndex < 0) return;

            DataGridViewRow row = dataGridView7.Rows[e.RowIndex];

            textBox22.Text = row.Cells["庫存編號"].Value?.ToString();

            textBox23.Text = row.Cells["商品名稱"].Value?.ToString();

            textBox24.Text = row.Cells["庫存數量"].Value?.ToString();
        }

        private void button33_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox22.Text))
            {
                MessageBox.Show("請先選擇商品");
                return;
            }

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"UPDATE Inventory
                       SET StockQty = @qty
                       WHERE InventoryID = @id";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@qty", textBox24.Text);

                cmd.Parameters.AddWithValue("@id", textBox22.Text);

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("庫存更新成功");

                LoadInventoryData();
            }
            catch (Exception ex)
            {
                MessageBox.Show("更新失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button34_Click(object sender, EventArgs e)
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"SELECT
                       InventoryID as '庫存編號',
                       ProductName as '商品名稱',
                       StockQty as '庫存數量'
                       FROM Inventory
                       WHERE ProductName  LIKE @search
                       OR InventoryID LIKE @search";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue(
                    "@search",
                    "%" + textBox25.Text.Trim() + "%"
                );

                SqlDataAdapter adapter = new SqlDataAdapter(cmd);

                DataTable dt = new DataTable();

                adapter.Fill(dt);

                dataGridView7.DataSource = dt;

                conn.Close();

                if (dt.Rows.Count == 0)
                {
                    MessageBox.Show("查無資料");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("查詢失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button35_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(textBox22.Text))
            {
                MessageBox.Show("請先選擇庫存資料");
                return;
            }

            DialogResult result = MessageBox.Show(
                "確定刪除這筆庫存？",
                "警告",
                MessageBoxButtons.YesNo
            );

            if (result == DialogResult.No)
                return;

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = "DELETE FROM Inventory WHERE InventoryID=@id";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@id", textBox22.Text);

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("刪除成功");

                LoadInventoryData();

                textBox22.Clear();
                textBox23.Clear();
                textBox24.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("刪除失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button32_Click(object sender, EventArgs e)
        {
            // 基本檢查
            if (string.IsNullOrWhiteSpace(textBox23.Text))
            {
                MessageBox.Show("請輸入商品名稱");
                return;
            }

            if (string.IsNullOrWhiteSpace(textBox24.Text))
            {
                MessageBox.Show("請輸入庫存數量");
                return;
            }

            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"INSERT INTO Inventory
                       (ProductName, StockQty)
                       VALUES
                       (@name, @qty)";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue("@name", textBox23.Text.Trim());

                cmd.Parameters.AddWithValue("@qty", textBox24.Text.Trim());

                cmd.ExecuteNonQuery();

                conn.Close();

                MessageBox.Show("庫存新增成功");

                // 重新整理
                LoadInventoryData();

                // 清空
                textBox22.Clear();
                textBox23.Clear();
                textBox24.Clear();
            }
            catch (Exception ex)
            {
                MessageBox.Show("新增失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        void LoadSupplierList()
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();
                conn.Open();
                string sql = "SELECT SupplierName FROM Supplier";

                SqlCommand cmd = new SqlCommand(sql, conn);

                SqlDataReader reader = cmd.ExecuteReader();

                comboBox4.Items.Clear();

                while (reader.Read())
                {
                    comboBox4.Items.Add(
                        reader["SupplierName"].ToString()
                    );
                }

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("載入供應商失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
        void LoadPurchaseData()
        {
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"SELECT
                       PurchaseID as '採購單號',
                       SupplierName as '供應商',
                       PurchaseDate as '採購日期',
                       TotalAmount as '總金額'
                       FROM PurchaseOrder
                       ORDER BY PurchaseDate DESC";

                SqlDataAdapter adapter = new SqlDataAdapter(sql, conn);

                DataTable dt = new DataTable();

                adapter.Fill(dt);

                dataGridView8.DataSource = dt;

                conn.Close();
            }
            catch (Exception ex)
            {
                MessageBox.Show("讀取採購資料失敗：" + ex.Message);

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }

        private void button36_Click(object sender, EventArgs e)
        {
            // 檢查供應商
            if (comboBox4.SelectedIndex == -1)
            {
                MessageBox.Show("請選擇供應商");
                return;
            }
            try
            {
                if (conn.State == ConnectionState.Open)
                    conn.Close();

                conn.Open();

                string sql = @"INSERT INTO PurchaseOrder
                       (SupplierName, PurchaseDate, TotalAmount)
                       OUTPUT INSERTED.PurchaseID
                       VALUES
                       (@supplier, GETDATE(), @total)";

                SqlCommand cmd = new SqlCommand(sql, conn);

                cmd.Parameters.AddWithValue(
                    "@supplier",
                    comboBox4.Text
                );

                // 新建採購單時先給 0
                cmd.Parameters.AddWithValue(
                    "@total",
                    0
                );

                // 取得新採購單號
                int newPurchaseID =
                    (int)cmd.ExecuteScalar();

                conn.Close();

                // 顯示單號
                textBox26.Text =
                    newPurchaseID.ToString();

                textBox27.Text = "0";

                MessageBox.Show(
                    $"採購主檔建立成功！\n採購單號：{newPurchaseID}"
                );

                // 重新整理
                LoadPurchaseData();
            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    "新增採購主檔失敗：" + ex.Message
                );

                if (conn.State == ConnectionState.Open)
                    conn.Close();
            }
        }
    }
}
