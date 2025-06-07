using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.Caching;
using System.Diagnostics;

namespace UCP1
{
    public partial class FormTambahTransaksi : Form
    {
        private int idEdit = -1;
        private int idKategoriEdit = -1;

        private string connectionString = "Data Source=MSI\\RIZKYPP;Initial Catalog=SisTemManajemenKeuangan;Integrated Security=True";

        private readonly MemoryCache _cache = MemoryCache.Default;
        private readonly CacheItemPolicy _policy = new CacheItemPolicy
        {
            AbsoluteExpiration = DateTimeOffset.Now.AddMinutes(5)
        };
        private const string CacheKey = "DataTransaksi";
        public FormTambahTransaksi()
        {
            InitializeComponent();
            dataGridView1.CellClick += dataGridView1_CellClick;


            button1.Text = "Lihat Laporan";
            button1.BackColor = Color.LightGreen;
            button1.Font = new Font("Arial", 9, FontStyle.Bold);
            button1.Click += BtnLihatLaporan_Click;
        }

        private void FormTambahTransaksi_Load(object sender, EventArgs e)
        {
            EnsureIndexes();
            LoadData();
        }
        
        private void EnsureIndexes()
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                var indexScript = @"
                    IF EXISTS (SELECT * FROM sys.indexes WHERE name = N'idx_Kategori_TipeKategori')
                    BEGIN
                        IF EXISTS (SELECT * FROM sys.columns 
                                   WHERE object_id = OBJECT_ID(N'dbo.Kategori') 
                                   AND object_id = OBJECT_ID 'dbo.kategori')
                        BEGIN
                            IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = N'idx_Kategori_TipeKategori')
                                CREATE NONCLUSTERED INDEX idx_Kategori_TipeKategori ON dbo.kategori(tipe);
                        END
                    END

                    IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = N'idx_Kategori_TipeKategori')
                        CREATE NONCLUSTERED INDEX idx_Kategori_TeipeKategori ON dbo.Kategori(tipe);
                ";

                using (var cmd = new SqlCommand(indexScript, conn))
                {
                    cmd.ExecuteReader();
                }
            }
        }

        private void LoadData()
        {
            DataTable dt;
            Stopwatch stopwatch = new Stopwatch();
            int rowCount = 0;

            if (_cache.Contains(CacheKey))
            {
                dt = _cache.Get(CacheKey) as DataTable;
                rowCount = dt?.Rows.Count ?? 0;
            }
            else
            {
                dt = new DataTable();
                stopwatch.Start(); // Mulai hitung waktu
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var query = @"
                        SELECT t.id_transaksi, k.nama_kategori, k.tipe, t.jumlah, t.tanggal
                        FROM transaksi t
                        JOIN kategori k ON t.id_kategori = k.id_kategori
                    WHERE k.tipe = 'pengeluaran'";
                    var da = new SqlDataAdapter(query, conn);
                    da.Fill(dt);
                    rowCount = dt.Rows.Count;
                }
                stopwatch.Stop(); // Selesai hitung waktu
                _cache.Add(CacheKey, dt, _policy);
            }

            dataGridView1.AutoGenerateColumns = true;
            dataGridView1.DataSource = dt;

            // Menampilkan informasi statistik
            string stats = $"STATISTICS INFO:\n" +
                           $"- Rows returned: {rowCount}\n" +
                           $"- Elapsed Time: {stopwatch.ElapsedMilliseconds} ms";

            MessageBox.Show(stats, "STATISTICS INFO");
        }


        private void AnalyzeQuery(string sqlQuery)
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.InfoMessage += (s, e) => MessageBox.Show(e.Message, "STATISTICS INFO");

                conn.Open();

                // Gabungkan query statistik dan query utama
                var wrapped = $@"
            SET NOCOUNT ON;
            SET STATISTICS IO ON;
            SET STATISTICS TIME ON;

            {sqlQuery}

            SET STATISTICS TIME OFF;
            SET STATISTICS IO OFF;";

                using (var cmd = new SqlCommand(wrapped, conn))
                {
                    cmd.ExecuteNonQuery(); // BUKAN ExecuteReader
                }
            }
        }



        private void BtnLihatLaporan_Click(object sender, EventArgs e)
        {
            try
            {

                button1.BackColor = Color.Lime;
                button1.Refresh();


                TampilkanLaporan();


                Timer timer = new Timer { Interval = 300 };
                timer.Tick += (s, args) =>
                {
                    button1.BackColor = Color.LightGreen;
                    timer.Stop();
                    timer.Dispose();
                };
                timer.Start();


                if (dataGridView1.Rows.Count > 0)
                {
                    dataGridView1.FirstDisplayedScrollingRowIndex = dataGridView1.Rows.Count - 1;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Gagal memuat laporan: {ex.Message}", "Error",
                              MessageBoxButtons.OK, MessageBoxIcon.Error);
                button1.BackColor = Color.Salmon;
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("pemasukan");
            comboBox1.Items.Add("pengeluaran");
            comboBox1.SelectedIndex = 0;
            dateTimePicker1.MinDate = new DateTime(2020, 1, 1);

            TampilkanLaporan();
        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            string nama = textBox1.Text; // Ini adalah nama_kategori
            string jumlahText = textBox2.Text.Trim();
            string keterangan = textBox3.Text;
            DateTime tanggal = dateTimePicker1.Value;
            string tipe = comboBox1.SelectedItem.ToString(); // Ini adalah tipe kategori

            // ... (Validasi input Anda tetap di sini) ...
            if (string.IsNullOrWhiteSpace(nama) || !System.Text.RegularExpressions.Regex.IsMatch(nama, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Nama kategori harus terdiri dari huruf saja.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            decimal jumlah;
            if (!decimal.TryParse(jumlahText, out jumlah) || jumlah <= 0)
            {
                MessageBox.Show("Jumlah harus berupa angka yang lebih dari 0.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (tanggal.Year < 2020)
            {
                MessageBox.Show("Data hanya berlaku untuk tahun 2020 ke atas.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var confirmSave = MessageBox.Show("Apakah Anda akan menyimpan data ini?",
                                             "Konfirmasi Simpan", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (confirmSave != DialogResult.Yes) return;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlTransaction sqlTransaction = conn.BeginTransaction(); // Mulai transaksi

                try
                {
                    if (idEdit == -1) // Mode Insert (Tambah Baru)
                    {
                        // Logika INSERT Anda yang sudah ada
                        string insertKategori = "INSERT INTO kategori (nama_kategori, tipe) VALUES (@nama, @tipe); SELECT SCOPE_IDENTITY();";
                        SqlCommand cmdKategori = new SqlCommand(insertKategori, conn, sqlTransaction);
                        cmdKategori.Parameters.AddWithValue("@nama", nama);
                        cmdKategori.Parameters.AddWithValue("@tipe", tipe);
                        int idKategoriBaru = Convert.ToInt32(cmdKategori.ExecuteScalar());

                        string insertTransaksi = "INSERT INTO transaksi (id_kategori, jumlah, tanggal, keterangan) VALUES (@id_kategori, @jumlah, @tanggal, @keterangan)";
                        SqlCommand cmdTransaksi = new SqlCommand(insertTransaksi, conn, sqlTransaction);
                        cmdTransaksi.Parameters.AddWithValue("@id_kategori", idKategoriBaru);
                        cmdTransaksi.Parameters.AddWithValue("@jumlah", jumlah);
                        cmdTransaksi.Parameters.AddWithValue("@tanggal", tanggal);
                        cmdTransaksi.Parameters.AddWithValue("@keterangan", keterangan);
                        cmdTransaksi.ExecuteNonQuery();

                        MessageBox.Show("Transaksi berhasil ditambahkan.");
                    }
                    else // Mode Update
                    {

                        // 1. Update tabel kategori
                        string updateKategoriQuery = "UPDATE kategori SET nama_kategori = @nama, tipe = @tipe WHERE id_kategori = @id_kategori_edit";
                        SqlCommand cmdUpdateKategori = new SqlCommand(updateKategoriQuery, conn, sqlTransaction);
                        cmdUpdateKategori.Parameters.AddWithValue("@nama", nama); // nama dari textBox1
                        cmdUpdateKategori.Parameters.AddWithValue("@tipe", tipe); // tipe dari comboBox1
                        cmdUpdateKategori.Parameters.AddWithValue("@id_kategori_edit", idKategoriEdit);
                        MessageBox.Show($"UPDATE:\nID: {idKategoriEdit}\nNama: {nama}\nTipe: {tipe}");


                        cmdUpdateKategori.ExecuteNonQuery();
                        int kategoriUpdated = cmdUpdateKategori.ExecuteNonQuery();
                        MessageBox.Show("Kategori rows updated: " + kategoriUpdated);


                        // 2. Update tabel transaksi
                        //    Catatan: id_kategori di tabel transaksi TIDAK diubah di sini,
                        //    karena kita mengedit detail kategori yang sama.
                        //    Jika Anda ingin transaksi ini dipindahkan ke kategori LAIN, logikanya akan berbeda.
                        string updateTransaksiQuery = "UPDATE transaksi SET jumlah=@jumlah, tanggal=@tanggal, keterangan=@keterangan WHERE id_transaksi=@id_transaksi_edit";
                        SqlCommand cmdUpdateTransaksi = new SqlCommand(updateTransaksiQuery, conn, sqlTransaction);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@jumlah", jumlah);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@tanggal", tanggal);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@keterangan", keterangan);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@id_transaksi_edit", idEdit); // idEdit adalah id_transaksi
                        cmdUpdateTransaksi.ExecuteNonQuery();

                        MessageBox.Show("Transaksi berhasil diupdate.");
                    }

                    sqlTransaction.Commit(); // Jika semua berhasil, commit transaksi
                }
                catch (Exception ex)
                {
                    sqlTransaction.Rollback(); // Jika ada error, batalkan semua perubahan
                    MessageBox.Show($"Error saat menyimpan data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally // Pastikan idEdit dan idKategoriEdit direset setelah operasi
                {
                    idEdit = -1;
                    idKategoriEdit = -1;
                }
            }

            ClearForm();
            TampilkanLaporan();
        }

        private void dataGridView1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            try
            {
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return; // Klik di header atau di luar batas

                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                // Abaikan jika yang diklik adalah baris TOTAL
                if (row.Cells["nama_kategori"].Value != null && row.Cells["nama_kategori"].Value.ToString() == "TOTAL")
                {
                    return;
                }


                string clickedColumnName = dataGridView1.Columns[e.ColumnIndex].Name;

                // Pastikan id_transaksi dan id_kategori ada dan valid sebelum melanjutkan
                if (row.Cells["id_transaksi"].Value == DBNull.Value || row.Cells["id_transaksi"].Value == null) return;
                int currentIdTransaksi = Convert.ToInt32(row.Cells["id_transaksi"].Value);


                if (clickedColumnName == "Update")
                {
                    var confirmUpdate = MessageBox.Show("Apakah Anda yakin ingin mengisi form dengan data ini untuk diupdate?",
                                                        "Konfirmasi Update", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                    if (confirmUpdate != DialogResult.Yes) return;

                    textBox1.Text = row.Cells["nama_kategori"].Value.ToString();
                    comboBox1.SelectedItem = row.Cells["tipe"].Value.ToString();
                    textBox2.Text = row.Cells["jumlah"].Value.ToString();
                    textBox3.Text = row.Cells["keterangan"].Value.ToString();

                    if (row.Cells["tanggal"].Value != DBNull.Value)
                    {
                        dateTimePicker1.Value = Convert.ToDateTime(row.Cells["tanggal"].Value);
                    }

                    idEdit = currentIdTransaksi;
                    idKategoriEdit = Convert.ToInt32(row.Cells["id_kategori"].Value); // ✅ Ini bagian penting

                    // ✅ Ini dia perbaikannya:
                    if (row.Cells["id_kategori"].Value != DBNull.Value && row.Cells["id_kategori"].Value != null)
                    {
                        idKategoriEdit = Convert.ToInt32(row.Cells["id_kategori"].Value);
                    }
                    else
                    {
                        idKategoriEdit = -1; // fallback
                    }
                }

                else if (clickedColumnName == "Delete")
                {
                    var confirmDelete = MessageBox.Show("Yakin ingin menghapus transaksi ini?", "Konfirmasi Hapus", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                    if (confirmDelete == DialogResult.Yes)
                    {
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            // Pertimbangkan menggunakan SqlTransaction jika proses delete melibatkan banyak tabel atau ada trigger kompleks
                            SqlCommand cmd = new SqlCommand("DELETE FROM transaksi WHERE id_transaksi = @id", conn);
                            cmd.Parameters.AddWithValue("@id", currentIdTransaksi);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Transaksi berhasil dihapus.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            TampilkanLaporan(); // Refresh data di grid
                            ClearForm();      // Bersihkan form jika data yang sama sedang diedit
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error pada operasi DataGridView: {ex.Message}\n\nStackTrace: {ex.StackTrace}", "Error Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
                // Saat error, pastikan ID reset agar tidak terjadi update/insert yang salah
                idEdit = -1;
                idKategoriEdit = -1;
            }
        }


        private void TampilkanLaporan()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"SELECT id_transaksi, id_kategori, nama_kategori, tipe, jumlah, 
                 CONVERT(varchar, tanggal, 103) as tanggal, keterangan 
                 FROM view_transaksi_lengkap 
                 ORDER BY tanggal DESC";


                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    decimal totalPemasukan = 0, totalPengeluaran = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["tipe"].ToString() == "pemasukan")
                            totalPemasukan += Convert.ToDecimal(row["jumlah"]);
                        else if (row["tipe"].ToString() == "pengeluaran")
                            totalPengeluaran += Convert.ToDecimal(row["jumlah"]);
                    }

                    DataRow totalRow = dt.NewRow();
                    totalRow["nama_kategori"] = "TOTAL";
                    totalRow["tipe"] = DBNull.Value;
                    totalRow["jumlah"] = DBNull.Value;
                    totalRow["tanggal"] = DBNull.Value;
                    totalRow["keterangan"] = $"Pemasukan: Rp{totalPemasukan:N0} | Pengeluaran: Rp{totalPengeluaran:N0} | Saldo: Rp{(totalPemasukan - totalPengeluaran):N0}";
                    dt.Rows.Add(totalRow);

                    SetupDataGridView(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saat menampilkan laporan: {ex.Message}");
            }
        }

        private void SetupDataGridView(DataTable dt)
        {
            dataGridView1.DataSource = null; // Hapus source lama
            dataGridView1.Columns.Clear();   // Hapus kolom lama
            dataGridView1.AutoGenerateColumns = false; // Nonaktifkan pembuatan kolom otomatis

            // Definisikan Kolom Data (sesuaikan DataPropertyName dengan nama kolom di DataTable/View)
            AddColumn("id_transaksi", "ID Transaksi", false); // Kolom ID biasanya disembunyikan
            AddColumn("id_kategori", "ID Kategori", false);  // Kolom ID Kategori juga disembunyikan
            AddColumn("nama_kategori", "Kategori");
            AddColumn("tipe", "Jenis");
            AddColumn("jumlah", "Jumlah", true, "N0", DataGridViewContentAlignment.MiddleRight); // Format angka
            AddColumn("tanggal", "Tanggal", true, "dd/MM/yyyy"); // Format tanggal, pastikan tipe data di DataTable sesuai
            AddColumn("keterangan", "Keterangan");

            // Definisikan Kolom Tombol Aksi
            AddButtonColumn("Update", "Update", Color.LightBlue);
            AddButtonColumn("Delete", "Delete", Color.LightCoral);

            dataGridView1.DataSource = dt; // Set source data baru

            FormatGridAppearance();
        }


        private void AddColumn(string dataPropertyName, string headerText, bool visible = true,
                               string format = null, DataGridViewContentAlignment alignment = DataGridViewContentAlignment.MiddleLeft)
        {
            var col = new DataGridViewTextBoxColumn
            {
                Name = dataPropertyName, // Name bisa sama dengan DataPropertyName agar mudah diakses
                DataPropertyName = dataPropertyName,
                HeaderText = headerText,
                Visible = visible,
                AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells, // Ukuran kolom
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = alignment }
            };

            if (!string.IsNullOrEmpty(format))
            {
                col.DefaultCellStyle.Format = format;
            }
            dataGridView1.Columns.Add(col);
        }

        private void AddButtonColumn(string name, string text, Color backColor)
        {
            var btnCol = new DataGridViewButtonColumn
            {
                Name = name,
                Text = text,
                HeaderText = "Aksi",
                UseColumnTextForButtonValue = true, // Tombol akan menampilkan Text di atas
                FlatStyle = FlatStyle.Flat,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = backColor,
                    ForeColor = Color.Black, // Warna teks tombol
                    Alignment = DataGridViewContentAlignment.MiddleCenter,
                    Padding = new Padding(2)
                },
                AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader // Ukuran kolom tombol
            };
            dataGridView1.Columns.Add(btnCol);
        }

        private void FormatGridAppearance()
        {
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 9, FontStyle.Bold);
            dataGridView1.EnableHeadersVisualStyles = false; // Memungkinkan kustomisasi header
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.Black;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill; // Sesuaikan ukuran kolom otomatis
            dataGridView1.RowHeadersVisible = false; // Sembunyikan header baris
            dataGridView1.AllowUserToAddRows = false; // Jangan biarkan user menambah baris langsung di grid
            dataGridView1.AllowUserToDeleteRows = false; // Jangan biarkan user menghapus baris langsung di grid
            dataGridView1.SelectionMode = DataGridViewSelectionMode.FullRowSelect; // Pilih seluruh baris
            dataGridView1.MultiSelect = false; // Hanya satu baris yang bisa dipilih

            // Format baris TOTAL jika ada
            if (dataGridView1.Rows.Count > 0)
            {
                DataGridViewRow lastRow = dataGridView1.Rows[dataGridView1.Rows.Count - 1];
                if (lastRow.Cells["nama_kategori"].Value != null && lastRow.Cells["nama_kategori"].Value.ToString() == "TOTAL")
                {
                    lastRow.DefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                    lastRow.DefaultCellStyle.BackColor = Color.LightYellow; // Warna latar beda untuk total
                    // Membuat kolom Update dan Delete tidak bisa diklik/kosong untuk baris TOTAL
                    lastRow.Cells["Update"].Value = string.Empty;
                    lastRow.Cells["Delete"].Value = string.Empty;
                }
            }
        }


        private void ClearForm()
        {
            textBox1.Clear();
            textBox2.Clear();
            textBox3.Clear();
            comboBox1.SelectedIndex = 0; // Atau -1 jika ingin kosong tanpa pilihan default
            dateTimePicker1.Value = DateTime.Now;
            idEdit = -1;        // Reset ID transaksi yang diedit
            idKategoriEdit = -1; // Reset ID kategori yang diedit
            textBox1.Focus();   // Fokuskan kembali ke input pertama
        }

        // Event untuk memilih file dan mempreview data
        private void button3_Click(object sender, EventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Excel Files|*.xlsx;*.xlsm;";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                string filePath = openFileDialog.FileName;
                PreviewData(filePath); // Display preview before importing
            }
        }

        // Method untuk menampilkan preview data di DataGridView
        private void PreviewData(string filePath)
        {
            try
            {
                using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fs); // Membuka workbook Excel
                    ISheet sheet = workbook.GetSheetAt(0);     // Mendapatkan worksheet pertama
                    DataTable dt = new DataTable();

                    // Membaca header kolom
                    IRow headerRow = sheet.GetRow(0);
                    foreach (var cell in headerRow.Cells)
                    {
                        dt.Columns.Add(cell.ToString());
                    }

                    // Membaca sisa data
                    for (int i = 1; i <= sheet.LastRowNum; i++) // Lewati baris header
                    {
                        IRow dataRow = sheet.GetRow(i);
                        DataRow newRow = dt.NewRow();
                        int cellIndex = 0;
                        foreach (var cell in dataRow.Cells)
                        {
                            newRow[cellIndex] = cell.ToString();
                            cellIndex++;
                        }
                        dt.Rows.Add(newRow);
                    }

                    // Membuka PreviewForm dan mengirimkan DataTable ke form tersebut
                    PreviewData previewForm = new PreviewData(dt);
                    previewForm.ShowDialog(); // Tampilkan PreviewForm
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error reading the Excel file: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void BtnAnalyze_Click(object sender, EventArgs e)
        {
            string sqlQuery = "SELECT id_kategori, nama_kategori, tipe FROM dbo.kategori WHERE nama_kategori LIKE 'A%'";

            AnalyzeQuery(sqlQuery); // Menampilkan statistik dari SQL Server

            // Jika mau juga tampilkan hasil data:
        }




    }
}