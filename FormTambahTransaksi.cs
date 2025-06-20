﻿using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using System.Runtime.Caching;
using System.Diagnostics;
using MathNet.Numerics;
using NPOI.POIFS.Crypt.Dsig;
using NPOI.SS.Formula.Functions;
using static NPOI.SS.Formula.PTG.ArrayPtg;

namespace UCP1
{
    public partial class FormTambahTransaksi : Form
    {
        private int idEdit = -1;
        private int idKategoriEdit = -1;

        private readonly string connectionString = "Data Source=PACARWELLY\\AULIANURFITRIA;Initial Catalog=KOAT;Integrated Security=True";

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
            LoadData(); // Memuat data awal untuk DataGridView

            // Pindahkan isi dari Form1_Load ke sini:
            comboBox1.Items.Add("pemasukan");
            comboBox1.Items.Add("pengeluaran");
            comboBox1.SelectedIndex = 0;
            dateTimePicker1.MinDate = DateTime.Now.AddYears(-5);
            dateTimePicker1.MaxDate = DateTime.Now;

            // Panggil TampilkanLaporan() di sini agar grid terisi saat form pertama kali dimuat
            TampilkanLaporan();
        }

        private void EnsureIndexes()
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.Open();
                var indexScript = @"
                    IF NOT EXISTS (SELECT * FROM sys.indexes WHERE name = N'idx_Kategori_TipeKategori' AND object_id = OBJECT_ID(N'dbo.Kategori'))
                    BEGIN
                        CREATE NONCLUSTERED INDEX idx_Kategori_TipeKategori ON dbo.Kategori(tipe);
                    END";
                using (var cmd = new SqlCommand(indexScript, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void LoadData()
        {
            DataTable dt;
            if (_cache.Contains(CacheKey))
            {
                dt = _cache.Get(CacheKey) as DataTable;
            }
            else
            {
                dt = new DataTable();
                using (var conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    var query = @"SELECT TOP 10 k.nama_kategori, k.tipe, t.jumlah FROM transaksi t JOIN kategori k ON t.id_kategori = k.id_kategori WHERE k.tipe = 'pengeluaran'";
                    var da = new SqlDataAdapter(query, conn);
                    da.Fill(dt);
                }
                _cache.Add(CacheKey, dt, _policy);
            }
        }

        private void AnalyzeQuery(string sqlQuery)
        {
            using (var conn = new SqlConnection(connectionString))
            {
                conn.InfoMessage += (s, e) => MessageBox.Show(e.Message, "STATISTICS INFO");
                conn.Open();
                var wrapped = $@"SET NOCOUNT ON; SET STATISTICS IO ON; SET STATISTICS TIME ON; {sqlQuery} SET STATISTICS TIME OFF; SET STATISTICS IO OFF;";
                using (var cmd = new SqlCommand(wrapped, conn))
                {
                    cmd.ExecuteNonQuery();
                }
            }
        }

        private void BtnLihatLaporan_Click(object sender, EventArgs e)
        {
            TampilkanLaporan();
        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            string nama = textBox1.Text.Trim();
            string jumlahText = textBox2.Text.Trim();
            string keterangan = textBox3.Text.Trim();
            DateTime tanggal = dateTimePicker1.Value;
            string tipe = comboBox1.SelectedItem.ToString();

            // --- INI BAGIAN PALING PENTING UNTUK MEMBLOKIR TANGGAL LAMA ---
            // 1. Tentukan batas tanggal paling lama yang diizinkan (5 tahun dari sekarang)
            DateTime batasTanggalMundur = DateTime.Now.AddYears(-5);

            // 2. Cek apakah tanggal yang diinput kurang dari batas tersebut
            if (tanggal.Date < batasTanggalMundur.Date)
            {
                // 3. Jika ya, tampilkan notifikasi dan hentikan proses dengan 'return'
                MessageBox.Show("Tanggal transaksi tidak boleh kurang dari 5 tahun terakhir.", "Tanggal Tidak Valid", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // PENTING: Menghentikan eksekusi, jadi tidak akan menyimpan data.
            }

            // 4. Cek juga agar tanggal tidak di masa depan
            if (tanggal.Date > DateTime.Now.Date)
            {
                MessageBox.Show("Tanggal transaksi tidak boleh melebihi tanggal hari ini.", "Tanggal Tidak Valid", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return; // PENTING: Menghentikan eksekusi.
            }
            // --- AKHIR DARI BLOK VALIDASI TANGGAL ---


            if (string.IsNullOrWhiteSpace(nama) || !System.Text.RegularExpressions.Regex.IsMatch(nama, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Nama kategori harus diisi dan hanya boleh terdiri dari huruf.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }
            if (!decimal.TryParse(jumlahText, out decimal jumlah) || jumlah <= 0)
            {
                MessageBox.Show("Jumlah harus berupa angka yang lebih dari 0.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // --- AWAL DARI BLOK VALIDASI TAMBAH/UPDATE ---
            DialogResult a;
            if (idEdit == -1) // Mode Insert
            {
                a = MessageBox.Show("Apakah Anda yakin ingin menambahkan data ini?", "Konfirmasi Tambah Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }
            else // Mode Update
            {
                a = MessageBox.Show("Apakah Anda yakin ingin memperbarui data ini?", "Konfirmasi Update Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            }

            if (a == DialogResult.No)
            {
                return; // Jika pengguna memilih 'No', batalkan operasi
            }
            // --- AKHIR DARI BLOK VALIDASI TAMBAH/UPDATE ---


            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlTransaction sqlTransaction = conn.BeginTransaction();
                try
                {
                    if (idEdit == -1) // Mode Insert
                    {
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
                    }
                    else // Mode Update
                    {
                        string updateKategoriQuery = "UPDATE kategori SET nama_kategori = @nama, tipe = @tipe WHERE id_kategori = @id_kategori_edit";
                        SqlCommand cmdUpdateKategori = new SqlCommand(updateKategoriQuery, conn, sqlTransaction);
                        cmdUpdateKategori.Parameters.AddWithValue("@nama", nama);
                        cmdUpdateKategori.Parameters.AddWithValue("@tipe", tipe);
                        cmdUpdateKategori.Parameters.AddWithValue("@id_kategori_edit", idKategoriEdit);
                        cmdUpdateKategori.ExecuteNonQuery();

                        string updateTransaksiQuery = "UPDATE transaksi SET jumlah=@jumlah, tanggal=@tanggal, keterangan=@keterangan WHERE id_transaksi=@id_transaksi_edit";
                        SqlCommand cmdUpdateTransaksi = new SqlCommand(updateTransaksiQuery, conn, sqlTransaction);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@jumlah", jumlah);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@tanggal", tanggal);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@keterangan", keterangan);
                        cmdUpdateTransaksi.Parameters.AddWithValue("@id_transaksi_edit", idEdit);
                        cmdUpdateTransaksi.ExecuteNonQuery();
                    }
                    sqlTransaction.Commit();
                    MessageBox.Show("Data berhasil disimpan.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
                catch (Exception ex)
                {
                    sqlTransaction.Rollback();
                    MessageBox.Show($"Error saat menyimpan data: {ex.Message}", "Database Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                }
                finally
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
                if (e.RowIndex < 0 || e.ColumnIndex < 0) return;
                DataGridViewRow row = dataGridView1.Rows[e.RowIndex];

                if (row.Cells["nama_kategori"].Value?.ToString() == "SALDO AKHIR") return;

                string clickedColumnName = dataGridView1.Columns[e.ColumnIndex].Name;
                if (row.Cells["id_transaksi"].Value == DBNull.Value) return;
                int currentIdTransaksi = Convert.ToInt32(row.Cells["id_transaksi"].Value);

                if (clickedColumnName == "Update")
                {
                    textBox1.Text = row.Cells["nama_kategori"].Value.ToString();
                    comboBox1.SelectedItem = row.Cells["tipe"].Value.ToString();
                    textBox2.Text = row.Cells["jumlah"].Value.ToString();
                    textBox3.Text = row.Cells["keterangan"].Value.ToString();
                    if (row.Cells["tanggal"].Value != DBNull.Value)
                    {
                        dateTimePicker1.Value = DateTime.ParseExact(row.Cells["tanggal"].Value.ToString(), "dd/MM/yyyy", null);
                    }
                    idEdit = currentIdTransaksi;
                    if (row.Cells["id_kategori"].Value != DBNull.Value)
                    {
                        idKategoriEdit = Convert.ToInt32(row.Cells["id_kategori"].Value);
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
                            SqlCommand cmd = new SqlCommand("DELETE FROM transaksi WHERE id_transaksi = @id", conn);
                            cmd.Parameters.AddWithValue("@id", currentIdTransaksi);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Transaksi berhasil dihapus.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                            TampilkanLaporan();
                            ClearForm();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error pada operasi DataGridView: {ex.ToString()}", "Error Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void TampilkanLaporan()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"SELECT id_transaksi, id_kategori, nama_kategori, tipe, jumlah, CONVERT(varchar, tanggal, 103) as tanggal, keterangan FROM view_transaksi_lengkap ORDER BY CONVERT(date, tanggal, 103) DESC, id_transaksi DESC";
                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    decimal totalPemasukan = 0;
                    decimal totalPengeluaran = 0;
                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["tipe"] == DBNull.Value || row["jumlah"] == DBNull.Value) continue;
                        string tipeTransaksi = row["tipe"].ToString().Trim();
                        decimal jumlah = Convert.ToDecimal(row["jumlah"]);
                        if (tipeTransaksi.Equals("pemasukan", StringComparison.OrdinalIgnoreCase)) totalPemasukan += jumlah;
                        else if (tipeTransaksi.Equals("pengeluaran", StringComparison.OrdinalIgnoreCase)) totalPengeluaran += jumlah;
                    }
                    DataRow totalRow = dt.NewRow();
                    totalRow["nama_kategori"] = "SALDO AKHIR";
                    totalRow["tipe"] = "Ringkasan";
                    totalRow["jumlah"] = totalPemasukan - totalPengeluaran;
                    totalRow["keterangan"] = $"Total Pemasukan: Rp{totalPemasukan:N0} | Total Pengeluaran: Rp{totalPengeluaran:N0}";
                    dt.Rows.Add(totalRow);
                    SetupDataGridView(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saat menampilkan laporan: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void SetupDataGridView(DataTable dt)
        {
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            dataGridView1.AutoGenerateColumns = false;
            AddColumn("id_transaksi", "ID Transaksi", false);
            AddColumn("id_kategori", "ID Kategori", false);
            AddColumn("nama_kategori", "Kategori");
            AddColumn("tipe", "Jenis");
            AddColumn("jumlah", "Jumlah", true, "N0", DataGridViewContentAlignment.MiddleRight);
            AddColumn("tanggal", "Tanggal");
            AddColumn("keterangan", "Keterangan");
            AddButtonColumn("Update", "Update", Color.LightBlue);
            AddButtonColumn("Delete", "Delete", Color.LightCoral);
            dataGridView1.DataSource = dt;
            FormatGridAppearance();
        }

        private void AddColumn(string dataPropertyName, string headerText, bool visible = true, string format = null, DataGridViewContentAlignment alignment = DataGridViewContentAlignment.MiddleLeft)
        {
            var col = new DataGridViewTextBoxColumn { Name = dataPropertyName, DataPropertyName = dataPropertyName, HeaderText = headerText, Visible = visible, AutoSizeMode = DataGridViewAutoSizeColumnMode.DisplayedCells, DefaultCellStyle = new DataGridViewCellStyle { Alignment = alignment, Format = format } };
            dataGridView1.Columns.Add(col);
        }

        private void AddButtonColumn(string name, string text, Color backColor)
        {
            var btnCol = new DataGridViewButtonColumn { Name = name, Text = text, HeaderText = "Aksi", UseColumnTextForButtonValue = true, FlatStyle = FlatStyle.Flat, DefaultCellStyle = new DataGridViewCellStyle { BackColor = backColor, ForeColor = Color.Black, Alignment = DataGridViewContentAlignment.MiddleCenter, Padding = new Padding(2) }, AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCellsExceptHeader };
            dataGridView1.Columns.Add(btnCol);
        }

        private void FormatGridAppearance()
        {
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 9, FontStyle.Bold);
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
            dataGridView1.RowHeadersVisible = false;

            if (dataGridView1.Rows.Count > 0)
            {
                DataGridViewRow lastRow = dataGridView1.Rows[dataGridView1.Rows.Count - 1];
                if (lastRow.Cells["nama_kategori"].Value?.ToString() == "SALDO AKHIR")
                {
                    lastRow.DefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                    lastRow.DefaultCellStyle.BackColor = Color.LightYellow;
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
            comboBox1.SelectedIndex = 0;
            dateTimePicker1.Value = DateTime.Now;
            idEdit = -1;
            idKategoriEdit = -1;
            textBox1.Focus();
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Tidak ada data untuk diekspor.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                sfd.FileName = $"Laporan_Transaksi_{DateTime.Now:dd-MM-yyyy}.xlsx";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    try
                    {
                        IWorkbook workbook = new XSSFWorkbook();
                        ISheet sheet = workbook.CreateSheet("Laporan Transaksi");

                        IRow headerRow = sheet.CreateRow(0);
                        int colIndex = 0;
                        for (int i = 0; i < dataGridView1.Columns.Count; i++)
                        {
                            if (dataGridView1.Columns[i].Visible && !(dataGridView1.Columns[i] is DataGridViewButtonColumn))
                            {
                                headerRow.CreateCell(colIndex).SetCellValue(dataGridView1.Columns[i].HeaderText);
                                colIndex++;
                            }
                        }

                        int rowIndex = 1;
                        foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                        {
                            IRow dataRow = sheet.CreateRow(rowIndex);
                            colIndex = 0;
                            for (int i = 0; i < dataGridView1.Columns.Count; i++)
                            {
                                if (dataGridView1.Columns[i].Visible && !(dataGridView1.Columns[i] is DataGridViewButtonColumn))
                                {
                                    string cellText = dgvRow.Cells[i].FormattedValue?.ToString() ?? "";
                                    dataRow.CreateCell(colIndex).SetCellValue(cellText);
                                    colIndex++;
                                }
                            }
                            rowIndex++;
                        }

                        using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                        {
                            workbook.Write(fs);
                        }
                        MessageBox.Show("Data berhasil diekspor ke file Excel.", "Ekspor Berhasil", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Terjadi error saat mengekspor data.\n\nDETAIL ERROR:\n{ex.ToString()}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }
            }
        }

        private void btnImportData_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog ofd = new OpenFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
            {
                if (ofd.ShowDialog() == DialogResult.OK)
                {
                    IWorkbook workbook;
                    using (FileStream fs = new FileStream(ofd.FileName, FileMode.Open, FileAccess.Read))
                    {
                        workbook = new XSSFWorkbook(fs);
                    }

                    ISheet sheet = workbook.GetSheetAt(0);
                    int berhasil = 0;
                    int gagal = 0;

                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        SqlTransaction transaction = conn.BeginTransaction();

                        try
                        {
                            for (int i = 1; i <= sheet.LastRowNum; i++)
                            {
                                IRow row = sheet.GetRow(i);
                                if (row == null) continue;

                                try
                                {
                                    string namaKategori = row.GetCell(0)?.ToString().Trim();
                                    string tipe = row.GetCell(1)?.ToString().Trim().ToLower();
                                    string jumlahStr = row.GetCell(2)?.ToString().Trim();
                                    string tanggalStr = row.GetCell(3)?.ToString().Trim();
                                    string keterangan = row.GetCell(4)?.ToString().Trim() ?? "";

                                    if (string.IsNullOrWhiteSpace(namaKategori) ||
                                        (tipe != "pemasukan" && tipe != "pengeluaran") ||
                                        !decimal.TryParse(jumlahStr, out decimal jumlah) ||
                                        !DateTime.TryParse(tanggalStr, out DateTime tanggal))
                                    {
                                        gagal++;
                                        continue;
                                    }

                                    int idKategori;
                                    string checkKategoriQuery = "SELECT id_kategori FROM kategori WHERE nama_kategori = @nama AND tipe = @tipe";
                                    SqlCommand cmdCheck = new SqlCommand(checkKategoriQuery, conn, transaction);
                                    cmdCheck.Parameters.AddWithValue("@nama", namaKategori);
                                    cmdCheck.Parameters.AddWithValue("@tipe", tipe);

                                    object result = cmdCheck.ExecuteScalar();

                                    if (result != null)
                                    {
                                        idKategori = Convert.ToInt32(result);
                                    }
                                    else
                                    {
                                        string insertKategoriQuery = "INSERT INTO kategori (nama_kategori, tipe) VALUES (@nama, @tipe); SELECT SCOPE_IDENTITY();";
                                        SqlCommand cmdInsertKategori = new SqlCommand(insertKategoriQuery, conn, transaction);
                                        cmdInsertKategori.Parameters.AddWithValue("@nama", namaKategori);
                                        cmdInsertKategori.Parameters.AddWithValue("@tipe", tipe);
                                        idKategori = Convert.ToInt32(cmdInsertKategori.ExecuteScalar());
                                    }

                                    string insertTransaksiQuery = "INSERT INTO transaksi (id_kategori, jumlah, tanggal, keterangan) VALUES (@id_kategori, @jumlah, @tanggal, @keterangan)";
                                    SqlCommand cmdInsertTransaksi = new SqlCommand(insertTransaksiQuery, conn, transaction);
                                    cmdInsertTransaksi.Parameters.AddWithValue("@id_kategori", idKategori);
                                    cmdInsertTransaksi.Parameters.AddWithValue("@jumlah", jumlah);
                                    cmdInsertTransaksi.Parameters.AddWithValue("@tanggal", tanggal);
                                    cmdInsertTransaksi.Parameters.AddWithValue("@keterangan", keterangan);

                                    cmdInsertTransaksi.ExecuteNonQuery();
                                    berhasil++;
                                }
                                catch
                                {
                                    gagal++;
                                }
                            }

                            transaction.Commit();
                            MessageBox.Show($"Impor Selesai.\n\nBerhasil: {berhasil} baris\nGagal: {gagal} baris", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            MessageBox.Show($"Terjadi error saat impor data. Semua perubahan dibatalkan.\n\nDETAIL: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
                    }

                    TampilkanLaporan();
                }
            }
        }

        private void btnReport_Click(object sender, EventArgs e)
        {
            FormReportExport fk = new FormReportExport();
            fk.Show();
        }
    }
}