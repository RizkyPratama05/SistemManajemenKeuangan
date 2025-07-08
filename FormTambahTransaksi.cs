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

using MathNet.Numerics;

using NPOI.POIFS.Crypt.Dsig;

using NPOI.SS.Formula.Functions;

using static NPOI.SS.Formula.PTG.ArrayPtg;

using System.Collections.Generic;

using System.Text.RegularExpressions;
using System.Linq;
using System.Drawing;




namespace UCP1

{

    public partial class FormTambahTransaksi : Form

    {
        Koneksi kn = new Koneksi();
        string strKonek = "";

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
            strKonek = kn.connectionString();

            dataGridView1.CellClick += dataGridView1_CellClick;

            this.Load += FormTambahTransaksi_Load;



            button1.Text = "Lihat Laporan";

            button1.Font = new Font("Arial", 9, FontStyle.Bold);

            button1.Click += BtnLihatLaporan_Click;

        }



        private void FormTambahTransaksi_Load(object sender, EventArgs e)

        {

            EnsureIndexes();

            LoadData();

            comboBox1.SelectedIndex = 0;

            dateTimePicker1.MinDate = DateTime.Now.AddYears(-5);

            dateTimePicker1.MaxDate = DateTime.Now;



            // 1. Tampilkan dulu data ke DataGridView

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
            SqlConnection conn = new SqlConnection(strKonek);
            SqlDataAdapter da = new SqlDataAdapter("SELECT * FROM Transaksi", conn);
            DataTable dt = new DataTable();
            da.Fill(dt);
            dataGridView1.DataSource = dt;

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
            List<string> errors = new List<string>();

            string nama = textBox1.Text.Trim();
            string jumlahText = textBox2.Text.Trim();
            string keterangan = textBox3.Text.Trim();
            string tipe = comboBox1.Text.Trim();
            DateTime tanggal = dateTimePicker1.Value;

            // Validasi nama
            if (string.IsNullOrWhiteSpace(nama))
                errors.Add("Kolom 'Nama' wajib diisi.");
            else if (!Regex.IsMatch(nama, @"^[a-zA-Z\s]+$"))
                errors.Add("Kolom 'Nama' hanya boleh berisi huruf dan spasi.");

            // Validasi jumlah
            decimal jumlah = 0;
            if (string.IsNullOrWhiteSpace(jumlahText))
                errors.Add("Kolom 'Jumlah' wajib diisi.");
            else if (!decimal.TryParse(jumlahText, out jumlah))
                errors.Add("Kolom 'Jumlah' harus berupa angka.");
            else if (jumlah <= 0)
                errors.Add("Kolom 'Jumlah' harus lebih dari 0.");

            // Validasi kategori
            if (string.IsNullOrWhiteSpace(tipe))
                errors.Add("Kolom 'Kategori' wajib dipilih.");
            else if (tipe != "Pemasukan" && tipe != "Pengeluaran")
                errors.Add("Kolom 'Kategori' hanya boleh 'Pemasukan' atau 'Pengeluaran'.");

            // Validasi keterangan
            if (string.IsNullOrWhiteSpace(keterangan))
                errors.Add("Kolom 'Keterangan' wajib diisi.");

            // Validasi tanggal
            DateTime batasTanggalMundur = DateTime.Now.AddYears(-5);
            if (tanggal.Date < batasTanggalMundur.Date)
                errors.Add("Tanggal transaksi tidak boleh lebih dari 5 tahun lalu.");
            if (tanggal.Date > DateTime.Now.Date)
                errors.Add("Tanggal transaksi tidak boleh di masa depan.");

            // Tampilkan semua error sekaligus
            if (errors.Count > 0)
            {
                string errorMessage = string.Join("\n", errors);
                MessageBox.Show(errorMessage, "Validasi Input", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Konfirmasi Tambah/Update
            DialogResult a = (idEdit == -1)
                ? MessageBox.Show("Apakah Anda yakin ingin menambahkan data ini?", "Konfirmasi Tambah Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question)
                : MessageBox.Show("Apakah Anda yakin ingin memperbarui data ini?", "Konfirmasi Update Data", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (a == DialogResult.No) return;

            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();
                SqlTransaction sqlTransaction = conn.BeginTransaction();
                try
                {
                    if (idEdit == -1) // Insert
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
                    else // Update
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

                if (!dataGridView1.Columns.Contains("Pilih"))
                {
                    DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn();
                    chk.HeaderText = "Pilih";
                    chk.Name = "Pilih";
                    chk.Width = 50;
                    dataGridView1.Columns.Insert(0, chk);
                }

                if (clickedColumnName == "Update")
                {
                    textBox1.Text = row.Cells["nama_kategori"].Value.ToString();
                    comboBox1.SelectedItem = row.Cells["tipe"].Value.ToString();
                    textBox2.Text = row.Cells["jumlah"].Value.ToString();
                    textBox3.Text = row.Cells["keterangan"].Value.ToString();

                    if (row.Cells["tanggal"].Value != DBNull.Value)
                    {
                        DateTime parsedDate;
                        // Coba parsing dengan "dd/MM/yyyy" terlebih dahulu, karena kueri SQL Anda menggunakan 103 untuk format tanggal.
                        if (DateTime.TryParseExact(row.Cells["tanggal"].Value.ToString(), "dd/MM/yyyy", null, System.Globalization.DateTimeStyles.None, out parsedDate))
                        {
                            // **Perubahan Penting Di Sini:** Ganti sementara MaxDate untuk memungkinkan pengaturan tanggal valid apa pun untuk pembaruan.
                            // Validasi Anda di btnSimpan_Click akan tetap menangkap tanggal yang tidak valid saat menyimpan.
                            DateTime originalMaxDate = dateTimePicker1.MaxDate; // Simpan yang asli
                            DateTime originalMinDate = dateTimePicker1.MinDate; // Simpan yang asli

                            dateTimePicker1.MaxDate = DateTimePicker.MaximumDateTime; // Izinkan tanggal apa pun untuk diatur
                            dateTimePicker1.MinDate = DateTimePicker.MinimumDateTime; // Izinkan tanggal apa pun untuk diatur

                            dateTimePicker1.Value = parsedDate;

                            // Pulihkan tanggal Min/Max asli setelah mengatur nilai
                            dateTimePicker1.MaxDate = originalMaxDate;
                            dateTimePicker1.MinDate = originalMinDate;
                        }
                        else
                        {
                            // Fallback ke Parse umum jika format spesifik gagal (meskipun 103 harus konsisten)
                            // Jika parsing gagal lagi, Anda mungkin ingin mencatatnya atau mengatur tanggal default.
                            // Untuk saat ini, mari kita asumsikan seharusnya selalu parsing dengan benar mengingat CONVERT(varchar, tanggal, 103)
                            // Anda bahkan dapat mempertimbangkan hanya Convert.ToDateTime jika kolom dasar adalah tipe datetime secara langsung
                            try
                            {
                                parsedDate = Convert.ToDateTime(row.Cells["tanggal"].Value);

                                DateTime originalMaxDate = dateTimePicker1.MaxDate;
                                DateTime originalMinDate = dateTimePicker1.MinDate;

                                dateTimePicker1.MaxDate = DateTimePicker.MaximumDateTime;
                                dateTimePicker1.MinDate = DateTimePicker.MinimumDateTime;

                                dateTimePicker1.Value = parsedDate;

                                dateTimePicker1.MaxDate = originalMaxDate;
                                dateTimePicker1.MinDate = originalMinDate;
                            }
                            catch (Exception dateConvertEx)
                            {
                                MessageBox.Show($"Tidak dapat mengonversi tanggal '{row.Cells["tanggal"].Value}' dari database: {dateConvertEx.Message}", "Kesalahan Konversi Tanggal", MessageBoxButtons.OK, MessageBoxIcon.Error);
                                // Secara opsional atur nilai default atau kosongkan pemilih tanggal
                                // dateTimePicker1.Value = DateTime.Now;
                            }
                        }
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
                MessageBox.Show($"Error pada operasi DataGridView: {ex.Message}", "Error Grid", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void AturTampilanDataGridView()
        {
            dataGridView1.BorderStyle = System.Windows.Forms.BorderStyle.Fixed3D;
            dataGridView1.CellBorderStyle = DataGridViewCellBorderStyle.SingleHorizontal;
            dataGridView1.DefaultCellStyle.SelectionBackColor = Color.LightSteelBlue;
            dataGridView1.DefaultCellStyle.SelectionForeColor = Color.Black;
            dataGridView1.EnableHeadersVisualStyles = false;
            dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.SteelBlue;
            dataGridView1.ColumnHeadersDefaultCellStyle.ForeColor = Color.White;
            dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Segoe UI", 10F, FontStyle.Bold);
            dataGridView1.RowTemplate.Height = 28;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
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
                             ORDER BY CONVERT(date, tanggal, 103) DESC, id_transaksi DESC";

                    SqlDataAdapter adapter = new SqlDataAdapter(query, conn);
                    DataTable dt = new DataTable();
                    adapter.Fill(dt);

                    decimal totalPemasukan = 0;
                    decimal totalPengeluaran = 0;

                    foreach (DataRow row in dt.Rows)
                    {
                        if (row["tipe"] == DBNull.Value || row["jumlah"] == DBNull.Value) continue;

                        string tipe = row["tipe"].ToString().Trim().ToLower();
                        decimal jumlah = Convert.ToDecimal(row["jumlah"]);

                        if (tipe == "pemasukan")
                            totalPemasukan += jumlah;
                        else if (tipe == "pengeluaran")
                            totalPengeluaran += jumlah;
                    }

                    // Tambah baris saldo akhir
                    DataRow totalRow = dt.NewRow();
                    totalRow["nama_kategori"] = "SALDO AKHIR";
                    totalRow["tipe"] = "Ringkasan";
                    totalRow["jumlah"] = totalPemasukan - totalPengeluaran;
                    totalRow["keterangan"] = $"Total Pemasukan: Rp{totalPemasukan:N0} | Total Pengeluaran: Rp{totalPengeluaran:N0}";
                    dt.Rows.Add(totalRow);

                    // Tampilkan ke DataGridView
                    SetupDataGridView(dt);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saat menampilkan laporan: {ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            // Tambahkan kolom checkbox jika belum ada
            if (!dataGridView1.Columns.Contains("Pilih"))
            {
                DataGridViewCheckBoxColumn chk = new DataGridViewCheckBoxColumn
                {
                    HeaderText = "Pilih",
                    Name = "Pilih",
                    Width = 50
                };
                dataGridView1.Columns.Insert(0, chk);
            }

            // Hapus checkbox header sebelumnya jika ada
            var existingHeader = dataGridView1.Controls.Find("chkPilihSemua", true);
            if (existingHeader.Length > 0)
            {
                dataGridView1.Controls.Remove(existingHeader[0]);
            }

            // Tambahkan checkbox di header kolom "Pilih"
            Rectangle rect = dataGridView1.GetCellDisplayRectangle(0, -1, true);
            CheckBox chkHeader = new CheckBox
            {
                Name = "chkPilihSemua",
                Size = new Size(18, 18),
                Location = new Point(rect.Location.X + (rect.Width - 18) / 2, rect.Location.Y + 4),
                BackColor = Color.Transparent
            };

            chkHeader.CheckedChanged += (s, e) =>
            {
                foreach (DataGridViewRow row in dataGridView1.Rows)
                {
                    if (!row.IsNewRow)
                    {
                        row.Cells["Pilih"].Value = chkHeader.Checked;
                    }
                }
            };

            dataGridView1.Controls.Add(chkHeader);
        }
       


        private void btnHapusTerpilih_Click(object sender, EventArgs e)
        {
            List<int> idTransaksiTerpilih = new List<int>();

            foreach (DataGridViewRow row in dataGridView1.Rows)
            {
                DataGridViewCheckBoxCell chk = row.Cells["pilih"] as DataGridViewCheckBoxCell;
                if (chk != null && chk.Value != null && (bool)chk.Value == true)
                {
                    if (row.Cells["id_transaksi"].Value != DBNull.Value)
                    {
                        idTransaksiTerpilih.Add(Convert.ToInt32(row.Cells["id_transaksi"].Value));
                    }
                }
            }

            if (idTransaksiTerpilih.Count == 0)
            {
                MessageBox.Show("Tidak ada data yang dipilih untuk dihapus.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult result = MessageBox.Show($"Yakin ingin menghapus {idTransaksiTerpilih.Count} data terpilih?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Warning);

            if (result == DialogResult.Yes)
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    SqlTransaction tr = conn.BeginTransaction();

                    try
                    {
                        foreach (int id in idTransaksiTerpilih)
                        {
                            SqlCommand cmd = new SqlCommand("DELETE FROM transaksi WHERE id_transaksi = @id", conn, tr);
                            cmd.Parameters.AddWithValue("@id", id);
                            cmd.ExecuteNonQuery();
                        }

                        tr.Commit();
                        MessageBox.Show("Data berhasil dihapus.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    }
                    catch (Exception ex)
                    {
                        tr.Rollback();
                        MessageBox.Show("Gagal menghapus data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    }
                }

                // Refresh datagrid dan bersihkan form
                TampilkanLaporan();
                ClearForm();
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

            idEdit = -1;

            idKategoriEdit = -1;

            textBox1.Focus();

        }



        private void btnExport_Click(object sender, EventArgs e)
        {
            AppContext.SetSwitch("System.Drawing.EnableUnixSupport", true);

            if (dataGridView1.Rows.Count == 0)
            {
                MessageBox.Show("Tidak ada data yang tersedia.", "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            DialogResult pilihan = MessageBox.Show(
                "Apakah Anda ingin mengekspor data terlebih dahulu sebelum melihat laporan?\n\nYes = Ekspor & Lihat\nNo = Lihat Saja\nCancel = Batal",
                "Laporan Transaksi",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Question
            );

            if (pilihan == DialogResult.Cancel)
            {
                return;
            }
            else if (pilihan == DialogResult.No)
            {
                // Langsung buka form laporan tanpa ekspor
                FormReportExport fk = new FormReportExport();
                fk.Show();
            }
            else if (pilihan == DialogResult.Yes)
            {
                using (SaveFileDialog sfd = new SaveFileDialog() { Filter = "Excel Workbook|*.xlsx", ValidateNames = true })
                {
                    sfd.FileName = $"Laporan_Transaksi_{DateTime.Now:dd-MM-yyyy}.xlsx";

                    if (sfd.ShowDialog() == DialogResult.OK)
                    {
                        try
                        {
                            IWorkbook workbook = new XSSFWorkbook();
                            ISheet sheet = workbook.CreateSheet("Laporan Transaksi");

                            // Header
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

                            // Isi data
                            int rowIndex = 1;
                            foreach (DataGridViewRow dgvRow in dataGridView1.Rows)
                            {
                                if (dgvRow.IsNewRow) continue;

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

                            // Simpan file
                            using (FileStream fs = new FileStream(sfd.FileName, FileMode.Create, FileAccess.Write))
                            {
                                workbook.Write(fs);
                            }

                            if (File.Exists(sfd.FileName))
                            {
                                MessageBox.Show("Data berhasil diekspor ke file Excel.", "Ekspor Berhasil", MessageBoxButtons.OK, MessageBoxIcon.Information);

                                // Buka laporan setelah ekspor
                                FormReportExport fk = new FormReportExport();
                                fk.Show();
                            }
                            else
                            {
                                MessageBox.Show("File tidak berhasil dibuat.", "Ekspor Gagal", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                            }
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show($"Terjadi error saat mengekspor data.\n\nDETAIL ERROR:\n{ex.Message}", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                        }
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
                    DialogResult konfirmasi = MessageBox.Show(
                        "Apakah Anda yakin ingin mengimpor data dari file Excel ini?\n\nPastikan format sudah sesuai.",
                        "Konfirmasi Impor Data",
                        MessageBoxButtons.YesNo,
                        MessageBoxIcon.Question
                    );

                    if (konfirmasi == DialogResult.No)
                    {
                        return; // User membatalkan proses impor
                    }

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
                            int totalBarisDiproses = 0;

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

                                    totalBarisDiproses++;

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

                            // Jika semua gagal karena format
                            if (totalBarisDiproses == 0)
                            {
                                transaction.Rollback();
                                MessageBox.Show(
                                    "Impor dibatalkan karena semua data tidak sesuai format.\n\nPastikan semua kolom terisi dengan benar:\n- Nama Kategori\n- Tipe (pemasukan/pengeluaran)\n- Jumlah (angka)\n- Tanggal (format valid)",
                                    "Format Tidak Sesuai",
                                    MessageBoxButtons.OK,
                                    MessageBoxIcon.Warning
                                );
                                return;
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





        private void cari_TextChanged(object sender, EventArgs e)
        {
            if (dataGridView1.DataSource is DataTable dt)
            {
                string keyword = textBox4.Text.Trim().ToLower();
                dt.DefaultView.RowFilter = $@"
            Convert(nama_kategori, 'System.String') LIKE '%{keyword}%' OR
            Convert(tipe, 'System.String') LIKE '%{keyword}%' OR
            Convert(jumlah, 'System.String') LIKE '%{keyword}%' OR
            Convert(tanggal, 'System.String') LIKE '%{keyword}%' OR
            Convert(keterangan, 'System.String') LIKE '%{keyword}%'";
            }
        }

        private void btnGrafik_Click(object sender, EventArgs e)
        {
            FormGrafik grafik = new FormGrafik();
            grafik.Show(); // Gunakan .ShowDialog() jika ingin form modal
        }

    }

}