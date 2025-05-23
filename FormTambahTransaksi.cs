using System;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Windows.Forms;

namespace UCP1
{
    public partial class FormTambahTransaksi : Form
    {
        private int idEdit = -1;
        private string connectionString = "Data Source=MSI\\RIZKYPP;Initial Catalog=A4;Integrated Security=True";

        public FormTambahTransaksi()
        {
            InitializeComponent();
            dataGridView1.CellClick += dataGridView1_CellClick;


            button1.Text = "Lihat Laporan";
            button1.BackColor = Color.LightGreen;
            button1.Font = new Font("Arial", 9, FontStyle.Bold);
            button1.Click += BtnLihatLaporan_Click;
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

        private void FormTambahTransaksi_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("pemasukan");
            comboBox1.Items.Add("pengeluaran");
            comboBox1.SelectedIndex = 0;
            dateTimePicker1.MinDate = new DateTime(2020, 1, 1);

            TampilkanLaporan();
        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            string nama = textBox1.Text;
            string jumlahText = textBox2.Text.Trim(); 
            string keterangan = textBox3.Text;
            DateTime tanggal = dateTimePicker1.Value;
            string tipe = comboBox1.SelectedItem.ToString();

            // Validasi nama: hanya huruf dan spasi
            if (string.IsNullOrWhiteSpace(nama) || !System.Text.RegularExpressions.Regex.IsMatch(nama, @"^[a-zA-Z\s]+$"))
            {
                MessageBox.Show("Nama kategori harus terdiri dari huruf saja.", "Validasi Error", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            // Validasi jumlah
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
                                  "Konfirmasi Simpan",
                                  MessageBoxButtons.YesNo,
                                  MessageBoxIcon.Question);
            if (confirmSave != DialogResult.Yes)
                return;


            using (SqlConnection conn = new SqlConnection(connectionString))
            {
                conn.Open();

                if (idEdit == -1)
                {
                    string insertKategori = "INSERT INTO kategori (nama_kategori, tipe) VALUES (@nama, @tipe); SELECT SCOPE_IDENTITY();";
                    SqlCommand cmdKategori = new SqlCommand(insertKategori, conn);
                    cmdKategori.Parameters.AddWithValue("@nama", nama);
                    cmdKategori.Parameters.AddWithValue("@tipe", tipe);
                    int idKategori = Convert.ToInt32(cmdKategori.ExecuteScalar());

                    string insertTransaksi = "INSERT INTO transaksi (id_kategori, jumlah, tanggal, keterangan) VALUES (@id_kategori, @jumlah, @tanggal, @keterangan)";
                    SqlCommand cmdTransaksi = new SqlCommand(insertTransaksi, conn);
                    cmdTransaksi.Parameters.AddWithValue("@id_kategori", idKategori);
                    cmdTransaksi.Parameters.AddWithValue("@jumlah", jumlah);
                    cmdTransaksi.Parameters.AddWithValue("@tanggal", tanggal);
                    cmdTransaksi.Parameters.AddWithValue("@keterangan", keterangan);
                    cmdTransaksi.ExecuteNonQuery();

                    MessageBox.Show("Transaksi berhasil ditambahkan.");
                }
                else
                {

                 

                    string updateTransaksi = "UPDATE transaksi SET jumlah=@jumlah, tanggal=@tanggal, keterangan=@keterangan WHERE id_transaksi=@id";
                    SqlCommand cmdUpdate = new SqlCommand(updateTransaksi, conn);
                    cmdUpdate.Parameters.AddWithValue("@jumlah", jumlah);
                    cmdUpdate.Parameters.AddWithValue("@tanggal", tanggal);
                    cmdUpdate.Parameters.AddWithValue("@keterangan", keterangan);
                    cmdUpdate.Parameters.AddWithValue("@id", idEdit);
                    cmdUpdate.ExecuteNonQuery();

                    MessageBox.Show("Transaksi berhasil diupdate.");
                    idEdit = -1;
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

                string columnName = dataGridView1.Columns[e.ColumnIndex].Name;
                if (columnName != "Update" && columnName != "Delete") return;

                var idValue = dataGridView1.Rows[e.RowIndex].Cells["id_transaksi"].Value;
                if (idValue == DBNull.Value || idValue == null || idValue.ToString() == "") return;

                int idTransaksi = Convert.ToInt32(idValue);

                if (columnName == "Update")
                {

                    var confirmUpdate = MessageBox.Show("Apakah Anda yakin ingin mengupdate data ini?",
                                       "Konfirmasi Update",
                                       MessageBoxButtons.YesNo,
                                       MessageBoxIcon.Question);
                    if (confirmUpdate != DialogResult.Yes)
                    {
                        return;
                    }

                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells["nama_kategori"].Value.ToString();
                    comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells["tipe"].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells["jumlah"].Value.ToString();
                    textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells["keterangan"].Value.ToString();
                    idEdit = idTransaksi;
                }
                else if (columnName == "Delete")
                {
                    var confirm = MessageBox.Show("Yakin ingin menghapus transaksi ini?", "Konfirmasi", MessageBoxButtons.YesNo);
                    if (confirm == DialogResult.Yes)
                    {
                        using (SqlConnection conn = new SqlConnection(connectionString))
                        {
                            conn.Open();
                            SqlCommand cmd = new SqlCommand("DELETE FROM transaksi WHERE id_transaksi = @id", conn);
                            cmd.Parameters.AddWithValue("@id", idTransaksi);
                            cmd.ExecuteNonQuery();
                            MessageBox.Show("Transaksi berhasil dihapus.");
                            TampilkanLaporan();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}\n\nDetail: {ex.StackTrace}");
            }
        }

        private void TampilkanLaporan()
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(connectionString))
                {
                    conn.Open();
                    string query = @"SELECT id_transaksi, nama_kategori, tipe, jumlah, 
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

            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            dataGridView1.AutoGenerateColumns = false;


            AddColumn("id_transaksi", "ID", false);
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

        private void AddColumn(string name, string headerText, bool visible = true,
                             string format = null, DataGridViewContentAlignment alignment = DataGridViewContentAlignment.MiddleLeft)
        {
            var col = new DataGridViewTextBoxColumn
            {
                Name = name,
                DataPropertyName = name,
                HeaderText = headerText,
                Visible = visible,
                DefaultCellStyle = new DataGridViewCellStyle { Alignment = alignment }
            };

            if (!string.IsNullOrEmpty(format))
                col.DefaultCellStyle.Format = format;

            dataGridView1.Columns.Add(col);
        }

        private void AddButtonColumn(string name, string text, Color backColor)
        {
            var btnCol = new DataGridViewButtonColumn
            {
                Name = name,
                Text = text,
                HeaderText = "Aksi",
                UseColumnTextForButtonValue = true,
                FlatStyle = FlatStyle.Flat,
                DefaultCellStyle = new DataGridViewCellStyle
                {
                    BackColor = backColor,
                    ForeColor = Color.Black,
                    Alignment = DataGridViewContentAlignment.MiddleCenter
                }
            };
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
                var lastRow = dataGridView1.Rows[dataGridView1.Rows.Count - 1];
                lastRow.DefaultCellStyle.Font = new Font(dataGridView1.Font, FontStyle.Bold);
                lastRow.DefaultCellStyle.BackColor = Color.LightGray;
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
        }
    }
}