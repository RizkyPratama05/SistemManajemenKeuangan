using System;
using System.Data;
using System.Data.SqlClient;
using System.Windows.Forms;
using System.Drawing;

namespace TransaksiApp
{
    public partial class FormTambahTransaksi : Form
    {
        private int idEdit = -1;
        private string connectionString = "Data Source=MSI\\RIZKYPP;Initial Catalog=A4;Integrated Security=True";

        public FormTambahTransaksi()
        {
            InitializeComponent();
            dataGridView1.CellClick += dataGridView1_CellClick;
        }

        private void FormTambahTransaksi_Load(object sender, EventArgs e)
        {
            comboBox1.Items.Add("pemasukan");
            comboBox1.Items.Add("pengeluaran");
            comboBox1.SelectedIndex = 0;
            TampilkanLaporan();
        }

        private void btnSimpan_Click(object sender, EventArgs e)
        {
            string nama = textBox1.Text;
            decimal jumlah = Convert.ToDecimal(textBox2.Text);
            string keterangan = textBox3.Text;
            DateTime tanggal = dateTimePicker1.Value;
            string tipe = comboBox1.SelectedItem.ToString();

            if (string.IsNullOrWhiteSpace(nama) || jumlah <= 0)
            {
                MessageBox.Show("Nama dan jumlah harus diisi dengan benar.");
                return;
            }

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
                    
                    textBox1.Text = dataGridView1.Rows[e.RowIndex].Cells["nama_kategori"].Value.ToString();
                    comboBox1.SelectedItem = dataGridView1.Rows[e.RowIndex].Cells["tipe"].Value.ToString();
                    textBox2.Text = dataGridView1.Rows[e.RowIndex].Cells["jumlah"].Value.ToString();
                    textBox3.Text = dataGridView1.Rows[e.RowIndex].Cells["keterangan"].Value.ToString();
                    dateTimePicker1.Value = Convert.ToDateTime(dataGridView1.Rows[e.RowIndex].Cells["tanggal"].Value);
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

                    string query = "SELECT id_transaksi, nama_kategori, tipe, jumlah, tanggal, keterangan FROM view_transaksi_lengkap ORDER BY tanggal DESC";
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
                    totalRow["tipe"] = "";
                    totalRow["jumlah"] = DBNull.Value;
                    totalRow["keterangan"] = $"Pemasukan: Rp{totalPemasukan:N0} | Pengeluaran: Rp{totalPengeluaran:N0} | Saldo: Rp{(totalPemasukan - totalPengeluaran):N0}";
                    dt.Rows.Add(totalRow);

                    
                    dataGridView1.DataSource = null;
                    dataGridView1.Columns.Clear();
                    dataGridView1.AutoGenerateColumns = false;

                    
                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "id_transaksi",
                        HeaderText = "ID",
                        Name = "id_transaksi",
                        Visible = false 
                    });

                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "nama_kategori",
                        HeaderText = "Kategori",
                        Name = "nama_kategori"
                    });

                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "tipe",
                        HeaderText = "Jenis",
                        Name = "tipe"
                    });

                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "jumlah",
                        HeaderText = "Jumlah",
                        Name = "jumlah",
                        DefaultCellStyle = new DataGridViewCellStyle() { Format = "N0" }
                    });

                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "tanggal",
                        HeaderText = "Tanggal",
                        Name = "tanggal",
                        DefaultCellStyle = new DataGridViewCellStyle() { Format = "dd/MM/yyyy" }
                    });

                    dataGridView1.Columns.Add(new DataGridViewTextBoxColumn()
                    {
                        DataPropertyName = "keterangan",
                        HeaderText = "Keterangan",
                        Name = "keterangan"
                    });

                   
                    DataGridViewButtonColumn btnUpdate = new DataGridViewButtonColumn()
                    {
                        HeaderText = "Aksi",
                        Name = "Update",
                        Text = "Update",
                        UseColumnTextForButtonValue = true,
                        FlatStyle = FlatStyle.Flat,
                        DefaultCellStyle = new DataGridViewCellStyle()
                        {
                            BackColor = Color.LightBlue,
                            ForeColor = Color.Black
                        }
                    };

                    DataGridViewButtonColumn btnDelete = new DataGridViewButtonColumn()
                    {
                        Name = "Delete",
                        Text = "Delete",
                        UseColumnTextForButtonValue = true,
                        FlatStyle = FlatStyle.Flat,
                        DefaultCellStyle = new DataGridViewCellStyle()
                        {
                            BackColor = Color.LightCoral,
                            ForeColor = Color.Black
                        }
                    };

                    dataGridView1.Columns.Add(btnUpdate);
                    dataGridView1.Columns.Add(btnDelete);

                    
                    dataGridView1.DataSource = dt;

                    
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font("Arial", 9, FontStyle.Bold);
                    dataGridView1.EnableHeadersVisualStyles = false;
                    dataGridView1.ColumnHeadersDefaultCellStyle.BackColor = Color.LightGray;
                    dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
                    dataGridView1.RowHeadersVisible = false;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error saat menampilkan laporan: {ex.Message}");
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