using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace UCP1
{
    public partial class PreviewData : Form
    {
        private string connectionString = "Data Source=MSI\\RIZKYPP;Initial Catalog=SisTemManajemenKeuangan;Integrated Security=True";
        // Konstruktor menerima DataTable dan menampilkannya di DataGridView
        public PreviewData(DataTable data)
        {
            InitializeComponent();
            // Menyimpan data source DataGridView ke DataTable yang diterima
            dataGridView1.DataSource = data;
        }

        // Event ketika form dimuat
        private void PreviewForm_Load(object sender, EventArgs e)
        {
            // Opsional: Sesuaikan DataGridView jika perlu
            dataGridView1.AutoResizeColumns(); // Menyesuaikan ukuran kolom
        }

        // Event ketika tombol OK ditekan
        private void button1_Click(object sender, EventArgs e)
        {
            // Menanyakan kepada pengguna jika mereka ingin mengimpor data
            DialogResult result = MessageBox.Show("Apakah Anda ingin mengimpor data ini ke database?", "Konfirmasi", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            if (result == DialogResult.Yes)
            {
                // Mengimpor data dari DataGridView ke database
                ImportDataToDatabase();
            }
        }

        private void ImportDataToDatabase()
        {
            try
            {
                DataTable dt = (DataTable)dataGridView1.DataSource;


                foreach (DataRow row in dt.Rows)
                {
                    // Ganti query ini dengan INSERT INTO
                    string query = @"INSERT INTO transaksi (id_kategori, jumlah, tanggal, keterangan)
                 VALUES (@id_kategori, @jumlah, @tanggal, @keterangan)";


                    using (SqlConnection conn = new SqlConnection(connectionString))
                    {
                        conn.Open();
                        using (SqlCommand cmd = new SqlCommand(query, conn))
                        {
                            cmd.Parameters.AddWithValue("@id_kategori", row["id_kategori"]);
                            cmd.Parameters.AddWithValue("@Jumlah", row["Jumlah"]);
                            string tanggalStr = row["Tanggal"].ToString().Trim();
                            string[] formats = { "dd/MM/yyyy", "d/M/yyyy", "yyyy-MM-dd", "dd-MMM-yyyy" };

                            if (!DateTime.TryParseExact(tanggalStr, formats, CultureInfo.InvariantCulture, DateTimeStyles.None, out DateTime parsedDate))
                            {
                                throw new FormatException($"Format tanggal tidak dikenali: '{tanggalStr}'");
                            }

                            // hanya tanggal (jam diabaikan karena SQL tipe DATE)
                            cmd.Parameters.AddWithValue("@tanggal", parsedDate.Date);   
                            cmd.Parameters.AddWithValue("@Keterangan", row["Keterangan"]);
                            cmd.ExecuteNonQuery();
                        }
                    }

                }

                MessageBox.Show("Data berhasil diimpor ke database.", "Sukses", MessageBoxButtons.OK, MessageBoxIcon.Information);
                this.Close(); // Tutup PreviewForm setelah data diimpor
            }
            catch (Exception ex)
            {
                MessageBox.Show("Terjadi kesalahan saat mengimpor data: " + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }


    }

}
