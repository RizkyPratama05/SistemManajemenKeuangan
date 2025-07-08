using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.DataVisualization.Charting;

namespace UCP1
{
    public partial class FormGrafik : Form
    {

        Koneksi kn = new Koneksi();
        string connect = "";
        public FormGrafik()
        {
            InitializeComponent();
            this.Load += new System.EventHandler(this.FormGrafik_Load);
        }

        private void chartKeuangan_Click(object sender, EventArgs e)
        {
            // Isi ComboBox
            //cmbJenis.Items.AddRange(new string[] { "Semua", "Pemasukan", "Pengeluaran" });
            //cmbJenis.SelectedIndex = 0;

            LoadChartData("Semua");

            // Tambahkan event handler jika belum
            cmbJenis.SelectedIndexChanged += cmbJenis_SelectedIndexChanged;
        }

        private void cmbJenis_SelectedIndexChanged(object sender, EventArgs e)
        {
            string selected = cmbJenis.SelectedItem.ToString();
            LoadChartData(selected);
        }

        private void LoadChartData(string filter)
        {
            // Bersihkan isi chart
            chartKeuangan.Series.Clear();
            chartKeuangan.Titles.Clear();
            chartKeuangan.Legends.Clear();
            chartKeuangan.ChartAreas.Clear();

            // Buat area chart
            ChartArea ca = new ChartArea("MainArea");
            ca.AxisX.Title = "Tipe Transaksi";
            ca.AxisY.Title = "Jumlah (Rp)";
            ca.AxisX.LabelStyle.Angle = -45;
            ca.BackColor = System.Drawing.Color.LightGoldenrodYellow;
            chartKeuangan.ChartAreas.Add(ca);

            // Tambahkan legend
            Legend legend = new Legend("Legenda");
            chartKeuangan.Legends.Add(legend);

            // Query SQL untuk menjumlahkan transaksi per tipe
            connect = kn.connectionString();
            string query = @"
        SELECT
            k.tipe,
            SUM(t.jumlah) AS total_jumlah
        FROM transaksi t
        JOIN kategori k ON t.id_kategori = k.id_kategori
        GROUP BY k.tipe";

            // Dictionary untuk menyimpan total per tipe
            Dictionary<string, decimal> totals = new Dictionary<string, decimal>()
    {
        { "pemasukan", 0 },
        { "pengeluaran", 0 }
    };

            using (SqlConnection conn = new SqlConnection(kn.connectionString()))
            {
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();
                SqlDataReader reader = cmd.ExecuteReader();
                while (reader.Read())
                {
                    string tipe = reader["tipe"].ToString().Trim().ToLower(); // normalisasi huruf kecil
                    decimal total = Convert.ToDecimal(reader["total_jumlah"]);

                    if (totals.ContainsKey(tipe))
                        totals[tipe] = total;
                }
            }

            // Buat series grafik batang
            Series series = new Series("Transaksi")
            {
                ChartType = SeriesChartType.Column,
                IsValueShownAsLabel = true
            };

            // Tambahkan data ke grafik sesuai filter
            if (filter.ToLower() == "semua" || filter.ToLower() == "pemasukan")
                series.Points.AddXY("Pemasukan", totals["pemasukan"]);

            if (filter.ToLower() == "semua" || filter.ToLower() == "pengeluaran")
                series.Points.AddXY("Pengeluaran", totals["pengeluaran"]);

            chartKeuangan.Series.Add(series);
            chartKeuangan.Titles.Add("Grafik Keuangan Organisasi");
        }



        private void FormGrafik_Load(object sender, EventArgs e)
        {
            //cmbJenis.Items.AddRange(new string[] { "Semua", "Pemasukan", "Pengeluaran" });
            //cmbJenis.SelectedIndex = 0;

            // Tambahkan event handler
            cmbJenis.SelectedIndexChanged += cmbJenis_SelectedIndexChanged;

            // Load data pertama kali
            LoadChartData("Semua");
        }


    }
}
