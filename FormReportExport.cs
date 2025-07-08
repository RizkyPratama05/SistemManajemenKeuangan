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
using Microsoft.Reporting.WinForms;
using System.IO;// Pastikan ini ada dan digunakan

namespace UCP1
{
    public partial class FormReportExport : Form
    {
        Koneksi kn = new Koneksi();
        string connect = "";
        public FormReportExport()
        {
            InitializeComponent();
        }

        private void SetupReportViewer()
        {
            // Connection string to your database
            connect = kn.connectionString();

            // SQL query to retrieve the required data from the database
            string query = @"
                SELECT
                    t.id_transaksi,
                    k.id_kategori,
                    k.nama_kategori,
                    k.tipe,
                    t.jumlah,
                    t.tanggal,
                    t.keterangan
                FROM transaksi t
                    JOIN kategori k ON t.id_kategori = k.id_kategori";

            // Create a DataTable to store the data
            DataTable dt = new DataTable();

            // Use SqlDataAdapter to fill the DataTable with data from the database
            using (SqlConnection conn = new SqlConnection(kn.connectionString()))
            {
                SqlDataAdapter da = new SqlDataAdapter(query, conn);
                da.Fill(dt);
            }

            // Create a ReportDataSource
            // INI BAGIAN YANG DIPERBAIKI: Harus ReportDataSource, bukan FormReportExport
            ReportDataSource rds = new ReportDataSource("DataSet1", dt);
            reportViewer1.LocalReport.DataSources.Clear();
            reportViewer1.LocalReport.DataSources.Add(rds);

            string reportPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "KeuanganReport.rdlc");
            reportViewer1.LocalReport.ReportPath = reportPath;
            reportViewer1.RefreshReport();
        }

        private void FormReportExport_Load(object sender, EventArgs e)
        {
            SetupReportViewer();
        }

    }
}