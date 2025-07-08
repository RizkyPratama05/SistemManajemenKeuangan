namespace UCP1
{
    partial class FormGrafik
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            System.Windows.Forms.DataVisualization.Charting.ChartArea chartArea4 = new System.Windows.Forms.DataVisualization.Charting.ChartArea();
            System.Windows.Forms.DataVisualization.Charting.Legend legend4 = new System.Windows.Forms.DataVisualization.Charting.Legend();
            System.Windows.Forms.DataVisualization.Charting.Series series4 = new System.Windows.Forms.DataVisualization.Charting.Series();
            this.chartKeuangan = new System.Windows.Forms.DataVisualization.Charting.Chart();
            this.cmbJenis = new System.Windows.Forms.ComboBox();
            this.lblCariBerdasarkan = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.chartKeuangan)).BeginInit();
            this.SuspendLayout();
            // 
            // chartKeuangan
            // 
            chartArea4.Name = "ChartArea1";
            this.chartKeuangan.ChartAreas.Add(chartArea4);
            legend4.Name = "Legend1";
            this.chartKeuangan.Legends.Add(legend4);
            this.chartKeuangan.Location = new System.Drawing.Point(37, 138);
            this.chartKeuangan.Name = "chartKeuangan";
            series4.ChartArea = "ChartArea1";
            series4.Legend = "Legend1";
            series4.Name = "Series1";
            this.chartKeuangan.Series.Add(series4);
            this.chartKeuangan.Size = new System.Drawing.Size(966, 355);
            this.chartKeuangan.TabIndex = 0;
            this.chartKeuangan.Text = "chartKeuangan";
            this.chartKeuangan.Click += new System.EventHandler(this.chartKeuangan_Click);
            // 
            // cmbJenis
            // 
            this.cmbJenis.FormattingEnabled = true;
            this.cmbJenis.Items.AddRange(new object[] {
            "Semua",
            "Pemasukan",
            "Pengeluaran"});
            this.cmbJenis.Location = new System.Drawing.Point(351, 85);
            this.cmbJenis.Name = "cmbJenis";
            this.cmbJenis.Size = new System.Drawing.Size(618, 28);
            this.cmbJenis.TabIndex = 1;
            // 
            // lblCariBerdasarkan
            // 
            this.lblCariBerdasarkan.AutoSize = true;
            this.lblCariBerdasarkan.Location = new System.Drawing.Point(87, 88);
            this.lblCariBerdasarkan.Name = "lblCariBerdasarkan";
            this.lblCariBerdasarkan.Size = new System.Drawing.Size(132, 20);
            this.lblCariBerdasarkan.TabIndex = 2;
            this.lblCariBerdasarkan.Text = "Cari Berdasarkan";
            // 
            // FormGrafik
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(9F, 20F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(192)))), ((int)(((byte)(192)))));
            this.ClientSize = new System.Drawing.Size(1502, 540);
            this.Controls.Add(this.lblCariBerdasarkan);
            this.Controls.Add(this.cmbJenis);
            this.Controls.Add(this.chartKeuangan);
            this.Name = "FormGrafik";
            this.Text = "FormGrafik";
            this.Load += new System.EventHandler(this.FormGrafik_Load);
            ((System.ComponentModel.ISupportInitialize)(this.chartKeuangan)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }


        private System.Windows.Forms.DataVisualization.Charting.Chart chartKeuangan;
        private System.Windows.Forms.ComboBox cmbJenis;
        private System.Windows.Forms.Label lblCariBerdasarkan;
    }
}