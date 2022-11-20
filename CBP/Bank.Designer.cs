namespace CBP
{
    partial class Bank
    {
        /// <summary>
        /// Variabile di progettazione necessaria.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Pulire le risorse in uso.
        /// </summary>
        /// <param name="disposing">ha valore true se le risorse gestite devono essere eliminate, false in caso contrario.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Codice generato da Progettazione Windows Form

        /// <summary>
        /// Metodo necessario per il supporto della finestra di progettazione. Non modificare
        /// il contenuto del metodo con l'editor di codice.
        /// </summary>
        private void InitializeComponent()
        {
            this.OpenOutput = new System.Windows.Forms.Button();
            this.OutputFolder = new System.Windows.Forms.TextBox();
            this.label11 = new System.Windows.Forms.Label();
            this.ExportMonthly = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.DataBonifico = new System.Windows.Forms.DateTimePicker();
            this.CreaExcelDaReportBase = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // OpenOutput
            // 
            this.OpenOutput.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.OpenOutput.FlatAppearance.BorderSize = 2;
            this.OpenOutput.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.OpenOutput.Font = new System.Drawing.Font("Calibri", 10.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.OpenOutput.ForeColor = System.Drawing.Color.White;
            this.OpenOutput.Location = new System.Drawing.Point(430, 62);
            this.OpenOutput.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.OpenOutput.Name = "OpenOutput";
            this.OpenOutput.Size = new System.Drawing.Size(45, 30);
            this.OpenOutput.TabIndex = 16;
            this.OpenOutput.Text = "...";
            this.OpenOutput.UseVisualStyleBackColor = true;
            this.OpenOutput.Click += new System.EventHandler(this.OpenOutput_Click);
            // 
            // OutputFolder
            // 
            this.OutputFolder.Location = new System.Drawing.Point(39, 61);
            this.OutputFolder.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.OutputFolder.Name = "OutputFolder";
            this.OutputFolder.Size = new System.Drawing.Size(372, 32);
            this.OutputFolder.TabIndex = 18;
            this.OutputFolder.Text = "C:\\Develop\\CBP\\Dati\\Output";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(34, 30);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(161, 26);
            this.label11.TabIndex = 17;
            this.label11.Text = "Cartella di output";
            // 
            // ExportMonthly
            // 
            this.ExportMonthly.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.ExportMonthly.FlatAppearance.BorderSize = 2;
            this.ExportMonthly.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.ExportMonthly.ForeColor = System.Drawing.Color.White;
            this.ExportMonthly.Location = new System.Drawing.Point(43, 393);
            this.ExportMonthly.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.ExportMonthly.Name = "ExportMonthly";
            this.ExportMonthly.Size = new System.Drawing.Size(432, 69);
            this.ExportMonthly.TabIndex = 0;
            this.ExportMonthly.Text = "AVVIA CREAZIONE XML BANCA";
            this.ExportMonthly.UseVisualStyleBackColor = true;
            this.ExportMonthly.Click += new System.EventHandler(this.ExportMonthly_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(34, 127);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(126, 26);
            this.label1.TabIndex = 19;
            this.label1.Text = "Data bonifico";
            // 
            // DataBonifico
            // 
            this.DataBonifico.Location = new System.Drawing.Point(39, 156);
            this.DataBonifico.Name = "DataBonifico";
            this.DataBonifico.Size = new System.Drawing.Size(372, 32);
            this.DataBonifico.TabIndex = 20;
            // 
            // CreaExcelDaReportBase
            // 
            this.CreaExcelDaReportBase.FlatAppearance.BorderColor = System.Drawing.Color.White;
            this.CreaExcelDaReportBase.FlatAppearance.BorderSize = 2;
            this.CreaExcelDaReportBase.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.CreaExcelDaReportBase.ForeColor = System.Drawing.Color.White;
            this.CreaExcelDaReportBase.Location = new System.Drawing.Point(43, 298);
            this.CreaExcelDaReportBase.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.CreaExcelDaReportBase.Name = "CreaExcelDaReportBase";
            this.CreaExcelDaReportBase.Size = new System.Drawing.Size(432, 69);
            this.CreaExcelDaReportBase.TabIndex = 21;
            this.CreaExcelDaReportBase.Text = "CREAZIONE EXCEL DA REPORT BASE PAGAMENTI";
            this.CreaExcelDaReportBase.UseVisualStyleBackColor = true;
            this.CreaExcelDaReportBase.Click += new System.EventHandler(this.CreaExcelDaReportBase_Click);
            // 
            // Bank
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(68)))), ((int)(((byte)(139)))));
            this.ClientSize = new System.Drawing.Size(513, 513);
            this.Controls.Add(this.CreaExcelDaReportBase);
            this.Controls.Add(this.DataBonifico);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.OpenOutput);
            this.Controls.Add(this.ExportMonthly);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.OutputFolder);
            this.Font = new System.Drawing.Font("Calibri Light", 12.2F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ForeColor = System.Drawing.Color.White;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(3, 5, 3, 5);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "Bank";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "CBP - Bank Management";
            this.Load += new System.EventHandler(this.Bank_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox OutputFolder;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button ExportMonthly;
        private System.Windows.Forms.Button OpenOutput;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DateTimePicker DataBonifico;
        private System.Windows.Forms.Button CreaExcelDaReportBase;
    }
}

