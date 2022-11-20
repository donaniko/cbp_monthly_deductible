namespace CBP
{
    partial class RecessoDisdetta
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

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.label17 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.NumeroPolizza = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.StatoPolizza = new System.Windows.Forms.TextBox();
            this.DataCancellazione = new System.Windows.Forms.DateTimePicker();
            this.label1 = new System.Windows.Forms.Label();
            this.RichiestaSottoscrittore = new System.Windows.Forms.CheckBox();
            this.CartaIdentita = new System.Windows.Forms.CheckBox();
            this.CodiceFiscale = new System.Windows.Forms.CheckBox();
            this.label13 = new System.Windows.Forms.Label();
            this.CointestatarioCartaIdentita = new System.Windows.Forms.CheckBox();
            this.label12 = new System.Windows.Forms.Label();
            this.Cointestatario = new System.Windows.Forms.CheckBox();
            this.label5 = new System.Windows.Forms.Label();
            this.DataLavorazione = new System.Windows.Forms.DateTimePicker();
            this.label10 = new System.Windows.Forms.Label();
            this.DataRimborso = new System.Windows.Forms.DateTimePicker();
            this.label9 = new System.Windows.Forms.Label();
            this.DebitoResiduo = new System.Windows.Forms.TextBox();
            this.label8 = new System.Windows.Forms.Label();
            this.NuovoStatoPolizza = new System.Windows.Forms.TextBox();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.Iban = new System.Windows.Forms.TextBox();
            this.CloseAndSave = new System.Windows.Forms.Button();
            this.Confirm = new System.Windows.Forms.Button();
            this.Cancel = new System.Windows.Forms.Button();
            this.DataDecorrenza = new System.Windows.Forms.TextBox();
            this.ListaBeneficiari = new System.Windows.Forms.DataGridView();
            this.tabControl1 = new System.Windows.Forms.TabControl();
            this.tabPage1 = new System.Windows.Forms.TabPage();
            this.tabPage2 = new System.Windows.Forms.TabPage();
            this.QuotaSum = new System.Windows.Forms.Label();
            this.AddBeneficiary = new System.Windows.Forms.Button();
            this.CointestatrioCognome = new System.Windows.Forms.TextBox();
            this.label14 = new System.Windows.Forms.Label();
            this.label15 = new System.Windows.Forms.Label();
            this.CointestatarioNome = new System.Windows.Forms.TextBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label16 = new System.Windows.Forms.Label();
            this.CointestatarioDataNascita = new System.Windows.Forms.DateTimePicker();
            ((System.ComponentModel.ISupportInitialize)(this.ListaBeneficiari)).BeginInit();
            this.tabControl1.SuspendLayout();
            this.tabPage1.SuspendLayout();
            this.tabPage2.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label17.Location = new System.Drawing.Point(453, 22);
            this.label17.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(172, 26);
            this.label17.TabIndex = 40;
            this.label17.Text = "Data di decorrenza";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label2.Location = new System.Drawing.Point(30, 22);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 26);
            this.label2.TabIndex = 36;
            this.label2.Text = "Numero polizza";
            // 
            // NumeroPolizza
            // 
            this.NumeroPolizza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NumeroPolizza.Location = new System.Drawing.Point(33, 52);
            this.NumeroPolizza.Margin = new System.Windows.Forms.Padding(4);
            this.NumeroPolizza.MaxLength = 10;
            this.NumeroPolizza.Name = "NumeroPolizza";
            this.NumeroPolizza.ReadOnly = true;
            this.NumeroPolizza.Size = new System.Drawing.Size(257, 32);
            this.NumeroPolizza.TabIndex = 1;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label3.Location = new System.Drawing.Point(298, 22);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 26);
            this.label3.TabIndex = 39;
            this.label3.Text = "Stato polizza";
            // 
            // StatoPolizza
            // 
            this.StatoPolizza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.StatoPolizza.Location = new System.Drawing.Point(302, 52);
            this.StatoPolizza.Margin = new System.Windows.Forms.Padding(4);
            this.StatoPolizza.Name = "StatoPolizza";
            this.StatoPolizza.ReadOnly = true;
            this.StatoPolizza.Size = new System.Drawing.Size(144, 32);
            this.StatoPolizza.TabIndex = 2;
            // 
            // DataCancellazione
            // 
            this.DataCancellazione.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.DataCancellazione.Location = new System.Drawing.Point(236, 82);
            this.DataCancellazione.Name = "DataCancellazione";
            this.DataCancellazione.Size = new System.Drawing.Size(267, 32);
            this.DataCancellazione.TabIndex = 5;
            this.DataCancellazione.Value = new System.DateTime(2020, 11, 5, 0, 0, 0, 0);
            this.DataCancellazione.ValueChanged += new System.EventHandler(this.DataCancellazione_ValueChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label1.Location = new System.Drawing.Point(24, 92);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(175, 26);
            this.label1.TabIndex = 42;
            this.label1.Text = "Data della richiesta";
            // 
            // RichiestaSottoscrittore
            // 
            this.RichiestaSottoscrittore.BackColor = System.Drawing.Color.Transparent;
            this.RichiestaSottoscrittore.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.RichiestaSottoscrittore.Location = new System.Drawing.Point(284, 196);
            this.RichiestaSottoscrittore.Name = "RichiestaSottoscrittore";
            this.RichiestaSottoscrittore.Size = new System.Drawing.Size(219, 32);
            this.RichiestaSottoscrittore.TabIndex = 7;
            this.RichiestaSottoscrittore.UseVisualStyleBackColor = false;
            // 
            // CartaIdentita
            // 
            this.CartaIdentita.BackColor = System.Drawing.Color.Transparent;
            this.CartaIdentita.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CartaIdentita.Location = new System.Drawing.Point(284, 253);
            this.CartaIdentita.Name = "CartaIdentita";
            this.CartaIdentita.Size = new System.Drawing.Size(219, 32);
            this.CartaIdentita.TabIndex = 8;
            this.CartaIdentita.UseVisualStyleBackColor = false;
            // 
            // CodiceFiscale
            // 
            this.CodiceFiscale.CheckAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.CodiceFiscale.Location = new System.Drawing.Point(284, 306);
            this.CodiceFiscale.Name = "CodiceFiscale";
            this.CodiceFiscale.Size = new System.Drawing.Size(219, 32);
            this.CodiceFiscale.TabIndex = 9;
            this.CodiceFiscale.UseVisualStyleBackColor = true;
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label13.Location = new System.Drawing.Point(11, 92);
            this.label13.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(142, 26);
            this.label13.TabIndex = 63;
            this.label13.Text = "Carta d\'identità";
            // 
            // CointestatarioCartaIdentita
            // 
            this.CointestatarioCartaIdentita.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.CointestatarioCartaIdentita.Location = new System.Drawing.Point(310, 89);
            this.CointestatarioCartaIdentita.Name = "CointestatarioCartaIdentita";
            this.CointestatarioCartaIdentita.Size = new System.Drawing.Size(16, 32);
            this.CointestatarioCartaIdentita.TabIndex = 62;
            this.CointestatarioCartaIdentita.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.CointestatarioCartaIdentita.UseVisualStyleBackColor = true;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label12.Location = new System.Drawing.Point(9, 46);
            this.label12.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(132, 26);
            this.label12.TabIndex = 61;
            this.label12.Text = "Cointestatario";
            // 
            // Cointestatario
            // 
            this.Cointestatario.CheckAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Cointestatario.Location = new System.Drawing.Point(310, 46);
            this.Cointestatario.Name = "Cointestatario";
            this.Cointestatario.Size = new System.Drawing.Size(16, 32);
            this.Cointestatario.TabIndex = 60;
            this.Cointestatario.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.Cointestatario.UseVisualStyleBackColor = true;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label5.Location = new System.Drawing.Point(24, 37);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(158, 26);
            this.label5.TabIndex = 59;
            this.label5.Text = "Data Lavorazione";
            // 
            // DataLavorazione
            // 
            this.DataLavorazione.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.DataLavorazione.Location = new System.Drawing.Point(236, 31);
            this.DataLavorazione.Name = "DataLavorazione";
            this.DataLavorazione.Size = new System.Drawing.Size(267, 32);
            this.DataLavorazione.TabIndex = 4;
            this.DataLavorazione.Value = new System.DateTime(2020, 11, 5, 0, 0, 0, 0);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label10.Location = new System.Drawing.Point(24, 147);
            this.label10.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(165, 26);
            this.label10.TabIndex = 57;
            this.label10.Text = "Data del rimborso";
            // 
            // DataRimborso
            // 
            this.DataRimborso.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.DataRimborso.Location = new System.Drawing.Point(236, 139);
            this.DataRimborso.Name = "DataRimborso";
            this.DataRimborso.Size = new System.Drawing.Size(267, 32);
            this.DataRimborso.TabIndex = 6;
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label9.Location = new System.Drawing.Point(24, 479);
            this.label9.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(137, 26);
            this.label9.TabIndex = 55;
            this.label9.Text = "Debito residuo";
            // 
            // DebitoResiduo
            // 
            this.DebitoResiduo.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.DebitoResiduo.Location = new System.Drawing.Point(236, 477);
            this.DebitoResiduo.Margin = new System.Windows.Forms.Padding(4);
            this.DebitoResiduo.MaxLength = 10;
            this.DebitoResiduo.Name = "DebitoResiduo";
            this.DebitoResiduo.Size = new System.Drawing.Size(267, 32);
            this.DebitoResiduo.TabIndex = 12;
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label8.Location = new System.Drawing.Point(24, 424);
            this.label8.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(178, 26);
            this.label8.TabIndex = 53;
            this.label8.Text = "Nuovo stato polizza";
            // 
            // NuovoStatoPolizza
            // 
            this.NuovoStatoPolizza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NuovoStatoPolizza.Location = new System.Drawing.Point(236, 420);
            this.NuovoStatoPolizza.Margin = new System.Windows.Forms.Padding(4);
            this.NuovoStatoPolizza.MaxLength = 10;
            this.NuovoStatoPolizza.Name = "NuovoStatoPolizza";
            this.NuovoStatoPolizza.Size = new System.Drawing.Size(267, 32);
            this.NuovoStatoPolizza.TabIndex = 11;
            this.NuovoStatoPolizza.Leave += new System.EventHandler(this.NuovoStatoPolizza_Leave);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label7.Location = new System.Drawing.Point(24, 312);
            this.label7.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(127, 26);
            this.label7.TabIndex = 51;
            this.label7.Text = "Codice fiscale";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label6.Location = new System.Drawing.Point(24, 257);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(142, 26);
            this.label6.TabIndex = 50;
            this.label6.Text = "Carta d\'identità";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label11.Location = new System.Drawing.Point(24, 202);
            this.label11.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(206, 26);
            this.label11.TabIndex = 49;
            this.label11.Text = "Richiesta sottoscrittore";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label4.Location = new System.Drawing.Point(24, 369);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(53, 26);
            this.label4.TabIndex = 48;
            this.label4.Text = "IBAN";
            // 
            // Iban
            // 
            this.Iban.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Iban.Location = new System.Drawing.Point(236, 363);
            this.Iban.Margin = new System.Windows.Forms.Padding(4);
            this.Iban.MaxLength = 27;
            this.Iban.Name = "Iban";
            this.Iban.Size = new System.Drawing.Size(267, 32);
            this.Iban.TabIndex = 10;
            // 
            // CloseAndSave
            // 
            this.CloseAndSave.Location = new System.Drawing.Point(208, 692);
            this.CloseAndSave.Name = "CloseAndSave";
            this.CloseAndSave.Size = new System.Drawing.Size(201, 39);
            this.CloseAndSave.TabIndex = 14;
            this.CloseAndSave.Text = "CHIUDI E SALVA";
            this.CloseAndSave.UseVisualStyleBackColor = true;
            this.CloseAndSave.Click += new System.EventHandler(this.CloseAndSave_Click);
            // 
            // Confirm
            // 
            this.Confirm.BackColor = System.Drawing.Color.SteelBlue;
            this.Confirm.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.Confirm.ForeColor = System.Drawing.Color.White;
            this.Confirm.Location = new System.Drawing.Point(35, 692);
            this.Confirm.Name = "Confirm";
            this.Confirm.Size = new System.Drawing.Size(167, 39);
            this.Confirm.TabIndex = 13;
            this.Confirm.Text = "PROCEDI";
            this.Confirm.UseVisualStyleBackColor = false;
            this.Confirm.Click += new System.EventHandler(this.Confirm_Click);
            // 
            // Cancel
            // 
            this.Cancel.Location = new System.Drawing.Point(415, 692);
            this.Cancel.Name = "Cancel";
            this.Cancel.Size = new System.Drawing.Size(220, 39);
            this.Cancel.TabIndex = 15;
            this.Cancel.Text = "CHIUDI SENZA SALVARE";
            this.Cancel.UseVisualStyleBackColor = true;
            this.Cancel.Click += new System.EventHandler(this.Cancel_Click);
            // 
            // DataDecorrenza
            // 
            this.DataDecorrenza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.DataDecorrenza.Location = new System.Drawing.Point(458, 52);
            this.DataDecorrenza.Margin = new System.Windows.Forms.Padding(4);
            this.DataDecorrenza.Name = "DataDecorrenza";
            this.DataDecorrenza.ReadOnly = true;
            this.DataDecorrenza.Size = new System.Drawing.Size(177, 32);
            this.DataDecorrenza.TabIndex = 47;
            // 
            // ListaBeneficiari
            // 
            this.ListaBeneficiari.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.ListaBeneficiari.Location = new System.Drawing.Point(18, 89);
            this.ListaBeneficiari.Name = "ListaBeneficiari";
            this.ListaBeneficiari.RowHeadersWidth = 51;
            this.ListaBeneficiari.RowTemplate.Height = 24;
            this.ListaBeneficiari.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.ListaBeneficiari.Size = new System.Drawing.Size(719, 439);
            this.ListaBeneficiari.TabIndex = 48;
            this.ListaBeneficiari.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.ListaBeneficiari_CellDoubleClick);
            // 
            // tabControl1
            // 
            this.tabControl1.Controls.Add(this.tabPage1);
            this.tabControl1.Controls.Add(this.tabPage2);
            this.tabControl1.Location = new System.Drawing.Point(33, 91);
            this.tabControl1.Name = "tabControl1";
            this.tabControl1.SelectedIndex = 0;
            this.tabControl1.Size = new System.Drawing.Size(890, 571);
            this.tabControl1.TabIndex = 51;
            // 
            // tabPage1
            // 
            this.tabPage1.BackColor = System.Drawing.Color.White;
            this.tabPage1.Controls.Add(this.groupBox1);
            this.tabPage1.Controls.Add(this.label5);
            this.tabPage1.Controls.Add(this.RichiestaSottoscrittore);
            this.tabPage1.Controls.Add(this.CartaIdentita);
            this.tabPage1.Controls.Add(this.DataCancellazione);
            this.tabPage1.Controls.Add(this.CodiceFiscale);
            this.tabPage1.Controls.Add(this.DataLavorazione);
            this.tabPage1.Controls.Add(this.Iban);
            this.tabPage1.Controls.Add(this.label10);
            this.tabPage1.Controls.Add(this.label1);
            this.tabPage1.Controls.Add(this.DataRimborso);
            this.tabPage1.Controls.Add(this.label4);
            this.tabPage1.Controls.Add(this.label9);
            this.tabPage1.Controls.Add(this.label11);
            this.tabPage1.Controls.Add(this.DebitoResiduo);
            this.tabPage1.Controls.Add(this.label6);
            this.tabPage1.Controls.Add(this.label8);
            this.tabPage1.Controls.Add(this.label7);
            this.tabPage1.Controls.Add(this.NuovoStatoPolizza);
            this.tabPage1.Location = new System.Drawing.Point(4, 33);
            this.tabPage1.Name = "tabPage1";
            this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage1.Size = new System.Drawing.Size(882, 534);
            this.tabPage1.TabIndex = 0;
            this.tabPage1.Text = "Dettaglio Recesso/Disdetta";
            // 
            // tabPage2
            // 
            this.tabPage2.BackColor = System.Drawing.Color.White;
            this.tabPage2.Controls.Add(this.QuotaSum);
            this.tabPage2.Controls.Add(this.ListaBeneficiari);
            this.tabPage2.Controls.Add(this.AddBeneficiary);
            this.tabPage2.Location = new System.Drawing.Point(4, 33);
            this.tabPage2.Name = "tabPage2";
            this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
            this.tabPage2.Size = new System.Drawing.Size(882, 534);
            this.tabPage2.TabIndex = 1;
            this.tabPage2.Text = "Beneficiari";
            // 
            // QuotaSum
            // 
            this.QuotaSum.AutoSize = true;
            this.QuotaSum.Font = new System.Drawing.Font("Calibri", 13.8F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.QuotaSum.Location = new System.Drawing.Point(19, 18);
            this.QuotaSum.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.QuotaSum.Name = "QuotaSum";
            this.QuotaSum.Size = new System.Drawing.Size(165, 29);
            this.QuotaSum.TabIndex = 51;
            this.QuotaSum.Text = "Somma quota: ";
            // 
            // AddBeneficiary
            // 
            this.AddBeneficiary.BackColor = System.Drawing.Color.SteelBlue;
            this.AddBeneficiary.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Center;
            this.AddBeneficiary.FlatStyle = System.Windows.Forms.FlatStyle.Flat;
            this.AddBeneficiary.ForeColor = System.Drawing.Color.Black;
            this.AddBeneficiary.Image = global::CBP.Properties.Resources.Add;
            this.AddBeneficiary.ImageAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.AddBeneficiary.Location = new System.Drawing.Point(598, 18);
            this.AddBeneficiary.Name = "AddBeneficiary";
            this.AddBeneficiary.Size = new System.Drawing.Size(139, 65);
            this.AddBeneficiary.TabIndex = 50;
            this.AddBeneficiary.Text = "Aggiungi";
            this.AddBeneficiary.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.AddBeneficiary.UseVisualStyleBackColor = false;
            this.AddBeneficiary.Click += new System.EventHandler(this.AddBeneficiary_Click);
            // 
            // CointestatrioCognome
            // 
            this.CointestatrioCognome.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.CointestatrioCognome.Location = new System.Drawing.Point(14, 183);
            this.CointestatrioCognome.Margin = new System.Windows.Forms.Padding(4);
            this.CointestatrioCognome.MaxLength = 27;
            this.CointestatrioCognome.Name = "CointestatrioCognome";
            this.CointestatrioCognome.Size = new System.Drawing.Size(312, 32);
            this.CointestatrioCognome.TabIndex = 64;
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label14.Location = new System.Drawing.Point(9, 153);
            this.label14.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(93, 26);
            this.label14.TabIndex = 66;
            this.label14.Text = "Cognome";
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label15.Location = new System.Drawing.Point(11, 231);
            this.label15.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(63, 26);
            this.label15.TabIndex = 67;
            this.label15.Text = "Nome";
            // 
            // CointestatarioNome
            // 
            this.CointestatarioNome.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.CointestatarioNome.Location = new System.Drawing.Point(14, 261);
            this.CointestatarioNome.Margin = new System.Windows.Forms.Padding(4);
            this.CointestatarioNome.MaxLength = 10;
            this.CointestatarioNome.Name = "CointestatarioNome";
            this.CointestatarioNome.Size = new System.Drawing.Size(312, 32);
            this.CointestatarioNome.TabIndex = 65;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.CointestatarioDataNascita);
            this.groupBox1.Controls.Add(this.label13);
            this.groupBox1.Controls.Add(this.CointestatarioCartaIdentita);
            this.groupBox1.Controls.Add(this.label16);
            this.groupBox1.Controls.Add(this.label14);
            this.groupBox1.Controls.Add(this.CointestatarioNome);
            this.groupBox1.Controls.Add(this.label15);
            this.groupBox1.Controls.Add(this.CointestatrioCognome);
            this.groupBox1.Controls.Add(this.label12);
            this.groupBox1.Controls.Add(this.Cointestatario);
            this.groupBox1.Location = new System.Drawing.Point(522, 16);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(344, 493);
            this.groupBox1.TabIndex = 68;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Dato cointestatario";
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label16.Location = new System.Drawing.Point(11, 310);
            this.label16.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(138, 26);
            this.label16.TabIndex = 69;
            this.label16.Text = "Data di nascita";
            // 
            // CointestatarioDataNascita
            // 
            this.CointestatarioDataNascita.CustomFormat = "dd/MM/yyyy";
            this.CointestatarioDataNascita.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.CointestatarioDataNascita.Location = new System.Drawing.Point(14, 347);
            this.CointestatarioDataNascita.Name = "CointestatarioDataNascita";
            this.CointestatarioDataNascita.Size = new System.Drawing.Size(312, 32);
            this.CointestatarioDataNascita.TabIndex = 70;
            // 
            // RecessoDisdetta
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(935, 766);
            this.Controls.Add(this.tabControl1);
            this.Controls.Add(this.DataDecorrenza);
            this.Controls.Add(this.Cancel);
            this.Controls.Add(this.CloseAndSave);
            this.Controls.Add(this.Confirm);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.NumeroPolizza);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.StatoPolizza);
            this.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.Name = "RecessoDisdetta";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Impostazione Recesso/Disdetta";
            ((System.ComponentModel.ISupportInitialize)(this.ListaBeneficiari)).EndInit();
            this.tabControl1.ResumeLayout(false);
            this.tabPage1.ResumeLayout(false);
            this.tabPage1.PerformLayout();
            this.tabPage2.ResumeLayout(false);
            this.tabPage2.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox NumeroPolizza;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox StatoPolizza;
        private System.Windows.Forms.DateTimePicker DataCancellazione;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.CheckBox RichiestaSottoscrittore;
        private System.Windows.Forms.CheckBox CartaIdentita;
        private System.Windows.Forms.CheckBox CodiceFiscale;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox Iban;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.TextBox NuovoStatoPolizza;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.TextBox DebitoResiduo;
        private System.Windows.Forms.Button CloseAndSave;
        private System.Windows.Forms.Button Confirm;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.DateTimePicker DataRimborso;
        private System.Windows.Forms.Button Cancel;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DateTimePicker DataLavorazione;
        private System.Windows.Forms.TextBox DataDecorrenza;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.CheckBox CointestatarioCartaIdentita;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.CheckBox Cointestatario;
        private System.Windows.Forms.DataGridView ListaBeneficiari;
        private System.Windows.Forms.Button AddBeneficiary;
        private System.Windows.Forms.TabControl tabControl1;
        private System.Windows.Forms.TabPage tabPage1;
        private System.Windows.Forms.TabPage tabPage2;
        private System.Windows.Forms.Label QuotaSum;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.DateTimePicker CointestatarioDataNascita;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.TextBox CointestatarioNome;
        private System.Windows.Forms.Label label15;
        private System.Windows.Forms.TextBox CointestatrioCognome;
    }
}