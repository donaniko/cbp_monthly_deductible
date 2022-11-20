namespace CBP
{
    partial class SearchBankTransactions
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.NumeroAdesione = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.StartSearch = new System.Windows.Forms.Button();
            this.NomeDistinta = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.BankTransactions = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ElencoSize = new System.Windows.Forms.ComboBox();
            this.NextAll = new System.Windows.Forms.Button();
            this.Prev = new System.Windows.Forms.Button();
            this.PrevAll = new System.Windows.Forms.Button();
            this.Next = new System.Windows.Forms.Button();
            this.PaginationLabel = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BankTransactions)).BeginInit();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.NumeroAdesione);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.StartSearch);
            this.panel1.Controls.Add(this.NomeDistinta);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Location = new System.Drawing.Point(15, 18);
            this.panel1.Margin = new System.Windows.Forms.Padding(4);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1128, 179);
            this.panel1.TabIndex = 0;
            // 
            // NumeroAdesione
            // 
            this.NumeroAdesione.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NumeroAdesione.Location = new System.Drawing.Point(515, 60);
            this.NumeroAdesione.Margin = new System.Windows.Forms.Padding(4);
            this.NumeroAdesione.Name = "NumeroAdesione";
            this.NumeroAdesione.Size = new System.Drawing.Size(416, 32);
            this.NumeroAdesione.TabIndex = 4;
            this.NumeroAdesione.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SearchText_KeyDown);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label2.Location = new System.Drawing.Point(510, 30);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(259, 26);
            this.label2.TabIndex = 3;
            this.label2.Text = "Ricerca per numero adesione";
            // 
            // StartSearch
            // 
            this.StartSearch.BackColor = System.Drawing.Color.SteelBlue;
            this.StartSearch.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.StartSearch.ForeColor = System.Drawing.Color.White;
            this.StartSearch.Location = new System.Drawing.Point(31, 114);
            this.StartSearch.Margin = new System.Windows.Forms.Padding(4);
            this.StartSearch.Name = "StartSearch";
            this.StartSearch.Size = new System.Drawing.Size(200, 42);
            this.StartSearch.TabIndex = 2;
            this.StartSearch.Text = "AVVIA RICERCA";
            this.StartSearch.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.StartSearch.UseVisualStyleBackColor = false;
            this.StartSearch.Click += new System.EventHandler(this.StartSearch_Click);
            // 
            // NomeDistinta
            // 
            this.NomeDistinta.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NomeDistinta.Location = new System.Drawing.Point(31, 60);
            this.NomeDistinta.Margin = new System.Windows.Forms.Padding(4);
            this.NomeDistinta.Name = "NomeDistinta";
            this.NomeDistinta.Size = new System.Drawing.Size(416, 32);
            this.NomeDistinta.TabIndex = 1;
            this.NomeDistinta.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SearchText_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label1.Location = new System.Drawing.Point(26, 30);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(228, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ricerca per nome distinta";
            // 
            // BankTransactions
            // 
            this.BankTransactions.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.BankTransactions.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.BankTransactions.Location = new System.Drawing.Point(15, 205);
            this.BankTransactions.Margin = new System.Windows.Forms.Padding(4);
            this.BankTransactions.MultiSelect = false;
            this.BankTransactions.Name = "BankTransactions";
            this.BankTransactions.ReadOnly = true;
            this.BankTransactions.RowHeadersWidth = 51;
            this.BankTransactions.RowTemplate.Height = 24;
            this.BankTransactions.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.BankTransactions.Size = new System.Drawing.Size(1128, 705);
            this.BankTransactions.TabIndex = 0;
            this.BankTransactions.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.BankTransactions_CellContentClick);
            this.BankTransactions.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.BankTransactions_CellDoubleClick);
            // 
            // panel2
            // 
            this.panel2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel2.BackColor = System.Drawing.Color.LightGray;
            this.panel2.Controls.Add(this.ElencoSize);
            this.panel2.Controls.Add(this.NextAll);
            this.panel2.Controls.Add(this.Prev);
            this.panel2.Controls.Add(this.PrevAll);
            this.panel2.Controls.Add(this.Next);
            this.panel2.Controls.Add(this.PaginationLabel);
            this.panel2.Location = new System.Drawing.Point(15, 918);
            this.panel2.Margin = new System.Windows.Forms.Padding(4);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1129, 68);
            this.panel2.TabIndex = 1;
            // 
            // ElencoSize
            // 
            this.ElencoSize.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.ElencoSize.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ElencoSize.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.ElencoSize.FormattingEnabled = true;
            this.ElencoSize.Location = new System.Drawing.Point(913, 21);
            this.ElencoSize.Margin = new System.Windows.Forms.Padding(4);
            this.ElencoSize.Name = "ElencoSize";
            this.ElencoSize.Size = new System.Drawing.Size(199, 32);
            this.ElencoSize.TabIndex = 5;
            this.ElencoSize.SelectedIndexChanged += new System.EventHandler(this.ElencoSize_SelectedIndexChanged);
            // 
            // NextAll
            // 
            this.NextAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.NextAll.Location = new System.Drawing.Point(760, 12);
            this.NextAll.Margin = new System.Windows.Forms.Padding(4);
            this.NextAll.Name = "NextAll";
            this.NextAll.Size = new System.Drawing.Size(48, 48);
            this.NextAll.TabIndex = 4;
            this.NextAll.Text = ">>";
            this.NextAll.UseVisualStyleBackColor = true;
            this.NextAll.Click += new System.EventHandler(this.NextAll_Click);
            // 
            // Prev
            // 
            this.Prev.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Prev.Location = new System.Drawing.Point(336, 12);
            this.Prev.Margin = new System.Windows.Forms.Padding(4);
            this.Prev.Name = "Prev";
            this.Prev.Size = new System.Drawing.Size(48, 48);
            this.Prev.TabIndex = 3;
            this.Prev.Text = "<";
            this.Prev.UseVisualStyleBackColor = true;
            this.Prev.Click += new System.EventHandler(this.Prev_Click);
            // 
            // PrevAll
            // 
            this.PrevAll.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PrevAll.Location = new System.Drawing.Point(281, 12);
            this.PrevAll.Margin = new System.Windows.Forms.Padding(4);
            this.PrevAll.Name = "PrevAll";
            this.PrevAll.Size = new System.Drawing.Size(48, 48);
            this.PrevAll.TabIndex = 2;
            this.PrevAll.Text = "<<";
            this.PrevAll.UseVisualStyleBackColor = true;
            this.PrevAll.Click += new System.EventHandler(this.PrevAll_Click);
            // 
            // Next
            // 
            this.Next.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.Next.Location = new System.Drawing.Point(705, 12);
            this.Next.Margin = new System.Windows.Forms.Padding(4);
            this.Next.Name = "Next";
            this.Next.Size = new System.Drawing.Size(48, 48);
            this.Next.TabIndex = 1;
            this.Next.Text = ">";
            this.Next.UseVisualStyleBackColor = true;
            this.Next.Click += new System.EventHandler(this.Next_Click);
            // 
            // PaginationLabel
            // 
            this.PaginationLabel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.PaginationLabel.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.PaginationLabel.Location = new System.Drawing.Point(392, 12);
            this.PaginationLabel.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.PaginationLabel.Name = "PaginationLabel";
            this.PaginationLabel.Size = new System.Drawing.Size(305, 48);
            this.PaginationLabel.TabIndex = 0;
            this.PaginationLabel.Text = "Pagina: ";
            this.PaginationLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // SearchBankTransactions
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(1159, 995);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.BankTransactions);
            this.Controls.Add(this.panel1);
            this.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "SearchBankTransactions";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Archivio Distinte";
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.BankTransactions)).EndInit();
            this.panel2.ResumeLayout(false);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button StartSearch;
        private System.Windows.Forms.TextBox NomeDistinta;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView BankTransactions;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label PaginationLabel;
        private System.Windows.Forms.Button NextAll;
        private System.Windows.Forms.Button Prev;
        private System.Windows.Forms.Button PrevAll;
        private System.Windows.Forms.Button Next;
        private System.Windows.Forms.ComboBox ElencoSize;
        private System.Windows.Forms.TextBox NumeroAdesione;
        private System.Windows.Forms.Label label2;
    }
}