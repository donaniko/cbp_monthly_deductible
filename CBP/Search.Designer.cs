namespace CBP
{
    partial class Search
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
            this.StartSearch = new System.Windows.Forms.Button();
            this.SearchText = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Policies = new System.Windows.Forms.DataGridView();
            this.panel2 = new System.Windows.Forms.Panel();
            this.ElencoSize = new System.Windows.Forms.ComboBox();
            this.NextAll = new System.Windows.Forms.Button();
            this.Prev = new System.Windows.Forms.Button();
            this.PrevAll = new System.Windows.Forms.Button();
            this.Next = new System.Windows.Forms.Button();
            this.PaginationLabel = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.NuovoRecesso = new System.Windows.Forms.Button();
            this.Recesso = new System.Windows.Forms.Button();
            this.FirstName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.LastName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.StatoPolizza = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.NumeroPolizza = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.Tipologia = new System.Windows.Forms.ComboBox();
            ((System.ComponentModel.ISupportInitialize)(this.Policies)).BeginInit();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.SuspendLayout();
            // 
            // StartSearch
            // 
            this.StartSearch.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.StartSearch.Location = new System.Drawing.Point(473, 107);
            this.StartSearch.Margin = new System.Windows.Forms.Padding(4);
            this.StartSearch.Name = "StartSearch";
            this.StartSearch.Size = new System.Drawing.Size(200, 54);
            this.StartSearch.TabIndex = 2;
            this.StartSearch.Text = "AVVIA RICERCA";
            this.StartSearch.UseVisualStyleBackColor = true;
            this.StartSearch.Click += new System.EventHandler(this.StartSearch_Click);
            // 
            // SearchText
            // 
            this.SearchText.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.SearchText.Location = new System.Drawing.Point(15, 54);
            this.SearchText.Margin = new System.Windows.Forms.Padding(4);
            this.SearchText.Name = "SearchText";
            this.SearchText.Size = new System.Drawing.Size(395, 32);
            this.SearchText.TabIndex = 1;
            this.SearchText.KeyDown += new System.Windows.Forms.KeyEventHandler(this.SearchText_KeyDown);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label1.Location = new System.Drawing.Point(10, 24);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(400, 26);
            this.label1.TabIndex = 0;
            this.label1.Text = "Ricerca per testo (n° pratica, cognome, nome)";
            // 
            // Policies
            // 
            this.Policies.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.Policies.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.Policies.Location = new System.Drawing.Point(15, 196);
            this.Policies.Margin = new System.Windows.Forms.Padding(4);
            this.Policies.MultiSelect = false;
            this.Policies.Name = "Policies";
            this.Policies.ReadOnly = true;
            this.Policies.RowHeadersWidth = 51;
            this.Policies.RowTemplate.Height = 24;
            this.Policies.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.Policies.Size = new System.Drawing.Size(785, 696);
            this.Policies.TabIndex = 0;
            this.Policies.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.Policies_CellContentClick);
            this.Policies.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.Policies_CellDoubleClick);
            this.Policies.SelectionChanged += new System.EventHandler(this.Policies_SelectionChanged);
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
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.NuovoRecesso);
            this.panel3.Controls.Add(this.Recesso);
            this.panel3.Controls.Add(this.FirstName);
            this.panel3.Controls.Add(this.label5);
            this.panel3.Controls.Add(this.LastName);
            this.panel3.Controls.Add(this.label4);
            this.panel3.Controls.Add(this.StatoPolizza);
            this.panel3.Controls.Add(this.label3);
            this.panel3.Controls.Add(this.NumeroPolizza);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Location = new System.Drawing.Point(807, 196);
            this.panel3.Margin = new System.Windows.Forms.Padding(4);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(336, 695);
            this.panel3.TabIndex = 2;
            // 
            // NuovoRecesso
            // 
            this.NuovoRecesso.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.NuovoRecesso.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NuovoRecesso.Location = new System.Drawing.Point(15, 636);
            this.NuovoRecesso.Margin = new System.Windows.Forms.Padding(4);
            this.NuovoRecesso.Name = "NuovoRecesso";
            this.NuovoRecesso.Size = new System.Drawing.Size(305, 34);
            this.NuovoRecesso.TabIndex = 10;
            this.NuovoRecesso.Text = "Nuova polizza manuale";
            this.NuovoRecesso.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.NuovoRecesso.UseVisualStyleBackColor = true;
            this.NuovoRecesso.Click += new System.EventHandler(this.NuovoRecesso_Click);
            // 
            // Recesso
            // 
            this.Recesso.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Recesso.Location = new System.Drawing.Point(21, 432);
            this.Recesso.Margin = new System.Windows.Forms.Padding(4);
            this.Recesso.Name = "Recesso";
            this.Recesso.Size = new System.Drawing.Size(304, 35);
            this.Recesso.TabIndex = 3;
            this.Recesso.Text = "Imposta Recesso/Disdetta";
            this.Recesso.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            this.Recesso.UseVisualStyleBackColor = true;
            this.Recesso.Click += new System.EventHandler(this.Recesso_Click);
            // 
            // FirstName
            // 
            this.FirstName.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.FirstName.Location = new System.Drawing.Point(21, 376);
            this.FirstName.Margin = new System.Windows.Forms.Padding(4);
            this.FirstName.Name = "FirstName";
            this.FirstName.ReadOnly = true;
            this.FirstName.Size = new System.Drawing.Size(304, 32);
            this.FirstName.TabIndex = 9;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label5.Location = new System.Drawing.Point(16, 346);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(63, 26);
            this.label5.TabIndex = 8;
            this.label5.Text = "Nome";
            // 
            // LastName
            // 
            this.LastName.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.LastName.Location = new System.Drawing.Point(20, 274);
            this.LastName.Margin = new System.Windows.Forms.Padding(4);
            this.LastName.Name = "LastName";
            this.LastName.ReadOnly = true;
            this.LastName.Size = new System.Drawing.Size(304, 32);
            this.LastName.TabIndex = 7;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label4.Location = new System.Drawing.Point(16, 244);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(93, 26);
            this.label4.TabIndex = 6;
            this.label4.Text = "Cognome";
            // 
            // StatoPolizza
            // 
            this.StatoPolizza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.StatoPolizza.Location = new System.Drawing.Point(19, 172);
            this.StatoPolizza.Margin = new System.Windows.Forms.Padding(4);
            this.StatoPolizza.Name = "StatoPolizza";
            this.StatoPolizza.ReadOnly = true;
            this.StatoPolizza.Size = new System.Drawing.Size(304, 32);
            this.StatoPolizza.TabIndex = 5;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label3.Location = new System.Drawing.Point(15, 142);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(120, 26);
            this.label3.TabIndex = 4;
            this.label3.Text = "Stato polizza";
            // 
            // NumeroPolizza
            // 
            this.NumeroPolizza.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.NumeroPolizza.Location = new System.Drawing.Point(20, 70);
            this.NumeroPolizza.Margin = new System.Windows.Forms.Padding(4);
            this.NumeroPolizza.Name = "NumeroPolizza";
            this.NumeroPolizza.ReadOnly = true;
            this.NumeroPolizza.Size = new System.Drawing.Size(304, 32);
            this.NumeroPolizza.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label2.Location = new System.Drawing.Point(16, 40);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(144, 26);
            this.label2.TabIndex = 2;
            this.label2.Text = "Numero polizza";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label6.Location = new System.Drawing.Point(10, 100);
            this.label6.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(185, 26);
            this.label6.TabIndex = 3;
            this.label6.Text = "Ricerca per tipologia";
            // 
            // Tipologia
            // 
            this.Tipologia.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.Tipologia.FormattingEnabled = true;
            this.Tipologia.Items.AddRange(new object[] {
            "Tutte le pratiche",
            "Recessi/Disdette non completate",
            "Recessi/Disdette completate"});
            this.Tipologia.Location = new System.Drawing.Point(15, 129);
            this.Tipologia.Name = "Tipologia";
            this.Tipologia.Size = new System.Drawing.Size(395, 32);
            this.Tipologia.TabIndex = 4;
            // 
            // Search
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.LightSteelBlue;
            this.ClientSize = new System.Drawing.Size(1159, 995);
            this.Controls.Add(this.Tipologia);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.SearchText);
            this.Controls.Add(this.StartSearch);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.Policies);
            this.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Margin = new System.Windows.Forms.Padding(4);
            this.Name = "Search";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Archivio Polizze";
            ((System.ComponentModel.ISupportInitialize)(this.Policies)).EndInit();
            this.panel2.ResumeLayout(false);
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button StartSearch;
        private System.Windows.Forms.TextBox SearchText;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView Policies;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label PaginationLabel;
        private System.Windows.Forms.Button NextAll;
        private System.Windows.Forms.Button Prev;
        private System.Windows.Forms.Button PrevAll;
        private System.Windows.Forms.Button Next;
        private System.Windows.Forms.ComboBox ElencoSize;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.TextBox FirstName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox LastName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox StatoPolizza;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox NumeroPolizza;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button NuovoRecesso;
        private System.Windows.Forms.Button Recesso;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ComboBox Tipologia;
    }
}