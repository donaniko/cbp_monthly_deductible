namespace CBP
{
    partial class BeneficiaryEditor
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
            this.FirstName = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.LastName = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.Iban = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.Share = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.BirthDate = new System.Windows.Forms.DateTimePicker();
            this.label5 = new System.Windows.Forms.Label();
            this.Save = new System.Windows.Forms.Button();
            this.Delete = new System.Windows.Forms.Button();
            this.ErrorText = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // FirstName
            // 
            this.FirstName.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.FirstName.Location = new System.Drawing.Point(48, 67);
            this.FirstName.Margin = new System.Windows.Forms.Padding(4);
            this.FirstName.MaxLength = 27;
            this.FirstName.Name = "FirstName";
            this.FirstName.Size = new System.Drawing.Size(428, 32);
            this.FirstName.TabIndex = 0;
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label4.Location = new System.Drawing.Point(43, 37);
            this.label4.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 26);
            this.label4.TabIndex = 50;
            this.label4.Text = "Nome";
            // 
            // LastName
            // 
            this.LastName.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.LastName.Location = new System.Drawing.Point(48, 155);
            this.LastName.Margin = new System.Windows.Forms.Padding(4);
            this.LastName.MaxLength = 27;
            this.LastName.Name = "LastName";
            this.LastName.Size = new System.Drawing.Size(428, 32);
            this.LastName.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label1.Location = new System.Drawing.Point(43, 125);
            this.label1.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(93, 26);
            this.label1.TabIndex = 52;
            this.label1.Text = "Cognome";
            // 
            // Iban
            // 
            this.Iban.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Iban.Location = new System.Drawing.Point(48, 341);
            this.Iban.Margin = new System.Windows.Forms.Padding(4);
            this.Iban.MaxLength = 27;
            this.Iban.Name = "Iban";
            this.Iban.Size = new System.Drawing.Size(428, 32);
            this.Iban.TabIndex = 4;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label2.Location = new System.Drawing.Point(43, 311);
            this.label2.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(53, 26);
            this.label2.TabIndex = 54;
            this.label2.Text = "IBAN";
            // 
            // Share
            // 
            this.Share.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.Share.Location = new System.Drawing.Point(284, 244);
            this.Share.Margin = new System.Windows.Forms.Padding(4);
            this.Share.MaxLength = 27;
            this.Share.Name = "Share";
            this.Share.Size = new System.Drawing.Size(192, 32);
            this.Share.TabIndex = 3;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label3.Location = new System.Drawing.Point(279, 214);
            this.label3.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(207, 26);
            this.label3.TabIndex = 56;
            this.label3.Text = "Quota ( . per decimali )";
            // 
            // BirthDate
            // 
            this.BirthDate.CustomFormat = "dd/MM/yyyy";
            this.BirthDate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.BirthDate.Location = new System.Drawing.Point(48, 244);
            this.BirthDate.Name = "BirthDate";
            this.BirthDate.Size = new System.Drawing.Size(192, 32);
            this.BirthDate.TabIndex = 2;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.label5.Location = new System.Drawing.Point(43, 214);
            this.label5.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(138, 26);
            this.label5.TabIndex = 58;
            this.label5.Text = "Data di nascita";
            // 
            // Save
            // 
            this.Save.BackColor = System.Drawing.Color.SteelBlue;
            this.Save.ForeColor = System.Drawing.Color.White;
            this.Save.Location = new System.Drawing.Point(284, 496);
            this.Save.Name = "Save";
            this.Save.Size = new System.Drawing.Size(95, 54);
            this.Save.TabIndex = 5;
            this.Save.Text = "Salva";
            this.Save.UseVisualStyleBackColor = false;
            this.Save.Click += new System.EventHandler(this.Save_Click);
            // 
            // Delete
            // 
            this.Delete.BackColor = System.Drawing.Color.DarkRed;
            this.Delete.ForeColor = System.Drawing.Color.White;
            this.Delete.Location = new System.Drawing.Point(381, 496);
            this.Delete.Name = "Delete";
            this.Delete.Size = new System.Drawing.Size(95, 54);
            this.Delete.TabIndex = 6;
            this.Delete.Text = "Elimina";
            this.Delete.UseVisualStyleBackColor = false;
            this.Delete.Click += new System.EventHandler(this.Delete_Click);
            // 
            // ErrorText
            // 
            this.ErrorText.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.ErrorText.ForeColor = System.Drawing.Color.DarkRed;
            this.ErrorText.Location = new System.Drawing.Point(43, 391);
            this.ErrorText.Margin = new System.Windows.Forms.Padding(4, 0, 4, 0);
            this.ErrorText.Name = "ErrorText";
            this.ErrorText.Size = new System.Drawing.Size(433, 88);
            this.ErrorText.TabIndex = 61;
            this.ErrorText.Text = "errore";
            // 
            // BeneficiaryEditor
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(10F, 24F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(521, 574);
            this.Controls.Add(this.ErrorText);
            this.Controls.Add(this.Delete);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.BirthDate);
            this.Controls.Add(this.Share);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.Iban);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.LastName);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.FirstName);
            this.Controls.Add(this.label4);
            this.Font = new System.Drawing.Font("Calibri Light", 12.2F);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Margin = new System.Windows.Forms.Padding(4);
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "BeneficiaryEditor";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "BeneficiaryEditor";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox FirstName;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox LastName;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox Iban;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox Share;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.DateTimePicker BirthDate;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Button Save;
        private System.Windows.Forms.Button Delete;
        private System.Windows.Forms.Label ErrorText;
    }
}