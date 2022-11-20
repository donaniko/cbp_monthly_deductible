using CBP.Helpers;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CBP
{
    public partial class BeneficiaryEditor : Form
    {
        private static BeneficiaryInfo Current { get; set; }
        private BeneficiaryEditor()
        {
            InitializeComponent();
        }

        public static BeneficiaryInfo CreateNew(Form owner, int tacticalRowId)
        {
            var form = new BeneficiaryEditor();

            form.ErrorText.Text = "";
            form.Delete.Hide();
            form.Save.Left = form.Delete.Left;
            Current = new BeneficiaryInfo();
            Current.TacticalRowId = tacticalRowId;

            if (form.ShowDialog(owner) == DialogResult.OK)
            {
                return Current;
            }

            return null;
        }

        public static BeneficiaryInfo Edit(Form owner, int beneficiaryId)
        {
            var form = new BeneficiaryEditor();

            form.ErrorText.Text = "";
            Current = HelperClass.GetBeneficiary(beneficiaryId);

            form.FirstName.Text = Current.FirstName;
            form.LastName.Text = Current.LastName;
            if (Current.BirthDate.HasValue)
            {
                form.BirthDate.Value = Current.BirthDate.Value;
            }
            form.Iban.Text = Current.Iban;
            form.Share.Text = Current.Share.ToString().Replace(",", ".");

            if (form.ShowDialog(owner) == DialogResult.OK)
            {
                return Current;
            }

            return null;
        }

        private void Save_Click(object sender, EventArgs e)
        {
            ErrorText.Text = "";

            Current.FirstName = FirstName.Text;
            Current.LastName = LastName.Text;
            Current.BirthDate = BirthDate.Value;
            Current.Iban = Iban.Text;

            if (string.IsNullOrWhiteSpace(Current.FirstName))
            {
                ErrorText.Text += "Inserire il nome - ";
            }

            if (string.IsNullOrWhiteSpace(Current.LastName))
            {
                ErrorText.Text += "Inserire il cognome - ";
            }

            if (string.IsNullOrWhiteSpace(Current.Iban))
            {
                ErrorText.Text += "Inserire un IBAN - ";
            }

            var quota = 0D;
            if (!double.TryParse(Share.Text.Replace(",", "."), 
                System.Globalization.NumberStyles.AllowDecimalPoint, 
                System.Globalization.NumberFormatInfo.InvariantInfo, 
                out quota))
            {
                ErrorText.Text += "Inserire un valore quota valido (usare il punto per i decimali) - ";
            }
            Current.Share = quota;

            if (!string.IsNullOrEmpty(ErrorText.Text))
            {
                return;
            }

            try
            {
                if (Current.Id == 0)
                {
                    Current = HelperClass.PostBeneficiary(Current);
                }
                else
                {
                    Current = HelperClass.PutBeneficiary(Current);
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                HelperClass.ShowError(this, ex.Message);
            }
        }

        private void Delete_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "Sicuro di voler eliminare il beneficiario corrente?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                HelperClass.DeleteBeneficiary(Current.Id);
                DialogResult = DialogResult.OK;
            }
        }
    }
}
