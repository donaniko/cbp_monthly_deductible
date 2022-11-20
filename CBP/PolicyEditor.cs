using CBP.DataLayer;
using CBP.Helpers;
using OfficeOpenXml.FormulaParsing.Excel.Functions.RefAndLookup;
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
    public partial class PolicyEditor : Form
    {
        /*
            PolicyNumber
            TransactionType
            DateOfSignature

            Surname
            FirstName
            Birth
            Sex

            Street
            Location
            Provincia
            ZipCode

            PremioLordo
            PremioNetto
            Tassi

            PremioNettoDanni
            CommissioniDanni

            PremioVita
            CommissioniVita


            Prodotto
            Convenzione
            Tariff

         */


        public TacticalRowInfo Policy;

        private void InitForm()
        {
            HelperClass.InitNumericTextBox(PremioLordo, 0, false);
            HelperClass.InitNumericTextBox(PremioNetto, 2, true);
            HelperClass.InitNumericTextBox(Tassi, 2, true);
            HelperClass.InitNumericTextBox(PremioNettoDanni, 2, true);
            HelperClass.InitNumericTextBox(CommissioniDanni, 2, true);
            HelperClass.InitNumericTextBox(PremioVita, 2, true);
            HelperClass.InitNumericTextBox(CommissioniVita, 2, true);
            HelperClass.InitNumericTextBox(InsuredSum, 2, true);

            DataDecorrenza.Value = HelperClass.GetDateNow();
            ExpirationDate.Value = HelperClass.GetDateNow();
            BirthDate.Value = HelperClass.GetDateNow();
            DataCancellazione.Value = DataCancellazione.MinDate;
        }

        public PolicyEditor(TacticalRowInfo policy)
        {
            InitializeComponent();
            InitForm();

            Policy = policy;
            LoadValues();
        }

        public PolicyEditor()
        {
            InitializeComponent();
            InitForm();

            Policy = new TacticalRowInfo();
            Policy.Sex = 1;
            LoadValues();
        }

        private void LoadValues()
        {
            NumeroPolizza.Text = Policy.PolicyNumber == 0 ? "" : Policy.PolicyNumber.ToString();
            StatoPolizza.Text = Policy.TransactionType ?? "";
            if (Policy.DateOfSignature.HasValue)
            {
                DataDecorrenza.Value = Policy.DateOfSignature.Value;
            }

            if (Policy.ExpirationDate.HasValue)
            {
                ExpirationDate.Value = Policy.ExpirationDate.Value;
            }

            LastName.Text = Policy.Surname ?? "";
            FirstName.Text = Policy.FirstName ?? "";

            if (Policy.Id > 0)
            {
                BirthDate.Value = Policy.Birth;
            }

            if (Policy.Sex > 0)
            {
                Gender.SelectedIndex = Policy.Sex - 1;
            }

            Street.Text = Policy.Street ?? "";
            Localita.Text = Policy.Location ?? "";
            Provincia.Text = Policy.Provincia ?? "";
            Cap.Text = Policy.ZipCode ?? "";
            Santander.Checked = Policy.Santander;

            if (Policy.DateOfCancellation.HasValue)
            {
                DataCancellazione.Value = Policy.DateOfCancellation.Value;
            }

            if (Policy.Id > 0)
            {
                PremioLordo.Text = HelperClass.FormatNumber(Policy.PremioLordo, 2);
                if (Policy.PremioNetto.HasValue)
                {
                    PremioNetto.Text = HelperClass.FormatNumber(Policy.PremioNetto.Value, 2);
                }

                if (Policy.Tassi.HasValue)
                {
                    Tassi.Text = HelperClass.FormatNumber(Policy.Tassi.Value, 2);
                }

                if (Policy.PremioNettoDanni.HasValue)
                {
                    PremioNettoDanni.Text = HelperClass.FormatNumber(Policy.PremioNettoDanni.Value, 2);
                }

                if (Policy.CommissioniDanni.HasValue)
                {
                    CommissioniDanni.Text = HelperClass.FormatNumber(Policy.CommissioniDanni.Value, 2);
                }

                if (Policy.PremioVita.HasValue)
                {
                    PremioVita.Text = HelperClass.FormatNumber(Policy.PremioVita.Value, 2);
                }

                if (Policy.CommissioniVita.HasValue)
                {
                    CommissioniVita.Text = HelperClass.FormatNumber(Policy.CommissioniVita.Value, 2);
                }

                if (Policy.InsuredSum.HasValue)
                {
                    InsuredSum.Text = HelperClass.FormatNumber(Policy.InsuredSum.Value, 2);
                }                
            }

            Tariffa.Text = Policy.Tariff ?? "";
            HE.Text = Policy.HE ?? "";
            Iban.Text = Policy.Iban ?? "";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void Confirm_Click(object sender, EventArgs e)
        {
            try
            {
                using (var uow = new UnitOfWork())
                {
                    TacticalRow entity = null;
                    if (Policy.Id == 0)
                    {
                        entity = new TacticalRow();
                    }
                    else
                    {
                        entity = uow.TacticalRowRepository.Get(tr => tr.Id == Policy.Id).First();
                    }

                    entity.PolicyNumber = HelperClass.GetIntValue(NumeroPolizza);
                    entity.TransactionType = StatoPolizza.Text;
                    entity.DateOfSignature = DataDecorrenza.Value;
                    entity.ExpirationDate = ExpirationDate.Value;
                    entity.Surname = LastName.Text;
                    entity.FirstName = FirstName.Text;
                    entity.Birth = BirthDate.Value;
                    entity.Sex = Gender.SelectedIndex + 1;
                    entity.Street = Street.Text;
                    entity.Location = Localita.Text;
                    entity.Provincia = Provincia.Text;
                    entity.ZipCode = Cap.Text;
                    entity.PremioLordo = HelperClass.GetDoubleValue(PremioLordo);
                    entity.PremioNetto = HelperClass.GetDoubleNullableValue(PremioNetto);
                    entity.Tassi = HelperClass.GetDoubleNullableValue(Tassi);
                    entity.PremioNettoDanni = HelperClass.GetDoubleNullableValue(PremioNettoDanni);
                    entity.CommissioniDanni = HelperClass.GetDoubleNullableValue(CommissioniDanni);
                    entity.PremioVita = HelperClass.GetDoubleNullableValue(PremioVita);
                    entity.CommissioniVita = HelperClass.GetDoubleNullableValue(CommissioniVita);
                    entity.InsuredSum = HelperClass.GetDoubleNullableValue(InsuredSum);
                    entity.Tariff = Tariffa.Text;
                    entity.HE = HE.Text;
                    entity.Iban = Iban.Text;
                    entity.Santander = Santander.Checked;
                    if (entity.Santander)
                    {
                        entity.RecessoDisdettaCompleto = true;
                    }
                    if (DataCancellazione.Value > DataCancellazione.MinDate)
                    {
                        entity.DateOfCancellation = DataCancellazione.Value;
                    }

                    if (entity.Id == 0)
                    {
                        uow.TacticalRowRepository.Insert(entity);
                    }
                    else
                    {
                        uow.TacticalRowRepository.Update(entity);
                    }

                    uow.Save();

                    Policy = new TacticalRowInfo(entity, uow);
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                ShowError("Errore nel salvataggio: " + ex.Message);
            }
        }

        private void ShowError(string err)
        {
            MessageBox.Show(this, err, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }
}
