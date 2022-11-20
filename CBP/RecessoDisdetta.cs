using CBP.DataLayer;
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
    public partial class RecessoDisdetta : Form
    {
        public int PolicyId;
        public bool ErrorQuotaSum;

        public RecessoDisdetta()
        {
            InitializeComponent();
        }

        public RecessoDisdetta(int policyId)
        {
            InitializeComponent();
            InitForm();

            PolicyId = policyId;
            LoadValues();
        }

        private void InitForm()
        {
            HelperClass.InitNumericTextBox(DebitoResiduo, 0, true);
            DataLavorazione.Value = HelperClass.GetDateNow();
            DataRimborso.Value = HelperClass.GetDateNow();
            DataCancellazione.Value = HelperClass.GetDateNow();
        }

        private void LoadValues()
        {
            using (var uow = new UnitOfWork())
            {
                var Policy = uow.TacticalRowRepository.Get(p => p.Id == PolicyId).First();

                NumeroPolizza.Text = Policy.PolicyNumber == 0 ? "" : Policy.PolicyNumber.ToString();
                StatoPolizza.Text = Policy.TransactionType ?? "";
                NuovoStatoPolizza.Text = Policy.TransactionType ?? "";

                if (Policy.DateOfSignature.HasValue)
                {
                    DataDecorrenza.Text = Policy.DateOfSignature.Value.ToString("dd/MM/yyyy");
                }

                var estinzioneParziale = Policy.CapitaleResiduo.HasValue;
                if (estinzioneParziale)
                {
                    if (Policy.DateOfCancellation3.HasValue)
                    {
                        DataCancellazione.Value = Policy.DateOfCancellation3.Value;
                    }
                }
                else
                {
                    if (Policy.DateOfCancellation.HasValue)
                    {
                        DataCancellazione.Value = Policy.DateOfCancellation.Value;
                    }
                }

                DataLavorazione.Value = DateTime.Now;

                if (Policy.DataRimborso.HasValue)
                {
                    DataRimborso.Value = Policy.DataRimborso.Value;
                }

                RichiestaSottoscrittore.Checked = Policy.RichiestaSottoscrittore == true;
                CartaIdentita.Checked = Policy.CartaIdentita == true;
                CodiceFiscale.Checked = Policy.CodiceFiscale == true;
                Cointestatario.Checked = Policy.Cointestatario == true;
                CointestatarioCartaIdentita.Checked = Policy.CointestatarioCartaIdentita == true;
                CointestatrioCognome.Text = Policy.CognomeCoIntestario;
                CointestatarioNome.Text = Policy.NomeCoIntestatario;
                if (Policy.CointestatarioDataNascita.HasValue)
                {
                    CointestatarioDataNascita.Value = Policy.CointestatarioDataNascita.Value;
                }
                Iban.Text = Policy.Iban ?? "";

                if (Policy.DebitoResiduo.HasValue)
                {
                    DebitoResiduo.Text = HelperClass.FormatNumber(Policy.DebitoResiduo.Value, 2);
                }

                
                ImpostaNuovoStatoPolizza();
            }

            LoadBeneficiaries();
        }

        private void LoadBeneficiaries()
        {
            var beneficiaries = new List<BeneficiaryInfo>();
            var quota = 0D;

            using (var uow = new UnitOfWork())
            {
                foreach (var beneficiary in uow.BeneficiaryRepository.Get(b => b.TacticalRowId == PolicyId))
                {
                    quota += beneficiary.Share;
                    beneficiaries.Add(new BeneficiaryInfo(beneficiary, uow));
                }
            }

            QuotaSum.Text = string.Format("Somma delle quote: {0}", quota.ToString());

            ErrorQuotaSum = false;
            if (quota != 100)
            {
                if (beneficiaries.Count > 0)
                {
                    ErrorQuotaSum = true;
                }
                QuotaSum.ForeColor = Color.DarkRed;
            }
            else
            {
                QuotaSum.ForeColor = Color.Black;
            }

            ListaBeneficiari.DataSource = beneficiaries;
        }

        private void ImpostaNuovoStatoPolizza()
        {
            if (StatoPolizza.Text == HelperClass.TRANSATION_TYPE_NEWBUSINESS)
            {
                if (!string.IsNullOrEmpty(DataDecorrenza.Text))
                {
                    var dataDecorrenza = new DateTime(int.Parse(DataDecorrenza.Text.Substring(6, 4)), int.Parse(DataDecorrenza.Text.Substring(3, 2)), int.Parse(DataDecorrenza.Text.Substring(0, 2)));
                    var days = (DataCancellazione.Value - dataDecorrenza).TotalDays;

                    if (days < 120)
                    {
                        NuovoStatoPolizza.Text = HelperClass.TRANSATION_TYPE_RECESSO;
                    }
                    else
                    {
                        NuovoStatoPolizza.Text = HelperClass.TRANSATION_TYPE_DISDETTA;
                    }
                }
            }
        }

        private void DataCancellazione_ValueChanged(object sender, EventArgs e)
        {
            ImpostaNuovoStatoPolizza();
        }

        private void NuovoStatoPolizza_Leave(object sender, EventArgs e)
        {
            if (NuovoStatoPolizza.Text != HelperClass.TRANSATION_TYPE_RECESSO && NuovoStatoPolizza.Text != HelperClass.TRANSATION_TYPE_DISDETTA)
            {
                ShowInfo("Impossibile impostare uno stato diverso da E, D");
                ImpostaNuovoStatoPolizza();
            }
        }

        private void ShowInfo(string msg)
        {
            MessageBox.Show(this, msg, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(string err)
        {
            MessageBox.Show(this, err, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private void CloseAndSave_Click(object sender, EventArgs e)
        {
            try
            {
                using (var uow = new UnitOfWork())
                {
                    var Policy = uow.TacticalRowRepository.Get(p => p.Id == PolicyId).First();
                    CopyValues(Policy);

                    uow.TacticalRowRepository.Update(Policy);
                    uow.Save();
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                ShowError("Errore nel salvataggio: " + ex.Message);
            }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
            Close();
        }

        private void Confirm_Click(object sender, EventArgs e)
        {
            if (ErrorQuotaSum)
            {
                ShowInfo("Impossibile procedere, la somma delle quote dei beneficiari è diversa da 100");
                return;
            }

            if (string.IsNullOrEmpty(DataDecorrenza.Text))
            {
                ShowInfo("Impossibile procedere, controlla la data di decorrenza della polizza");
                return;
            }

            if (!RichiestaSottoscrittore.Checked)
            {
                ShowInfo("Impossibile procedere, controlla la presenza della richiesta al sottoscrittore");
                return;
            }

            if (!CartaIdentita.Checked)
            {
                ShowInfo("Impossibile procedere, controlla la presenza della carta d'identità");
                return;
            }

            if (!CodiceFiscale.Checked)
            {
                ShowInfo("Impossibile procedere, controlla la presenza del codice fiscale");
                return;
            }

            if (Cointestatario.Checked
                && (!CointestatarioCartaIdentita.Checked
                     || string.IsNullOrEmpty(CointestatrioCognome.Text)
                     || string.IsNullOrEmpty(CointestatarioNome.Text))
               )
            {
                ShowInfo("Impossibile procedere, per una pratica con cointestario bisogna avere tutti i suoi dati");
                return;
            }

            if (string.IsNullOrEmpty(Iban.Text))
            {
                ShowInfo("Impossibile procedere, bisogna inserire l'IBAN");
                return;
            }

            if (NuovoStatoPolizza.Text.ToUpper() == "DS" && string.IsNullOrEmpty(DebitoResiduo.Text))
            {
                ShowInfo("Impossibile procedere, per le disdette bisogna inserire il debito residuo");
                return;
            }

            try
            {
                using (var uow = new UnitOfWork())
                {
                    var Policy = uow.TacticalRowRepository.Get(p => p.Id == PolicyId).First();
                    Policy.RecessoDisdettaCompleto = true;
                    CopyValues(Policy);

                    uow.TacticalRowRepository.Update(Policy);
                    uow.Save();
                }

                DialogResult = DialogResult.OK;
                Close();
            }
            catch (Exception ex)
            {
                ShowError("Errore nel salvataggio: " + ex.Message);
            }
        }

        private void CopyValues(TacticalRow Policy)
        {
            var estinzioneParziale = Policy.CapitaleResiduo.HasValue;

            Policy.TransactionType = NuovoStatoPolizza.Text;
            if (estinzioneParziale)
            {
                Policy.DateOfCancellation3 = DataCancellazione.Value;
            }
            else
            {
                Policy.DateOfCancellation = DataCancellazione.Value;
            }
            Policy.DataLavorazione = DataLavorazione.Value;
            Policy.DataRimborso = DataRimborso.Value;
            Policy.RichiestaSottoscrittore = RichiestaSottoscrittore.Checked;
            Policy.CartaIdentita = CartaIdentita.Checked;
            Policy.CodiceFiscale = CodiceFiscale.Checked;
            Policy.Iban = Iban.Text;
            Policy.DebitoResiduo = HelperClass.GetDoubleNullableValue(DebitoResiduo);
            Policy.Cointestatario = Cointestatario.Checked;
            Policy.CointestatarioCartaIdentita = CointestatarioCartaIdentita.Checked;
            Policy.CognomeCoIntestario = CointestatrioCognome.Text;
            Policy.NomeCoIntestatario = CointestatarioNome.Text;
            Policy.CointestatarioDataNascita = CointestatarioDataNascita.Value;
        }

        private void AddBeneficiary_Click(object sender, EventArgs e)
        {
            if (BeneficiaryEditor.CreateNew(this, PolicyId) != null)
            {
                LoadBeneficiaries();
            }
        }

        private void ListaBeneficiari_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (ListaBeneficiari.SelectedRows.Count > 0)
            {
                var row = (BeneficiaryInfo)ListaBeneficiari.SelectedRows[0].DataBoundItem;
                if (BeneficiaryEditor.Edit(this, row.Id) != null)
                {
                    LoadBeneficiaries();
                }
            }
            else
            {
                ShowInfo("Selezionare un beneficiario da modificare");
            }

        }
    }
}
