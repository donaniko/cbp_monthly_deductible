using CBP.DataLayer;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data.Entity.Infrastructure.Interception;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CBP.Helpers
{
    public static class HelperClass
    {
        public static string TRANSATION_TYPE_CONTRATTO_CHIUSO = "TF";
        public static string TRANSATION_TYPE_RECESSO = "R";
        public static string TRANSATION_TYPE_DISDETTA = "D";
        public static string TRANSATION_TYPE_EA = "91";
        public static string TRANSATION_TYPE_REVOCA = "45";
        public static string TRANSATION_TYPE_NEWBUSINESS = "N";
        public static string TRANSATION_TYPE_VARIAZIONE = "V";
        public static string FormatNumber(double val, int decimals)
        {
            if (decimals > 0)
            {
                return val.ToString("#,##0." + "0".PadLeft(decimals, '0'));
            }
            else
            {
                return val.ToString("#,##0");
            }
        }

        public static void InitNumericTextBox(TextBox text, int decimals, bool nullValue)
        {
            text.Tag = new TextboxInfo { Decimals = decimals, NullValue = nullValue };
            text.Enter += NumericTextbox_Enter;
            text.Leave += NumericTextbox_Leave;
        }

        private static void NumericTextbox_Leave(object sender, EventArgs e)
        {
            var textbox = (TextBox)sender;
            var textboxInfo = (TextboxInfo)textbox.Tag;

            var value = 0D;
            if (!Double.TryParse(textbox.Text.Replace(".", ","), out value))
            {
                if (textboxInfo.NullValue)
                {
                    textbox.Text = "";
                }
                else
                {
                    textbox.Text = value.ToString();
                }
            }
            else
            {
                textbox.Text = FormatNumber(value, 2);
            }
        }

        public static BeneficiaryInfo GetBeneficiary(int id)
        {
            using (var uow = new UnitOfWork())
            {
                var entity = uow.BeneficiaryRepository.Get(b => b.Id == id).First();
                return new BeneficiaryInfo(entity, uow);
            }
        }

        public static BeneficiaryInfo PostBeneficiary(BeneficiaryInfo info)
        {            
            using (var uow = new UnitOfWork())
            {
                var sumQuota = 0D;
                var others = uow.BeneficiaryRepository.Get(b => b.TacticalRowId == info.TacticalRowId).ToList();
                foreach (var other in others)
                {
                    sumQuota += other.Share;
                }

                if (sumQuota + info.Share > 100)
                {
                    throw new Exception("La quota eccede 100%");
                }

                var entity = new Beneficiary();
                entity.BirthDate = info.BirthDate;
                entity.FirstName = info.FirstName;
                entity.Iban = info.Iban;
                entity.LastName = info.LastName;
                entity.Share = info.Share;
                entity.TacticalRowId = info.TacticalRowId;
                uow.BeneficiaryRepository.Insert(entity);
                uow.Save();

                return new BeneficiaryInfo(entity, uow);
            }
        }

        public static BeneficiaryInfo PutBeneficiary(BeneficiaryInfo info)
        {
            using (var uow = new UnitOfWork())
            {
                var sumQuota = 0D;
                var others = uow.BeneficiaryRepository.Get(b => b.Id != info.Id && b.TacticalRowId == info.TacticalRowId).ToList();
                foreach (var other in others)
                {
                    sumQuota += other.Share;
                }

                if (sumQuota + info.Share > 100)
                {
                    throw new Exception("La quota eccede 100%");
                }

                var entity = uow.BeneficiaryRepository.Get(b => b.Id == info.Id).First();
                entity.BirthDate = info.BirthDate;
                entity.FirstName = info.FirstName;
                entity.Iban = info.Iban;
                entity.LastName = info.LastName;
                entity.Share = info.Share;
                entity.TacticalRowId = info.TacticalRowId;
                uow.BeneficiaryRepository.Update(entity);
                uow.Save();

                return new BeneficiaryInfo(entity, uow);
            }
        }

        public static BeneficiaryInfo DeleteBeneficiary(int id)
        {
            using (var uow = new UnitOfWork())
            {
                var entity = uow.BeneficiaryRepository.Get(b => b.Id == id).First();
                uow.BeneficiaryRepository.Delete(entity);
                uow.Save();

                return new BeneficiaryInfo(entity, uow);
            }
        }

        public static DateTime GetDateNow()
        {
            return new DateTime(DateTime.Now.Year, DateTime.Now.Month, DateTime.Now.Day);
        }

        public static double? GetDoubleNullableValue(TextBox textbox)
        {
            var retVal = new double?();

            if (!string.IsNullOrEmpty(textbox.Text))
            {
                var value = 0D;
                if (Double.TryParse(textbox.Text, out value))
                {
                    retVal = value;
                }                
            }

            return retVal;
        }

        public static double GetDoubleValue(TextBox textbox)
        {
            var value = 0D;
            Double.TryParse(textbox.Text, out value);

            return value;
        }

        public static int GetIntValue(TextBox textbox)
        {
            var value = 0;
            int.TryParse(textbox.Text, out value);

            return value;
        }
        
        private static void NumericTextbox_Enter(object sender, EventArgs e)
        {
            var textbox = (TextBox)sender;

            var value = 0D;
            if (Double.TryParse(textbox.Text, out value))
            {
                textbox.Text = value.ToString().Replace(".", ",");
            }
            else
            {
                textbox.Text = "";
            }
        }

        public static void ShowError(Form owner, string err)
        {
            MessageBox.Show(owner, err, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static void ShowWarning(Form owner, string msg)
        {
            MessageBox.Show(owner, msg, "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        public static void ShowInfo(Form owner, string msg)
        {
            MessageBox.Show(owner, msg, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static DateTime? PromptDate(Form owner, string text, string title)
        {
            var retVal = new DateTime?();

            var inputDateText = HelperClass.PromptString(owner, text, title);

            if (string.IsNullOrEmpty(inputDateText))
            {
                ShowInfo(owner, "Nessuna variazione richiesta");
                return retVal;
            }

            var tokens = inputDateText.Split(new char[] { '/' });
            if (tokens.Length != 3)
            {
                ShowError(owner, "Data non in format dd/mm/yyyy");
                return retVal;
            }

            var day = 0;
            if (!int.TryParse(tokens[0], out day))
            {
                ShowError(owner, "Data non in format dd/mm/yyyy");
                return retVal;
            }

            var month = 0;
            if (!int.TryParse(tokens[1], out month))
            {
                ShowError(owner, "Data non in format dd/mm/yyyy");
                return retVal;
            }

            var year = 0;
            if (!int.TryParse(tokens[2], out year))
            {
                ShowError(owner, "Data non in format dd/mm/yyyy");
                return retVal;
            }

            return new DateTime(year, month, day, 0, 0, 0);
        }

        public static string PromptString(Form owner, string text, string caption)
        {
            Form prompt = new Form()
            {
                Width = 500,
                Height = 160,
                FormBorderStyle = FormBorderStyle.FixedDialog,
                Text = caption,
                StartPosition = FormStartPosition.CenterScreen
            };
            prompt.FormBorderStyle = FormBorderStyle.FixedToolWindow;
            Label textLabel = new Label() { Left = 50, Top = 20, Text = text, AutoSize = true };
            TextBox textBox = new TextBox() { Left = 50, Top = 50, Width = 400 };
            Button confirmation = new Button() { Text = "Ok", Left = 350, Width = 100, Top = 80, DialogResult = DialogResult.OK };
            confirmation.Click += (sender, e) => { prompt.Close(); };
            prompt.Controls.Add(textBox);
            prompt.Controls.Add(confirmation);
            prompt.Controls.Add(textLabel);
            prompt.AcceptButton = confirmation;

            return prompt.ShowDialog(owner) == DialogResult.OK ? textBox.Text : "";
        }

        public static void ImpostaDataValuta(int bankTransferId, DateTime dataValuta)
        {
            using (var uow = new UnitOfWork())
            {
                var bankTransfer = uow.BankTransferRepository.Get(bt => bt.Id == bankTransferId).First();

                bankTransfer.DataValuta = dataValuta;
                uow.BankTransferRepository.Update(bankTransfer);

                var cancellations = uow.CancellationRepository.Get(c => c.BankTransferId == bankTransfer.Id).ToList();
                foreach (var cancellation in cancellations)
                {
                    if (bankTransfer.RimborsoDanni)
                    {
                        cancellation.DataValutaDanni = dataValuta;
                    }
                    else
                    {
                        cancellation.DataValutaVita = dataValuta;
                    }                    

                    uow.CancellationRepository.Update(cancellation);
                }

                uow.Save();
            }
        }

        public static void ImpostaDataDistinta(int bankTransferId, DateTime dataDistinta)
        {
            using (var uow = new UnitOfWork())
            {
                var bankTransfer = uow.BankTransferRepository.Get(bt => bt.Id == bankTransferId).First();

                bankTransfer.DataDistinta = dataDistinta;
                uow.BankTransferRepository.Update(bankTransfer);

                var cancellations = uow.CancellationRepository.Get(c => c.BankTransferId == bankTransfer.Id).ToList();
                foreach (var cancellation in cancellations)
                {
                    if (bankTransfer.RimborsoDanni)
                    {
                        cancellation.DataDistintaDanni = dataDistinta;
                    }
                    else
                    {
                        cancellation.DataDistintaVita = dataDistinta;
                    }

                    uow.CancellationRepository.Update(cancellation);
                }

                uow.Save();
            }
        }

        public static DateTime GetDate(object value)
        {
            if (value is string || value is double)
            {
                var text = value.ToString();

                if (text.Length == 10)
                {
                    return new DateTime(int.Parse(text.Substring(6, 4)), int.Parse(text.Substring(3, 2)), int.Parse(text.Substring(0, 2)));
                }
                else if (text.Length == 8)
                {
                    return new DateTime(int.Parse(text.Substring(0, 4)), int.Parse(text.Substring(4, 2)), int.Parse(text.Substring(6, 2)));
                }
                else
                {
                    throw new Exception("Data non riconosciuta: " + text);
                }
            }
            else
            {
                return (DateTime)value;
            }
        }

        public static DateTime? GetDateNullable(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is ExcelErrorValue)
            {
                return null;
            }

            if (value is DateTime)
            {
                var d = (DateTime)value;
                if (d.Year == 1900)
                {
                    return null;
                }
            }

            if (value is string && value.ToString().Length != 10)
            {
                return null;
            }

            if (value is int)
            {
                return null;
            }

            if (value is double && value.ToString().Length != 8)
            {
                return null;
            }

            return GetDate(value);
        }

        public static void CreaDistinta(List<PaymentInfo> payments, bool rimborsoDanni, string nomeDistinta)
        {
            using (var uow = new UnitOfWork())
            {
                var distinta = uow.BankTransferRepository.Get(bt => bt.Name == nomeDistinta).FirstOrDefault();

                if (distinta == null)
                {
                    distinta = new BankTransfer();
                    distinta.Name = nomeDistinta;
                    distinta.DataDistinta = DateTime.Now;
                    distinta.RimborsoDanni = rimborsoDanni;
                    uow.BankTransferRepository.Insert(distinta);
                    uow.Save();
                }
                else
                {
                    var btPolicies = uow.BankTransferPolicyRepository.Get(btp => btp.BankTransferId == distinta.Id).ToList();

                    foreach (var btPolicy in btPolicies)
                    {
                        uow.BankTransferPolicyRepository.Delete(btPolicy);
                    }

                    distinta.DataDistinta = DateTime.Now;

                    uow.BankTransferRepository.Update(distinta);

                    uow.Save();
                }

                foreach (var payment in payments)
                {
                    var btPolicy = new BankTransferPolicy();
                    btPolicy.BankTransferId = distinta.Id;
                    btPolicy.TacticalRowId = payment.TacticalRowId;
                    uow.BankTransferPolicyRepository.Insert(btPolicy);

                    if (payment.DataCancellazione.HasValue)
                    {
                        var cancellation = uow.CancellationRepository.Get(c => c.TacticalRowId == payment.TacticalRowId && c.Date == payment.DataCancellazione.Value).FirstOrDefault();

                        if (cancellation == null)
                        {
                            cancellation = new Cancellation();
                            cancellation.TacticalRowId = payment.TacticalRowId;
                            cancellation.Date = payment.DataCancellazione.Value;
                        }
                        cancellation.BankTransferId = distinta.Id;

                        if (rimborsoDanni)
                        {
                            cancellation.DataAccontonatoDanni = null;
                            cancellation.DataDistintaDanni = distinta.DataDistinta;
                            cancellation.PremioAccantonatoDanni = payment.PagamentoAccantonato;
                            cancellation.PremioRimborsatoDanni = true;
                        }
                        else
                        {
                            cancellation.DataAccantonatoVita = null;
                            cancellation.DataDistintaVita = distinta.DataDistinta;
                            cancellation.PremioAccantonatoVita = payment.PagamentoAccantonato;
                            cancellation.PremioRimborsatoVita = true;
                        }

                        if (cancellation.Id == 0)
                        {
                            uow.CancellationRepository.Insert(cancellation);
                        }
                        else
                        {
                            uow.CancellationRepository.Update(cancellation);
                        }
                    }
                }

                uow.Save();
            }
        }
    }

    class TextboxInfo
    {
        public bool NullValue { get; set; }
        public int Decimals { get; set; }
    }
}
