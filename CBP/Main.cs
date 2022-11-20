using CBP.DataLayer;
using CBP.Helpers;
using iTextSharp.text;
using iTextSharp.text.pdf;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.Entity.Infrastructure;
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace CBP
{
    public partial class Main : Form
    {
        public delegate void ImportStep(string worksheet, int row);
        public delegate void ExportStep(string header, int currentStepIndex, int totalStep);
        public event ImportStep ImportStepAdvanced;
        public event ImportStep ImportStepAdvancedEAP;
        public event ImportStep ImportStepAdvancedRecessi;
        public event ExportStep ExportStepAdvanced;
        public string TIPOLOGIA_VITA = "VITA";
        public string TIPOLOGIA_DANNI = "DANNI";


        public Main()
        {
            InitializeComponent();

            Text = Program.TITOLO_APP + " - " + Text;

            OutputFolder.Text = ConfigurationManager.AppSettings["OutputFolder"] ?? "";
            if (string.IsNullOrEmpty(OutputFolder.Text))
            {
                OutputFolder.Text = @"C:\Develop\CBP\Dati\Output";
            }

            MonthList.SelectedIndex = DateTime.Now.Month - 1;
            ExportYear.Text = DateTime.Now.Year.ToString();
            ImportStepAdvanced += Main_ImportStepAdvanced;
            ImportStepAdvancedEAP += Main_ImportStepAdvancedEAP;
            ImportStepAdvancedRecessi += Main_ImportStepAdvancedRecessi;
            ExportStepAdvanced += Main_ExportStepAdvanced;

            try
            {
                using (var uow = new UnitOfWork())
                {
                    var generalInfo = uow.GeneralInfoRepository.Get().FirstOrDefault();
                    if (generalInfo == null)
                    {
                        generalInfo = new GeneralInfo();
                        uow.GeneralInfoRepository.Insert(generalInfo);
                        uow.Save();
                    }

                    LastImported.Text = generalInfo.LastImportedFileName ?? "";
                    LastImportedDate.Text = generalInfo.LastImportedDate.HasValue ? generalInfo.LastImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";

                    LastEAPImported.Text = generalInfo.LastEAPImportedFileName ?? "";
                    LastEAPImportedDate.Text = generalInfo.LastEAPImportedDate.HasValue ? generalInfo.LastEAPImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";

                    LastRecessiImported.Text = generalInfo.LastRecessiImportedFileName ?? "";
                    LastRecessiImportedDate.Text = generalInfo.LastRecessiImportedDate.HasValue ? generalInfo.LastRecessiImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                }
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        private void Main_ExportStepAdvanced(string header, int currentStepIndex, int totalStep)
        {
            if (totalStep == 0)
            {
                ExportProgress.Text = string.Format("{0} - Elaborazione riga {1}", header, currentStepIndex);
            }
            else
            {
                ExportProgress.Text = string.Format("{0} - Elaborazione riga {1} di {2}", header, currentStepIndex, totalStep);
            }
            Application.DoEvents();
        }

        private void Main_ImportStepAdvanced(string worksheet, int row)
        {
            CurrentWorksheetText.Text = worksheet;
            CurrentWorksheetRowText.Text = row.ToString();
            Application.DoEvents();
        }

        private void Main_ImportStepAdvancedEAP(string worksheet, int row)
        {
            CurrentEAPWorksheetText.Text = worksheet;
            CurrentEAPWorksheetRowText.Text = row.ToString();
            Application.DoEvents();
        }

        private void Main_ImportStepAdvancedRecessi(string worksheet, int row)
        {
            CurrentRecessiWorksheetText.Text = worksheet;
            CurrentRecessiWorksheetRowText.Text = row.ToString();
            Application.DoEvents();
        }

        private void ImportFileLavorativo_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file lavorativo";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                var countRows = StartImport(openFileDialog.FileName, true);
            }
        }

        private void ExportMonthly_Click(object sender, EventArgs e)
        {
            EsportaMonthly(false);
        }

        private string GetCurrentAmlName(string monthlyFile)
        {
            var suffix = "";
            monthlyFile = Path.GetFileNameWithoutExtension(monthlyFile);
            var index = monthlyFile.LastIndexOf("_");
            if (index > -1)
            {
                suffix = monthlyFile.Substring(index + 1);
            }
            return Path.Combine(OutputFolder.Text, "AmlInputFile_" + suffix + ".xlsx");
        }

        private string GetCurrentBankName(string monthlyFile)
        {
            var suffix = "";
            monthlyFile = Path.GetFileNameWithoutExtension(monthlyFile);
            var index = monthlyFile.LastIndexOf("_");
            if (index > -1)
            {
                suffix = monthlyFile.Substring(index + 1);
            }
            return Path.Combine(OutputFolder.Text, "BankInputFile_" + suffix + ".xlsx");
        }

        private string GetCurrentMonthlyName(bool recessiDisdette)
        {
            if (recessiDisdette)
            {
                return Path.Combine(OutputFolder.Text, "Monthly Cancellation File_RecessiDisdette.xlsx");
            }
            return Path.Combine(OutputFolder.Text, "Monthly Cancellation File_" + ExportYear.Text + (MonthList.SelectedIndex + 1).ToString().PadLeft(2, '0') + ".xlsx");
        }

        private string GetCurrentRecessiBancaName()
        {
            var startDate = new DateTime(int.Parse(ExportYear.Text), MonthList.SelectedIndex + 1, 1);
            var endDate = startDate.AddMonths(1).AddSeconds(-1);

            return Path.Combine(OutputFolder.Text, string.Format("Cancellations from {0} to {1}.xlsx", startDate.ToString("ddMMyyyy"), endDate.ToString("ddMMyyyy")));
        }

        private string GetCurrentExportName()
        {
            return Path.Combine(OutputFolder.Text, "Export_" + DateTime.Now.ToString("yyyyMMddhhmmss") + ".xlsx");
        }

        private void WriteLog(List<string> logRows, string name)
        {
            var logName = string.Format("{0}{1}.txt", name, DateTime.Now.ToString("yyyyMMdd_HHmmss"));
            using (var writer = new StreamWriter(Path.Combine(OutputFolder.Text, logName), false))
            {
                foreach (var logRow in logRows)
                {
                    writer.WriteLine(logRow);
                }
            }
        }

        private void EsportaEstinzioni(ExcelWorksheet ws, int month, int year, bool recessiDisdette, string outputFolder, ref int row, List<string> logRows)
        {
            var dtFrom = new DateTime(year, month, 1);

            var baseOutputFolder = Path.Combine(OutputFolder.Text, "Letters_" + MonthList.SelectedItem.ToString() + "_" + ExportYear.Text);

            if (!recessiDisdette)
            {
                if (Directory.Exists(baseOutputFolder))
                {
                    Directory.Delete(baseOutputFolder, true);
                }

                System.Threading.Thread.Sleep(1000);

                Directory.CreateDirectory(baseOutputFolder);
                FontFactory.Register(GetFilePath("Modellini\\calibri.ttf"));
                FontFactory.Register(GetFilePath("Modellini\\calibriBold.ttf"));
            }

            using (var uow = new UnitOfWork())
            {
                var products = uow.ProductRepository.Get().ToList();

                var query = uow.Context.Set<TacticalRow>().Where(tr => !tr.Santander);

                if (recessiDisdette)
                {
                    query = query.Where(tr => tr.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA
                                            && tr.RecessoDisdettaCompleto
                                            && !tr.Cancellations.Any(c => c.DataDistintaDanni.HasValue)
                                            && !tr.Cancellations.Any(c => c.DataDistintaVita.HasValue));
                }
                else
                {
                    query = query.Where(tr => tr.TransactionType == HelperClass.TRANSATION_TYPE_EA || tr.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA);

                    var dtTo = dtFrom.AddMonths(1).AddDays(-1);
                    dtTo = new DateTime(dtTo.Year, dtTo.Month, dtTo.Day, 23, 59, 59);

                    query = query.Where(tr => (tr.DateOfCancellation >= dtFrom && tr.DateOfCancellation <= dtTo)
                                            || (tr.DateOfCancellation2 >= dtFrom && tr.DateOfCancellation2 <= dtTo)
                                            || (tr.DateOfCancellation3 >= dtFrom && tr.DateOfCancellation3 <= dtTo));
                }

                var tacticalRows = query.OrderBy(tr => tr.PolicyNumber).ToList();

                //var tacticalRows = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == 14483997).ToList();

                int stepIndex = 1;

                foreach (var tacticalRow in tacticalRows)
                {
                    //if (tacticalRow.PolicyNumber == 15291296)
                    //{
                    //    var a = 0;
                    //}
                    RaiseExportStepAdvanced("Esportazione estinzioni", stepIndex++, tacticalRows.Count);

                    var product = products.Where(p => p.Code == tacticalRow.Tariff).FirstOrDefault();

                    if (product == null)
                    {
                        continue;
                    }

                    var premioDaRimborsare = GetPremioDaRimborsare(tacticalRow, product);
                    ws.Cells["L" + row.ToString()].Value = premioDaRimborsare;
                    ws.Cells["M" + row.ToString()].Value = GetCommissioniDaRimborsare(premioDaRimborsare, product);

                    CheckNegative(ws, row, logRows, tacticalRow.PolicyNumber.ToString());

                    ws.Cells["A" + row.ToString()].Value = product.Convenzione;
                    ws.Cells["B" + row.ToString()].Value = product.Prodotto;
                    ws.Cells["C" + row.ToString()].Value = tacticalRow.Surname;
                    ws.Cells["D" + row.ToString()].Value = tacticalRow.FirstName;
                    ws.Cells["E" + row.ToString()].Value = tacticalRow.Birth.ToString("dd/MM/yyyy");
                    ws.Cells["F" + row.ToString()].Value = tacticalRow.PolicyNumber;
                    ws.Cells["G" + row.ToString()].Value = tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_EA ? "Estinzione Anticipata": "Disdetta";
                    ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
                    ws.Cells["I" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                    ws.Cells["J" + row.ToString()].Value = product.FeesB;
                    ws.Cells["K" + row.ToString()].Value = product.FeesA;

                    var importoRimborso = GetDoubleNullable(ws.Cells["L" + row.ToString()].Value);

                    if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
                    {
                        ScriviValoriDisdetta(tacticalRow, importoRimborso, ws, row, true);
                    }
                    else
                    {
                        if (importoRimborso < 0)
                        {
                            ws.Cells["N" + row.ToString()].Value = "N";
                            ws.Cells["O" + row.ToString()].Value = "N";
                            ws.Cells["P" + row.ToString()].Value = "";
                            ws.Cells["Q" + row.ToString()].Value = "";
                        }
                        else
                        {
                            ws.Cells["N" + row.ToString()].Value = "S";
                            ws.Cells["O" + row.ToString()].Value = "N";
                            ws.Cells["P" + row.ToString()].Value = "";
                            ws.Cells["Q" + row.ToString()].Value = ws.Cells["H" + row.ToString()].Value;
                        }
                    }

                    ws.Cells["R" + row.ToString()].Value = tacticalRow.Tariff;

                    if (RimborsatoAccantonato.Checked)
                    {
                        ScriviValoriRimborsatoAccantonatoNelMonthly(ws, row, true, tacticalRow.Id, uow);
                    }

                    if (!recessiDisdette)
                    {
                        try
                        {
                            if (string.IsNullOrEmpty(tacticalRow.Iban) && importoRimborso > 0)
                            {
                                AddLogRow(logRows, "IBAN mancante", row, tacticalRow.PolicyNumber.ToString());
                                //CreaLetteraEstinzioniAnticipateNoIbanNuova(baseOutputFolder, tacticalRow, dtFrom);
                            }

                            if (tacticalRow.ExpirationDate.HasValue
                                && tacticalRow.DateOfCancellation.HasValue
                                && tacticalRow.ExpirationDate < tacticalRow.DateOfCancellation)
                            {
                                AddLogRow(logRows, "Data scadenza precedente alla data di estinzione", row, tacticalRow.PolicyNumber.ToString());
                            }

                            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
                            {
                                CreaLetteraDisdettaGAP(baseOutputFolder, tacticalRow, product, dtFrom);
                            }
                            else
                            {
                                if (importoRimborso <= 0)
                                {
                                    //CreaLetteraEstinzioniAnticipataZeroNuova(baseOutputFolder, tacticalRow, worksheetUltimoEstinzioni, dtFrom);
                                }
                                else
                                {
                                    CreaLetteraEstinzioniAnticipataNuova(baseOutputFolder, tacticalRow, product, dtFrom);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            AddLogRow(logRows, "Errore generazione lettera: " + ex.Message, row, tacticalRow.PolicyNumber.ToString());
                        }
                    }

                    row++;
                }
            }
        }

        private void GestioneModellinoEAP(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioni, TacticalRow tacticalRow, Product productVita)
        {
            worksheetEstinzioni.Cells["C9"].Value = productVita.CodeCNPSI;
            worksheetEstinzioni.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioni.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioni.Cells["C12"].Value = tacticalRow.DateOfCancellation;
            worksheetEstinzioni.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
            worksheetEstinzioni.Cells["C14"].Value = tacticalRow.DebitoResiduo;
            worksheetEstinzioni.Cells["C15"].Value = Math.Abs(tacticalRow.ImportoEstinzione2.HasValue ? tacticalRow.ImportoEstinzione2.Value : tacticalRow.ImportoEstinzione.Value);

            worksheetEstinzioni.Calculate();

            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D70"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D67"].Value);

            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E70"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E67"].Value);

            ws.Cells["G" + row.ToString()].Value = "Estinzione Parziale";
            ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Parziale";

            ws.Cells["J" + row.ToString()].Value = 0;
            ws.Cells["J" + (row + 1).ToString()].Value = 0;
            ws.Cells["K" + row.ToString()].Value = 0;
            ws.Cells["K" + (row + 1).ToString()].Value = 0;
        }

        private void GestioneModellinoEAP_ULTIMO(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioni, TacticalRow tacticalRow, Product productVita)
        {
            try
            {
                worksheetEstinzioni.Cells["C9"].Value = productVita.CodeCNPSI;
                worksheetEstinzioni.Cells["C10"].Value = tacticalRow.DateOfSignature;
                worksheetEstinzioni.Cells["C11"].Value = tacticalRow.ExpirationDate;
                worksheetEstinzioni.Cells["C12"].Value = tacticalRow.DateOfCancellation;
                worksheetEstinzioni.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
                worksheetEstinzioni.Cells["C14"].Value = tacticalRow.DebitoResiduo;
                worksheetEstinzioni.Cells["C15"].Value = 0;
                worksheetEstinzioni.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione2.HasValue ? tacticalRow.ImportoEstinzione2.Value : tacticalRow.ImportoEstinzione.Value);
                worksheetEstinzioni.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
                worksheetEstinzioni.Cells["C18"].Value = Math.Abs(tacticalRow.ImportoEstinzione2.HasValue ? tacticalRow.ImportoEstinzione2.Value : tacticalRow.ImportoEstinzione.Value);
                worksheetEstinzioni.Calculate();

                ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D74"].Value);
                ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D71"].Value);

                ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E74"].Value);
                ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E71"].Value);

                ws.Cells["G" + row.ToString()].Value = "Estinzione Parziale";
                ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Parziale";

                ws.Cells["J" + row.ToString()].Value = 0;
                ws.Cells["J" + (row + 1).ToString()].Value = 0;
                ws.Cells["K" + row.ToString()].Value = 0;
                ws.Cells["K" + (row + 1).ToString()].Value = 0;
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private void GestioneModellinoEA(ExcelWorksheet ws, int row, ExcelWorksheet worksheetEstinzioni, TacticalRow tacticalRow, Product productVita, Product productDanni)
        {
            //Estizione totale diretta
            worksheetEstinzioni.Cells["C9"].Value = productVita.CodeCNPSI;
            worksheetEstinzioni.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioni.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioni.Cells["C12"].Value = tacticalRow.DateOfCancellation;
            worksheetEstinzioni.Cells["C13"].Value = tacticalRow.InsuredSum - tacticalRow.PremioLordo;
            worksheetEstinzioni.Cells["C14"].Value = tacticalRow.DebitoResiduo;
            worksheetEstinzioni.Calculate();



            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D56"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["D46"].Value);


            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E56"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["E46"].Value);

            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
            {
                ws.Cells["G" + row.ToString()].Value = "Disdetta";
                ws.Cells["G" + (row + 1).ToString()].Value = "Disdetta";
            }
            else
            {
                ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
                ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Anticipata";
            }

            ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
            ws.Cells["H" + (row + 1).ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");

            ws.Cells["J" + row.ToString()].Value = productVita.FeesB;
            ws.Cells["J" + (row + 1).ToString()].Value = productDanni.FeesB;
            ws.Cells["K" + row.ToString()].Value = productVita.FeesA;
            ws.Cells["K" + (row + 1).ToString()].Value = productDanni.FeesA;
            //modellino.SaveAs(new FileInfo(tacticalRow.PolicyNumber.ToString() + ".xlsx"));
        }

        private int GetDurataGiorni(TacticalRow tacticalRow)
        {
            var retVal = 0;

            if (tacticalRow.ExpirationDate.HasValue && tacticalRow.DateOfSignature.HasValue)
            {
                retVal = (tacticalRow.ExpirationDate.Value - tacticalRow.DateOfSignature.Value).Days;
            }

            return retVal;
        }

        private int GetGiorniGoduti(TacticalRow tacticalRow)
        {
            var retVal = 0;

            if (tacticalRow.DateOfCancellation.HasValue && tacticalRow.DateOfSignature.HasValue)
            {
                retVal = (tacticalRow.DateOfCancellation.Value - tacticalRow.DateOfSignature.Value).Days;
            }

            return retVal;
        }       
        
        private double GetCommissioniDaRimborsare(double premioDaRimborsare, Product product)
        {
            var tassoCommissioni = 0.362887D;
            var commissioniDaRimborsare = Math.Round(premioDaRimborsare * tassoCommissioni, 2);

            return commissioniDaRimborsare;
        }

        private double GetPremioDaRimborsare(TacticalRow tacticalRow, Product product)
        {
            //Premium*((duration - [effective cancellation date - inception date])/duration)

            var durata = GetDurataGiorni(tacticalRow) * 1D;
            var giorniGoduti = GetGiorniGoduti(tacticalRow) * 1D;
            var premioDaRimborsare = Math.Round(tacticalRow.PremioLordo * ((durata - giorniGoduti) / durata), 2);

            return premioDaRimborsare;
        }

        private static double GetE29(Product product, double E17, double E18, double C13)
        {
            double E29;
            if (product.CodeCNPSI == "ITG03" || product.CodeCNPSI == "ITG04" || product.CodeCNPSI == "ITG09" || product.CodeCNPSI == "ITG10")
            {
                //SE(C13>25000;0;ARROTONDA((C13*ARROTONDA((E17*100%/(100%+E18));6));4));
                if (C13 > 25000)
                {
                    E29 = 0;
                }
                else
                {
                    //ARROTONDA((C13*ARROTONDA((E17*100%/(100%+E18));6));4)
                    E29 = C13 * (E17 / (1 + E18));
                }
            }
            else
            {
                //ARROTONDA((C13*ARROTONDA((E17*100%/(100%+E18));6));4))
                E29 = C13 * (E17 / (1 + E18));
            }

            return E29;
        }

        private void _GestioneModellinoEA_ULTIMO(ExcelWorksheet ws, int row, ExcelWorksheet worksheetEstinzioni, TacticalRow tacticalRow, Product product)
        {
            //Estizione totale diretta
            worksheetEstinzioni.Cells["C9"].Value = product.CodeCNPSI;
            worksheetEstinzioni.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioni.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioni.Cells["C12"].Value = tacticalRow.DateOfCancellation;
            worksheetEstinzioni.Cells["C13"].Value = tacticalRow.PremioLordo;
            //worksheetEstinzioni.Cells["C14"].Value = tacticalRow.DebitoResiduo;
            worksheetEstinzioni.Calculate();



            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["C30"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioni.Cells["C31"].Value);

            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
            {
                ws.Cells["G" + row.ToString()].Value = "Disdetta";
            }
            else
            {
                ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
            }

            ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");

            //ws.Cells["J" + row.ToString()].Value = productVita.FeesB;
            //ws.Cells["J" + (row + 1).ToString()].Value = productDanni.FeesB;
            //ws.Cells["K" + row.ToString()].Value = productVita.FeesA;
            //ws.Cells["K" + (row + 1).ToString()].Value = productDanni.FeesA;            
        }

        private void GestioneModellinoEAdopoEAP(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioniParziali, TacticalRow tacticalRow, Product productVita, Product productDanni)
        {
            worksheetEstinzioniParziali.Cells["C9"].Value = productVita.CodeCNPSI;
            worksheetEstinzioniParziali.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioniParziali.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioniParziali.Cells["C12"].Value = tacticalRow.DateOfCancellation3;
            worksheetEstinzioniParziali.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
            worksheetEstinzioniParziali.Cells["C14"].Value = (tacticalRow.CapitaleResiduo ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione.Value);
            worksheetEstinzioniParziali.Cells["C15"].Value = tacticalRow.DebitoResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione.Value);
            worksheetEstinzioniParziali.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C18"].Value = Math.Abs(tacticalRow.ImportoEstinzione ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Calculate();

            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D92"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D98"].Value);

            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E92"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E98"].Value);

            ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
            ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Anticipata";

            ws.Cells["J" + row.ToString()].Value = productVita.FeesB;
            ws.Cells["J" + (row + 1).ToString()].Value = productDanni.FeesB;
            ws.Cells["K" + row.ToString()].Value = productVita.FeesA;
            ws.Cells["K" + (row + 1).ToString()].Value = productDanni.FeesA;
        }

        private void GestioneModellinoEAdopoEAP_ULTIMO(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioniParziali, TacticalRow tacticalRow, Product product)
        {
            worksheetEstinzioniParziali.Cells["C9"].Value = product.CodeCNPSI;
            worksheetEstinzioniParziali.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioniParziali.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioniParziali.Cells["C12"].Value = tacticalRow.DateOfCancellation3;
            worksheetEstinzioniParziali.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
            worksheetEstinzioniParziali.Cells["C14"].Value = tacticalRow.DebitoResiduoIniziale ?? 0;
            worksheetEstinzioniParziali.Cells["C15"].Value = tacticalRow.DebitoResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione.Value);
            worksheetEstinzioniParziali.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C18"].Value = Math.Abs(tacticalRow.ImportoEstinzione ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Calculate();

            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D92"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D89"].Value);

            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E92"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E89"].Value);

            ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
            ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Anticipata";

            //ws.Cells["J" + row.ToString()].Value = productVita.FeesB;
            //ws.Cells["J" + (row + 1).ToString()].Value = productDanni.FeesB;
            //ws.Cells["K" + row.ToString()].Value = productVita.FeesA;
            //ws.Cells["K" + (row + 1).ToString()].Value = productDanni.FeesA;
        }

        private void GestioneModellinoEAPdopoEAP(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioniParziali, TacticalRow tacticalRow, Product productVita)
        {
            worksheetEstinzioniParziali.Cells["C9"].Value = productVita.CodeCNPSI;
            worksheetEstinzioniParziali.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioniParziali.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioniParziali.Cells["C12"].Value = tacticalRow.DateOfCancellation2;
            worksheetEstinzioniParziali.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
            worksheetEstinzioniParziali.Cells["C14"].Value = (tacticalRow.DebitoResiduo ?? 0);
            worksheetEstinzioniParziali.Cells["C15"].Value = 0;
            worksheetEstinzioniParziali.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C18"].Value = Math.Abs(tacticalRow.ImportoEstinzione ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Calculate();

            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D74"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D80"].Value);

            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E74"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E80"].Value);

            ws.Cells["G" + row.ToString()].Value = "Estinzione Parziale";
            ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Parziale";

            ws.Cells["J" + row.ToString()].Value = 0;
            ws.Cells["J" + (row + 1).ToString()].Value = 0;
            ws.Cells["K" + row.ToString()].Value = 0;
            ws.Cells["K" + (row + 1).ToString()].Value = 0;
        }

        private void GestioneModellinoEAPdopoEAP_ULTIMO(int row, ExcelWorksheet ws, ExcelWorksheet worksheetEstinzioniParziali, TacticalRow tacticalRow, Product productVita)
        {
            worksheetEstinzioniParziali.Cells["C9"].Value = productVita.CodeCNPSI;
            worksheetEstinzioniParziali.Cells["C10"].Value = tacticalRow.DateOfSignature;
            worksheetEstinzioniParziali.Cells["C11"].Value = tacticalRow.ExpirationDate;
            worksheetEstinzioniParziali.Cells["C12"].Value = tacticalRow.DateOfCancellation2;
            worksheetEstinzioniParziali.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
            worksheetEstinzioniParziali.Cells["C14"].Value = tacticalRow.DebitoResiduoIniziale ?? 0;
            worksheetEstinzioniParziali.Cells["C15"].Value = tacticalRow.DebitoResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
            worksheetEstinzioniParziali.Cells["C18"].Value = Math.Abs(tacticalRow.ImportoEstinzione ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione2 ?? 0);
            worksheetEstinzioniParziali.Calculate();

            ws.Cells["L" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D92"].Value);
            ws.Cells["M" + row.ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["D98"].Value);

            ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E92"].Value);
            ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(worksheetEstinzioniParziali.Cells["E98"].Value);

            ws.Cells["G" + row.ToString()].Value = "Estinzione Parziale";
            ws.Cells["G" + (row + 1).ToString()].Value = "Estinzione Parziale";

            ws.Cells["J" + row.ToString()].Value = 0;
            ws.Cells["J" + (row + 1).ToString()].Value = 0;
            ws.Cells["K" + row.ToString()].Value = 0;
            ws.Cells["K" + (row + 1).ToString()].Value = 0;
        }

        private void ScriviValoriDisdetta(TacticalRow tacticalRow, double? importoRimborso, ExcelWorksheet ws, int row, bool vita)
        {
            //var somma = importoRimborsoVita + importoRimborsoDanni;

            //if (somma >= 0)
            if (importoRimborso >= 0)
            {
                //MODIFICHE DA TESTARE 10/09/2021
                if (tacticalRow.DataRimborso.HasValue)
                //if (tacticalRow.DataRimborso.HasValue && tacticalRow.RecessoDisdettaCompleto)
                {
                    ws.Cells["N" + row.ToString()].Value = "N";
                    ws.Cells["O" + row.ToString()].Value = "S";
                }
                else
                {
                    ws.Cells["N" + row.ToString()].Value = "S";
                    ws.Cells["O" + row.ToString()].Value = "N";
                }
            }
            else
            {
                ws.Cells["N" + row.ToString()].Value = "N";
                ws.Cells["O" + row.ToString()].Value = "N";
            }

            ws.Cells["P" + row.ToString()].Value = tacticalRow.DateOfCancellation;
            ws.Cells["Q" + row.ToString()].Value = "";
        }

        private string GetLetterDate(DateTime d)
        {
            var nextMonth = d.AddMonths(1);
            //var date = new DateTime(nextMonth.Year, nextMonth.Month, 15);
            //if (date.DayOfWeek == DayOfWeek.Saturday)
            //{
            //    date = new DateTime(date.Year, date.Month, 17);
            //}
            //if (date.DayOfWeek == DayOfWeek.Sunday)
            //{
            //    date = new DateTime(date.Year, date.Month, 16);
            //}
            var date = new DateTime(nextMonth.Year, nextMonth.Month, 28);
            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = new DateTime(date.Year, date.Month, 27);
            }
            if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = new DateTime(date.Year, date.Month, 25);
            }

            return date.ToString("dd/MM/yyyy");
        }

        private string GetBankDate(DateTime d)
        {
            var nextMonth = d.AddMonths(1);
            var date = new DateTime(nextMonth.Year, nextMonth.Month, 15);
            if (date.DayOfWeek == DayOfWeek.Saturday)
            {
                date = new DateTime(date.Year, date.Month, 17);
            }
            if (date.DayOfWeek == DayOfWeek.Sunday)
            {
                date = new DateTime(date.Year, date.Month, 16);
            }
            return date.ToString("dd/MM/yyyy");
        }

        private void CheckNegative(ExcelWorksheet ws, int row, List<string> logRows, string policyNumber)
        {
            var value = ws.Cells["L" + row.ToString()].Value;

            if (value == null)
            {
                AddLogRow(logRows, "Importo non calcolato", row, policyNumber);
            }

            if (GetDoubleNullable(value) < 0)
            {
                AddLogRow(logRows, "Importo negativo", row, policyNumber);
            }
        }

        private void EsportaLeasing(ExcelWorksheet ws, int month, int year, string outputFolder, ref int row, List<string> logRows)
        {
            var dtFrom = new DateTime(year, month, 1);

            using (var modellino = new ExcelPackage(new FileInfo(GetFilePath("Modellini\\EstinzioniAnticipate.xlsx"))))
            {
                var modellinoLeasing = modellino.Workbook.Worksheets["Refund calculator (Leasing)-NEW"];

                using (var uow = new UnitOfWork())
                {
                    var products = uow.ProductRepository.Get().ToList();

                    List<TacticalRow> tacticalRows = null;

                    var dtTo = dtFrom.AddMonths(1).AddDays(-1);
                    dtTo = new DateTime(dtTo.Year, dtTo.Month, dtTo.Day, 23, 59, 59);

                    tacticalRows = uow.TacticalRowRepository.Get(tr => tr.TransactionType == "EA"
                                                                   && tr.DateOfCancellation >= dtFrom
                                                                   && tr.DateOfCancellation <= dtTo
                                                                   && !tr.Santander)
                                                            .OrderBy(tr => tr.PolicyNumber).ToList();

                    int stepIndex = 1;
                    foreach (var tacticalRow in tacticalRows)
                    {
                        RaiseExportStepAdvanced("Esportazione leasing", stepIndex++, tacticalRows.Count);

                        var productVita = products.Where(p => p.Code == tacticalRow.Tariff && p.Tipologia == TIPOLOGIA_VITA).FirstOrDefault();
                        var productDanni = products.Where(p => p.Code == tacticalRow.Tariff && p.Tipologia == TIPOLOGIA_DANNI).FirstOrDefault();

                        if (productVita == null || productDanni == null)
                        {
                            continue;
                        }

                        modellinoLeasing.Cells["C9"].Value = productVita.CodeCNPSI;
                        modellinoLeasing.Cells["C10"].Value = tacticalRow.DateOfSignature;
                        modellinoLeasing.Cells["C11"].Value = tacticalRow.ExpirationDate;
                        modellinoLeasing.Cells["C12"].Value = tacticalRow.DateOfCancellation;
                        modellinoLeasing.Cells["C13"].Value = tacticalRow.InsuredSum;
                        modellinoLeasing.Calculate();

                        ws.Cells["L" + row.ToString()].Value = WriteCurrency(modellinoLeasing.Cells["D47"].Value);
                        ws.Cells["M" + row.ToString()].Value = WriteCurrency(modellinoLeasing.Cells["D40"].Value);

                        ws.Cells["L" + (row + 1).ToString()].Value = WriteCurrency(modellinoLeasing.Cells["E47"].Value);
                        ws.Cells["M" + (row + 1).ToString()].Value = WriteCurrency(modellinoLeasing.Cells["E40"].Value);

                        ws.Cells["A" + row.ToString()].Value = productVita.Convenzione;
                        ws.Cells["B" + row.ToString()].Value = productVita.Prodotto;
                        ws.Cells["C" + row.ToString()].Value = tacticalRow.Surname;
                        ws.Cells["D" + row.ToString()].Value = tacticalRow.FirstName;
                        ws.Cells["E" + row.ToString()].Value = tacticalRow.Birth.ToString("dd/MM/yyyy");
                        ws.Cells["F" + row.ToString()].Value = tacticalRow.PolicyNumber;
                        ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
                        ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
                        ws.Cells["I" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                        ws.Cells["J" + row.ToString()].Value = productVita.FeesB;
                        ws.Cells["K" + row.ToString()].Value = productVita.FeesA;
                        ws.Cells["N" + row.ToString()].Value = "S";
                        ws.Cells["O" + row.ToString()].Value = "N";
                        ws.Cells["P" + row.ToString()].Value = "";
                        ws.Cells["Q" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
                        ws.Cells["R" + row.ToString()].Value = tacticalRow.Tariff;

                        row++;

                        ws.Cells["A" + row.ToString()].Value = productDanni.Convenzione;
                        ws.Cells["B" + row.ToString()].Value = productDanni.Prodotto;
                        ws.Cells["C" + row.ToString()].Value = tacticalRow.Surname;
                        ws.Cells["D" + row.ToString()].Value = tacticalRow.FirstName;
                        ws.Cells["E" + row.ToString()].Value = tacticalRow.Birth.ToString("dd/MM/yyyy");
                        ws.Cells["F" + row.ToString()].Value = tacticalRow.PolicyNumber;
                        ws.Cells["G" + row.ToString()].Value = "Estinzione Anticipata";
                        ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
                        ws.Cells["I" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                        ws.Cells["J" + row.ToString()].Value = productDanni.FeesB;
                        ws.Cells["K" + row.ToString()].Value = productDanni.FeesA;
                        ws.Cells["N" + row.ToString()].Value = "S";
                        ws.Cells["O" + row.ToString()].Value = "N";
                        ws.Cells["P" + row.ToString()].Value = "";
                        ws.Cells["Q" + row.ToString()].Value = tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy");
                        ws.Cells["R" + row.ToString()].Value = tacticalRow.Tariff;

                        row++;
                    }
                }
            }
        }

        private object WriteCurrency(object value)
        {
            if (value is double)
            {
                return Math.Round(Convert.ToDouble(value), 2, MidpointRounding.AwayFromZero);
            }

            return value;
        }

        private string WriteCurrencyString(object value)
        {
            var d = Convert.ToDouble(value);
            d = Math.Round(Convert.ToDouble(d), 2);

            return d.ToString("#,##0.00");
        }

        private string WriteCurrencyStringPercentual(object value)
        {
            var d = (double)value;
            d = Math.Round(Convert.ToDouble(d) * 100, 2);

            return d.ToString("#,##0.00") + " %";
        }

        private void EsportaRevoche(ExcelWorksheet ws, int month, int year, string outputFolder, ref int row, List<string> logRows)
        {
            var dtFrom = new DateTime(year, month, 1);

            using (var uow = new UnitOfWork())
            {
                var products = uow.ProductRepository.Get().ToList();

                List<TacticalRow> tacticalRows = null;

                var dtTo = dtFrom.AddMonths(1).AddDays(-1);
                dtTo = new DateTime(dtTo.Year, dtTo.Month, dtTo.Day, 23, 59, 59);

                tacticalRows = uow.TacticalRowRepository.Get(tr => (tr.TransactionType == HelperClass.TRANSATION_TYPE_REVOCA)
                                                               && tr.DateOfCancellation >= dtFrom
                                                               && tr.DateOfCancellation <= dtTo
                                                               && !tr.Santander)
                                                        .OrderBy(tr => tr.PolicyNumber).ToList();

                int stepIndex = 1;
                foreach (var tacticalRow in tacticalRows)
                {
                    RaiseExportStepAdvanced("Esportazione revoche", stepIndex++, tacticalRows.Count);

                    var product = products.Where(p => p.Code == tacticalRow.Tariff).FirstOrDefault();

                    if (product == null)
                    {
                        continue;
                    }

                    ws.Cells["A" + row.ToString()].Value = product.Convenzione;
                    ws.Cells["B" + row.ToString()].Value = product.Prodotto;
                    ws.Cells["C" + row.ToString()].Value = tacticalRow.Surname;
                    ws.Cells["D" + row.ToString()].Value = tacticalRow.FirstName;
                    ws.Cells["E" + row.ToString()].Value = tacticalRow.Birth.ToString("dd/MM/yyyy");
                    ws.Cells["F" + row.ToString()].Value = tacticalRow.PolicyNumber;
                    ws.Cells["G" + row.ToString()].Value = "Revoca";
                    ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                    ws.Cells["I" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                    ws.Cells["J" + row.ToString()].Value = 0;
                    ws.Cells["K" + row.ToString()].Value = 0;
                    ws.Cells["L" + row.ToString()].Value = tacticalRow.PremioLordo;
                    ws.Cells["M" + row.ToString()].Value = tacticalRow.Commissioni;
                    ws.Cells["N" + row.ToString()].Value = "N";
                    ws.Cells["O" + row.ToString()].Value = "N";
                    ws.Cells["P" + row.ToString()].Value = "";
                    ws.Cells["Q" + row.ToString()].Value = "";
                    ws.Cells["R" + row.ToString()].Value = tacticalRow.Tariff;

                    row++;
                }
            }
        }

        private void EsportaRecessi(ExcelWorksheet ws, int month, int year, bool recessiDisdette, string outputFolder, ref int row, List<string> logRows)
        {
            var dtFrom = new DateTime(year, month, 1);

            var baseOutputFolder = Path.Combine(OutputFolder.Text, "Letters_" + MonthList.SelectedItem.ToString() + "_" + ExportYear.Text);

            using (var uow = new UnitOfWork())
            {
                var products = uow.ProductRepository.Get().ToList();

                var query = uow.Context.Set<TacticalRow>().Where(tr => !tr.Santander);

                if (recessiDisdette)
                {
                    query = query.Where(tr => (tr.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO)
                                            && tr.RecessoDisdettaCompleto
                                            && !tr.Cancellations.Any(c => c.DataDistintaDanni.HasValue)
                                            && !tr.Cancellations.Any(c => c.DataDistintaVita.HasValue));
                }
                else
                {
                    var dtTo = dtFrom.AddMonths(1).AddDays(-1);
                    dtTo = new DateTime(dtTo.Year, dtTo.Month, dtTo.Day, 23, 59, 59);

                    query = query.Where(tr => (tr.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO)
                                            && tr.DateOfCancellation >= dtFrom
                                            && tr.DateOfCancellation <= dtTo);
                }



                var tacticalRows = query.OrderBy(tr => tr.PolicyNumber).ToList();

                int stepIndex = 1;
                foreach (var tacticalRow in tacticalRows)
                {
                    RaiseExportStepAdvanced("Esportazione recessi", stepIndex++, tacticalRows.Count);

                    var product = products.Where(p => p.Code == tacticalRow.Tariff).FirstOrDefault();

                    if (product == null)
                    {
                        continue;
                    }

                    ws.Cells["A" + row.ToString()].Value = product.Convenzione;
                    ws.Cells["B" + row.ToString()].Value = product.Prodotto;
                    ws.Cells["C" + row.ToString()].Value = tacticalRow.Surname;
                    ws.Cells["D" + row.ToString()].Value = tacticalRow.FirstName;
                    ws.Cells["E" + row.ToString()].Value = tacticalRow.Birth.ToString("dd/MM/yyyy");
                    ws.Cells["F" + row.ToString()].Value = tacticalRow.PolicyNumber;
                    ws.Cells["G" + row.ToString()].Value = "Recesso";
                    ws.Cells["H" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                    ws.Cells["I" + row.ToString()].Value = tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy");
                    ws.Cells["J" + row.ToString()].Value = 0;
                    ws.Cells["K" + row.ToString()].Value = 0;
                    ws.Cells["L" + row.ToString()].Value = tacticalRow.PremioNetto;
                    ws.Cells["M" + row.ToString()].Value = tacticalRow.Commissioni;

                    if (tacticalRow.DataRimborso.HasValue)
                    {
                        ws.Cells["N" + row.ToString()].Value = "N";
                        ws.Cells["O" + row.ToString()].Value = "S";
                        ws.Cells["P" + row.ToString()].Value = tacticalRow.DateOfCancellation;
                        ws.Cells["Q" + row.ToString()].Value = "";
                    }
                    else
                    {
                        ws.Cells["N" + row.ToString()].Value = "S";
                        ws.Cells["O" + row.ToString()].Value = "N";
                        ws.Cells["P" + row.ToString()].Value = tacticalRow.DateOfCancellation;
                        ws.Cells["Q" + row.ToString()].Value = "";
                    }
                    ws.Cells["R" + row.ToString()].Value = tacticalRow.Tariff;

                    CreaLetteraRecesso(baseOutputFolder, tacticalRow, dtFrom);

                    row++;
                }
            }
        }

        private void AddLogRow(List<string> logRows, string text, int row, string policyNumber)
        {
            if (!logRows.Contains(text))
            {
                logRows.Add(string.Format("Riga {0}. Pratica {1}. Messaggio: {2}", row, policyNumber, text));
            }
        }

        private string GetFilePath(string filePath)
        {
            var exePath = new FileInfo(Application.ExecutablePath);
            var retVal = Path.Combine(exePath.DirectoryName, filePath);

            return retVal;
        }

        private int StartImportLavorativo(string fileName, bool endMessage)
        {
            int retVal = 0;
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    if (excelPackage.Workbook.Worksheets["Tactical CL"] == null)
                    {
                        ShowError("Foglio 'Tactical CL' non trovato");
                        return retVal;
                    }

                    if (excelPackage.Workbook.Worksheets["Tactical Leasing"] == null)
                    {
                        ShowError("Foglio 'Tactical Leasing' non trovato");
                        return retVal;
                    }

                    Cursor = Cursors.WaitCursor;
                    CurrentWorksheetText.Text = "";
                    CurrentWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();
                    var count = ImportTactical(excelPackage.Workbook.Worksheets["Tactical CL"], true, logRows, false);
                    count += ImportTactical(excelPackage.Workbook.Worksheets["Tactical Leasing"], false, logRows, true);

                    WriteLog(logRows, "ImportLog");

                    using (var uow = new UnitOfWork())
                    {
                        var generalInfo = uow.GeneralInfoRepository.Get().First();
                        generalInfo.LastImportedDate = DateTime.Now;
                        generalInfo.LastImportedFileName = Path.GetFileName(fileName);
                        uow.GeneralInfoRepository.Update(generalInfo);
                        uow.Save();

                        LastImported.Text = generalInfo.LastImportedFileName ?? "";
                        LastImportedDate.Text = generalInfo.LastImportedDate.HasValue ? generalInfo.LastImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                    }

                    retVal = count;

                    if (endMessage)
                    {
                        Cursor = Cursors.Default;
                        ShowInfo("Righe importate: " + count.ToString());

                        Process.Start(OutputFolder.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }

            return retVal;
        }

        private int StartImport(string fileName, bool endMessage)
        {
            int retVal = 0;
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    Cursor = Cursors.WaitCursor;
                    CurrentWorksheetText.Text = "";
                    CurrentWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();

                    var wsTactical1 = excelPackage.Workbook.Worksheets[Program.TACTICAL_SHEET_1];
                    var wsProduction1 = excelPackage.Workbook.Worksheets[Program.TACTICAL_SHEET_2];

                    if (wsTactical1 == null)
                    {
                        throw new Exception("Sheet 1 non trovato: " + Program.TACTICAL_SHEET_1);
                    }

                    if (wsProduction1 == null)
                    {
                        throw new Exception("Sheet 2 non trovato: " + Program.TACTICAL_SHEET_2);
                    }

                    var count = ImportaTactical(wsTactical1, 
                                                wsProduction1, 
                                                logRows);

                    WriteLog(logRows, "ImportLog");

                    using (var uow = new UnitOfWork())
                    {
                        var generalInfo = uow.GeneralInfoRepository.Get().First();
                        generalInfo.LastImportedDate = DateTime.Now;
                        generalInfo.LastImportedFileName = Path.GetFileName(fileName);
                        uow.GeneralInfoRepository.Update(generalInfo);
                        uow.Save();

                        LastImported.Text = generalInfo.LastImportedFileName ?? "";
                        LastImportedDate.Text = generalInfo.LastImportedDate.HasValue ? generalInfo.LastImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                    }

                    retVal = count;

                    if (endMessage)
                    {
                        Cursor = Cursors.Default;
                        ShowInfo("Righe importate: " + count.ToString());

                        Process.Start(OutputFolder.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }

            return retVal;
        }

        private void StartImportScb(string fileName, bool endMessage, int previousRows)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    if (excelPackage.Workbook.Worksheets["SCB-CL"] == null
                        && excelPackage.Workbook.Worksheets["HCBE-CL"] == null
                        && excelPackage.Workbook.Worksheets["TIM FIN-CL"] == null)
                    {
                        ShowError("Foglio 'SCB-CL' o 'HCBE-CL' o 'TIM FIN-CL' non trovato");
                        return;
                    }

                    if (excelPackage.Workbook.Worksheets["SCB-Leasing"] == null
                        && excelPackage.Workbook.Worksheets["HCBE-Leasing"] == null
                        && excelPackage.Workbook.Worksheets["TIM FIN-Leasing"] == null)
                    {
                        ShowError("Foglio 'SCB-Leasing' o 'HCBE-Leasing' o 'TIM FIN-Leasing' non trovato");
                        return;
                    }

                    Cursor = Cursors.WaitCursor;
                    CurrentWorksheetText.Text = "";
                    CurrentWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();

                    var clSheet = excelPackage.Workbook.Worksheets["SCB-CL"];
                    if (clSheet == null)
                    {
                        clSheet = excelPackage.Workbook.Worksheets["HCBE-CL"];
                    }
                    if (clSheet == null)
                    {
                        clSheet = excelPackage.Workbook.Worksheets["TIM FIN-CL"];
                    }

                    var leasingSheet = excelPackage.Workbook.Worksheets["SCB-Leasing"];
                    if (leasingSheet == null)
                    {
                        leasingSheet = excelPackage.Workbook.Worksheets["HCBE-Leasing"];
                    }
                    if (leasingSheet == null)
                    {
                        leasingSheet = excelPackage.Workbook.Worksheets["TIM FIN-Leasing"];
                    }

                    var count = ImportScbCl(clSheet, logRows);
                    count += ImportScbLeasing(leasingSheet, logRows);

                    WriteLog(logRows, "ImportLog");

                    using (var uow = new UnitOfWork())
                    {
                        var generalInfo = uow.GeneralInfoRepository.Get().First();
                        generalInfo.LastImportedDate = DateTime.Now;
                        generalInfo.LastImportedFileName = Path.GetFileName(fileName);
                        uow.GeneralInfoRepository.Update(generalInfo);
                        uow.Save();

                        LastImported.Text = generalInfo.LastImportedFileName ?? "";
                        LastImportedDate.Text = generalInfo.LastImportedDate.HasValue ? generalInfo.LastImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                    }

                    if (endMessage)
                    {
                        Cursor = Cursors.Default;
                        ShowInfo("Righe importate: " + (previousRows + count).ToString());

                        Process.Start(OutputFolder.Text);
                    }
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void StartImportDatiAggiuntivi(string fileName)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    if (excelPackage.Workbook.Worksheets.Count == 0)
                    {
                        ShowError("Foglio excel non valido");
                        return;
                    }

                    Cursor = Cursors.WaitCursor;
                    CurrentWorksheetText.Text = "";
                    CurrentWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();
                    var count = ImportaDatiAggiuntivi(excelPackage.Workbook.Worksheets[1], logRows);

                    WriteLog(logRows, "ImportLog");

                    Cursor = Cursors.Default;
                    ShowInfo("Righe importate: " + count.ToString());
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void StartImportCodiciHE(string fileName)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    if (excelPackage.Workbook.Worksheets.Count == 0)
                    {
                        ShowError("Foglio excel non valido");
                        return;
                    }

                    Cursor = Cursors.WaitCursor;
                    CurrentWorksheetText.Text = "";
                    CurrentWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();
                    var count = ImportaCodiciHE(excelPackage.Workbook.Worksheets[1], logRows);

                    WriteLog(logRows, "ImportLog");

                    Cursor = Cursors.Default;
                    ShowInfo("Righe importate: " + count.ToString());
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void StartImportEAP(string fileName)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    Cursor = Cursors.WaitCursor;
                    CurrentEAPWorksheetText.Text = "";
                    CurrentEAPWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();
                    var count = ImportEAPWorksheet(excelPackage.Workbook.Worksheets[1], logRows);

                    WriteLog(logRows, "ImportLogEAP");

                    using (var uow = new UnitOfWork())
                    {
                        var generalInfo = uow.GeneralInfoRepository.Get().First();
                        generalInfo.LastEAPImportedDate = DateTime.Now;
                        generalInfo.LastEAPImportedFileName = Path.GetFileName(fileName);
                        uow.GeneralInfoRepository.Update(generalInfo);
                        uow.Save();

                        LastEAPImported.Text = generalInfo.LastEAPImportedFileName ?? "";
                        LastEAPImportedDate.Text = generalInfo.LastEAPImportedDate.HasValue ? generalInfo.LastEAPImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                    }

                    Cursor = Cursors.Default;
                    ShowInfo("Righe importate: " + count.ToString());

                    Process.Start(OutputFolder.Text);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void StartImportRecessi(string fileName)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    Cursor = Cursors.WaitCursor;
                    CurrentEAPWorksheetText.Text = "";
                    CurrentEAPWorksheetRowText.Text = "";
                    Application.DoEvents();

                    var logRows = new List<string>();
                    var count = ImportRecessiWorksheet(excelPackage.Workbook.Worksheets[1], logRows);

                    WriteLog(logRows, "ImportLogRecessi");

                    using (var uow = new UnitOfWork())
                    {
                        var generalInfo = uow.GeneralInfoRepository.Get().First();
                        generalInfo.LastRecessiImportedDate = DateTime.Now;
                        generalInfo.LastRecessiImportedFileName = Path.GetFileName(fileName);
                        uow.GeneralInfoRepository.Update(generalInfo);
                        uow.Save();

                        LastRecessiImported.Text = generalInfo.LastRecessiImportedFileName ?? "";
                        LastRecessiImportedDate.Text = generalInfo.LastRecessiImportedDate.HasValue ? generalInfo.LastRecessiImportedDate.Value.ToString("dd/MM/yyyy HH:mm") : "";
                    }

                    Cursor = Cursors.Default;
                    ShowInfo("Righe importate: " + count.ToString());

                    Process.Start(OutputFolder.Text);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void StartImportProducts(string fileName)
        {
            try
            {
                using (var excelPackage = new ExcelPackage(new FileInfo(fileName)))
                {
                    Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    var ws = excelPackage.Workbook.Worksheets[1];
                    var row = 2;
                    var exit = false;

                    using (var uow = new UnitOfWork())
                    {
                        while (!exit)
                        {
                            try
                            {
                                if (ws.Cells[row, 2].Value == null)
                                {
                                    exit = true;
                                }
                                else
                                {
                                    var tipologia = GetString(ws.Cells[row, 2].Value);
                                    var productCodeSCB = GetString(ws.Cells[row, 3].Value);
                                    var productCodeCNPSI = GetString(ws.Cells[row, 4].Value);
                                    var productName = GetString(ws.Cells[row, 5].Value);
                                    var convenzione = GetString(ws.Cells[row, 6].Value);
                                    var feesA = GetDoubleNullable(ws.Cells[row, 7].Value);
                                    var feesB = GetDoubleNullable(ws.Cells[row, 8].Value);

                                    var product = uow.ProductRepository.Get(p => p.Code == productCodeSCB && p.Tipologia == tipologia).FirstOrDefault();

                                    if (product == null)
                                    {
                                        product = new Product();
                                        product.Code = productCodeSCB;
                                        product.Tipologia = tipologia;
                                    }

                                    product.CodeCNPSI = productCodeCNPSI;
                                    product.Prodotto = productName;
                                    product.Convenzione = convenzione;
                                    product.FeesA = feesA;
                                    product.FeesB = feesB;

                                    if (product.Id == 0)
                                    {
                                        uow.ProductRepository.Insert(product);
                                    }
                                    else
                                    {
                                        uow.ProductRepository.Update(product);
                                    }
                                    uow.Save();

                                    row++;
                                }
                            }
                            catch (Exception ex)
                            {
                                throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                            }
                        }
                    }

                    Cursor = Cursors.Default;
                    ShowInfo("Prodotti importati");
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void ShowWarning(string msg)
        {
            MessageBox.Show(this, msg, "Attenzione", MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
        }

        private void ShowInfo(string msg)
        {
            MessageBox.Show(this, msg, "Info", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowError(string err)
        {
            MessageBox.Show(this, err, "Errore", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        private int ImportScbCl(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells["J" + row.ToString()].Value);

                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                AddLogRow(logRows, "Polizza non ancora caricata", row, policyNumber.ToString());
                            }
                            else
                            {
                                var province = GetString(ws.Cells["AH" + row.ToString()].Value);
                                if (!string.IsNullOrEmpty(province))
                                {
                                    tacticalRow.Provincia = province;
                                }

                                var iban = GetString(ws.Cells["AT" + row.ToString()].Value);
                                if (!string.IsNullOrEmpty(iban))
                                {
                                    tacticalRow.Iban = iban;
                                }

                                var nominativoAderente = GetString(ws.Cells["AE" + row.ToString()].Value);
                                if (!string.IsNullOrEmpty(nominativoAderente))
                                {
                                    if (tacticalRow.Surname + " " + tacticalRow.FirstName != nominativoAderente)
                                    {
                                        //Aderente diverso da contraente
                                        tacticalRow.NominativoAderente = nominativoAderente;
                                        var codiceFiscaleAderente = GetString(ws.Cells["AF" + row.ToString()].Value);
                                        if (!string.IsNullOrEmpty(codiceFiscaleAderente))
                                        {
                                            tacticalRow.CodiceFiscaleAderente = codiceFiscaleAderente;
                                        }
                                    }
                                }

                                uow.TacticalRowRepository.Update(tacticalRow);

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        /// <summary>
        /// Get Date From Fiscal Code
        /// </summary>
        /// <param name=”fiscalCode”></param>
        /// <returns>dd/MM/yyyy otherwise string.Empty</returns>
        private DateTime GetDateFromFiscalCode(string fiscalCode)
        {
            try
            {
                Dictionary<string, int> month = new Dictionary<string, int>();
                // To Upper
                fiscalCode = fiscalCode.ToUpper();
                month.Add("A", 1);
                month.Add("B", 2);
                month.Add("C", 3);
                month.Add("D", 4);
                month.Add("E", 5);
                month.Add("H", 6);
                month.Add("L", 7);
                month.Add("M", 8);
                month.Add("P", 9);
                month.Add("R", 10);
                month.Add("S", 11);
                month.Add("T", 12);
                // Get Date
                string date = fiscalCode.Substring(6, 5);
                int y = int.Parse(date.Substring(0, 2));
                string yy = ((y < 9) ? "20" : "19") + y.ToString("00");
                int m = month[date.Substring(2, 1)];
                int d = int.Parse(date.Substring(3, 2));
                if (d > 31)
                {
                    d -= 40;
                }

                return new DateTime(int.Parse(yy), m, d);
            }
            catch
            {
                throw new Exception("errore calcolo data di nascita per " + fiscalCode);
            }
        }

        private int ImportScbLeasing(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells["J" + row.ToString()].Value);

                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                AddLogRow(logRows, "Polizza non ancora caricata", row, policyNumber.ToString());
                            }
                            else
                            {
                                var province = GetString(ws.Cells["AP" + row.ToString()].Value);
                                if (!string.IsNullOrEmpty(province))
                                {
                                    tacticalRow.Provincia = province;
                                }

                                //tacticalRow.Iban = GetString(ws.Cells[row, 46].Value);

                                uow.TacticalRowRepository.Update(tacticalRow);

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        /// <summary>
        /// Colonna 1: PolicyNumber
        /// Colonna 2: IBAN
        /// Colonna 3: Provincia
        /// Colonna 4: DateOfCancellation
        /// Colonna 5: DateOfCancellation2
        /// Colonna 6: DateOfCancellation3
        /// Colonna 7: Capitale residuo
        /// Colonna 8: Importo Estinzione
        /// Colonna 9: Importo Estinzione 2
        /// Colonna 10: TransactionType
        /// </summary>
        /// <param name="ws"></param>
        /// <param name="logRows"></param>
        /// <returns></returns>
        private int ImportaDatiAggiuntivi(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells[row, 1].Value);

                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                AddLogRow(logRows, "Polizza non ancora caricata", row, policyNumber.ToString());
                            }
                            else
                            {
                                var iban = GetString(ws.Cells[row, 2].Value);
                                var province = GetString(ws.Cells[row, 3].Value);
                                var dateOfCancellation = GetString(ws.Cells[row, 4].Value);
                                var dateOfCancellation2 = GetString(ws.Cells[row, 5].Value);
                                var dateOfCancellation3 = GetString(ws.Cells[row, 6].Value);
                                var capitaleResiduo = GetString(ws.Cells[row, 7].Value);
                                var importoEstinzione = GetString(ws.Cells[row, 8].Value);
                                var importoEstinzione2 = GetString(ws.Cells[row, 9].Value);
                                var transactionType = GetString(ws.Cells[row, 10].Value);
                                var debitoResiduo = GetString(ws.Cells[row, 11].Value);
                                var cognomeCoIntestario = GetString(ws.Cells[row, 12].Value);
                                var nomeCoIntestario = GetString(ws.Cells[row, 13].Value);
                                var debitoResiduoPrimaEap = GetString(ws.Cells[row, 14].Value);

                                if (!string.IsNullOrEmpty(iban))
                                {
                                    tacticalRow.Iban = iban;
                                }

                                if (!string.IsNullOrEmpty(province))
                                {
                                    tacticalRow.Provincia = province;
                                }

                                if (!string.IsNullOrEmpty(dateOfCancellation))
                                {
                                    if (dateOfCancellation.ToUpper() == "NULL")
                                    {
                                        tacticalRow.DateOfCancellation = null;
                                    }
                                    else
                                    {
                                        tacticalRow.DateOfCancellation = GetDate(ws.Cells[row, 4].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(dateOfCancellation2))
                                {
                                    if (dateOfCancellation2.ToUpper() == "NULL")
                                    {
                                        tacticalRow.DateOfCancellation2 = null;
                                    }
                                    else
                                    {
                                        tacticalRow.DateOfCancellation2 = GetDate(ws.Cells[row, 5].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(dateOfCancellation3))
                                {
                                    if (dateOfCancellation3.ToUpper() == "NULL")
                                    {
                                        tacticalRow.DateOfCancellation3 = null;
                                    }
                                    else
                                    {
                                        tacticalRow.DateOfCancellation3 = GetDate(ws.Cells[row, 6].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(capitaleResiduo))
                                {
                                    if (capitaleResiduo.ToUpper() == "NULL")
                                    {
                                        tacticalRow.CapitaleResiduo = null;
                                    }
                                    else
                                    {
                                        tacticalRow.CapitaleResiduo = GetDoubleNullable(ws.Cells[row, 7].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(importoEstinzione))
                                {
                                    if (importoEstinzione.ToUpper() == "NULL")
                                    {
                                        tacticalRow.ImportoEstinzione = null;
                                    }
                                    else
                                    {
                                        tacticalRow.ImportoEstinzione = GetDoubleNullable(ws.Cells[row, 8].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(importoEstinzione2))
                                {
                                    if (importoEstinzione2.ToUpper() == "NULL")
                                    {
                                        tacticalRow.ImportoEstinzione2 = null;
                                    }
                                    else
                                    {
                                        tacticalRow.ImportoEstinzione2 = GetDoubleNullable(ws.Cells[row, 9].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(transactionType))
                                {
                                    if (transactionType.ToUpper() == "NULL")
                                    {
                                        tacticalRow.TransactionType = null;
                                    }
                                    else
                                    {
                                        tacticalRow.TransactionType = GetString(ws.Cells[row, 10].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(debitoResiduo))
                                {
                                    if (debitoResiduo.ToUpper() == "NULL")
                                    {
                                        tacticalRow.DebitoResiduo = null;
                                    }
                                    else
                                    {
                                        tacticalRow.DebitoResiduo = GetDoubleNullable(ws.Cells[row, 11].Value);
                                    }
                                }

                                if (!string.IsNullOrEmpty(cognomeCoIntestario))
                                {
                                    if (cognomeCoIntestario.ToUpper() == "NULL")
                                    {
                                        tacticalRow.CognomeCoIntestario = null;
                                    }
                                    else
                                    {
                                        tacticalRow.CognomeCoIntestario = cognomeCoIntestario;
                                    }
                                }

                                if (!string.IsNullOrEmpty(nomeCoIntestario))
                                {
                                    if (nomeCoIntestario.ToUpper() == "NULL")
                                    {
                                        tacticalRow.NomeCoIntestatario = null;
                                    }
                                    else
                                    {
                                        tacticalRow.NomeCoIntestatario = nomeCoIntestario;
                                    }
                                }

                                if (!string.IsNullOrEmpty(debitoResiduoPrimaEap))
                                {
                                    if (debitoResiduoPrimaEap.ToUpper() == "NULL")
                                    {
                                        tacticalRow.DebitoResiduoIniziale = null;
                                    }
                                    else
                                    {
                                        tacticalRow.DebitoResiduoIniziale = GetDoubleNullable(ws.Cells[row, 14].Value);
                                    }
                                }

                                uow.TacticalRowRepository.Update(tacticalRow);

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        private int ImportaCodiciHE(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells[row, 1].Value);

                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                AddLogRow(logRows, "Polizza non ancora caricata", row, policyNumber.ToString());
                            }
                            else
                            {
                                var codiceHE = GetString(ws.Cells[row, 2].Value);

                                if (!string.IsNullOrEmpty(codiceHE))
                                {
                                    tacticalRow.HE = codiceHE;
                                }
                                else
                                {
                                    AddLogRow(logRows, "Codice HE non trovato nella colonna giusta", row, policyNumber.ToString());
                                }

                                uow.TacticalRowRepository.Update(tacticalRow);

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        private bool IsRecessoDisdetta(TacticalRow tacticalRow)
        {
            return tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO || tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA;
        }

        private int ImportaTactical(ExcelWorksheet ws, ExcelWorksheet wsProd, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                var products = uow.ProductRepository.Get().ToList();

                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells["A" + row.ToString()].Value);
                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                tacticalRow = new TacticalRow();
                                tacticalRow.PolicyNumber = policyNumber;
                            }

                            var nonModificareStato = false;
                            var importTransactionType = GetString(ws.Cells["I" + row.ToString()].Value);
                            if (tacticalRow.TransactionType != importTransactionType
                                && IsRecessoDisdetta(tacticalRow))
                            {
                                //La pratica è presente nel database come Recesso/Disdetta ma sta per essere caricata come new business
                                //Viene quindi bloccato update dello stato e verrà aggiornata data di lavorazione
                                AddLogRow(logRows, "Si stava importando una riga NB su una pratica in recesso/disdetta", row, tacticalRow.PolicyNumber.ToString());
                                nonModificareStato = true;
                            }

                            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_CONTRATTO_CHIUSO)
                            {
                                //La pratica è stata chiusa
                                AddLogRow(logRows, "Si stava cercando di modificare una pratica chiusa", row, tacticalRow.PolicyNumber.ToString());
                            }
                            else
                            {
                                var col = 2;

                                tacticalRow.Surname = GetString(ws.Cells["B" + row.ToString()].Value);
                                tacticalRow.FirstName = GetString(ws.Cells["C" + row.ToString()].Value);
                                tacticalRow.Birth = GetDate(ws.Cells["D" + row.ToString()].Value);
                                tacticalRow.CalculatedAge = GetInt(ws.Cells["E" + row.ToString()].Value);
                                tacticalRow.CodiceFiscaleAderente = GetString(wsProd.Cells["I" + row.ToString()].Value);
                                tacticalRow.Sex = GetGenderFromFiscalCode(tacticalRow.CodiceFiscaleAderente);
                                tacticalRow.ZipCode = GetString(ws.Cells["F" + row.ToString()].Value);
                                tacticalRow.Location = GetString(ws.Cells["G" + row.ToString()].Value);
                                tacticalRow.Street = GetString(ws.Cells["H" + row.ToString()].Value);
                                tacticalRow.Targa = GetString(wsProd.Cells["T" + row.ToString()].Value);
                                tacticalRow.Telaio = GetString(wsProd.Cells["U" + row.ToString()].Value);

                                if (!nonModificareStato)
                                {
                                    tacticalRow.TransactionType = GetString(ws.Cells["I" + row.ToString()].Value);
                                }

                                var newTariff = GetString(ws.Cells["J" + row.ToString()].Value);
                                if (tacticalRow.Id > 0 && tacticalRow.Tariff != newTariff)
                                {
                                    //Vuol dire che sono in UPDATE e il codice prodotto è cambiato quindi devo segnalarlo
                                    AddLogRow(logRows, string.Format("Codice prodotto variato da {0} a {1}", tacticalRow.Tariff, newTariff), row, tacticalRow.PolicyNumber.ToString());
                                }

                                //var productVita = products.Where(p => p.Code == newTariff && p.Tipologia == TIPOLOGIA_VITA).FirstOrDefault();
                                //var productDanni = products.Where(p => p.Code == newTariff && p.Tipologia == TIPOLOGIA_DANNI).FirstOrDefault();

                                var product = products.Where(p => p.Code == newTariff).FirstOrDefault();

                                if (product == null)
                                {
                                    AddLogRow(logRows, "Codice Banca non trovato: " + newTariff, row, tacticalRow.PolicyNumber.ToString());
                                }

                                tacticalRow.Tariff = newTariff;

                                tacticalRow.Duration = GetIntNullable(ws.Cells["M" + row.ToString()].Value);

                                var date = GetDateNullable(ws.Cells["P" + row.ToString()].Value);
                                if (tacticalRow.DateOfCancellation.HasValue && tacticalRow.DateOfCancellation.Value != date)
                                {
                                    //È arrivata un'estinzione totale dopo una parziale
                                    tacticalRow.DateOfCancellation3 = date;
                                }
                                else
                                {
                                    tacticalRow.DateOfCancellation = date;
                                }

                                tacticalRow.DateOfSignature = GetDateNullable(ws.Cells["Q" + row.ToString()].Value);

                                if (tacticalRow.Id > 0 && tacticalRow.ExpirationDate.HasValue && tacticalRow.ExpirationDate.Value != GetDateNullable(ws.Cells["R" + row.ToString()].Value))
                                {
                                    AddLogRow(logRows, "ExpirationDate variata", row, tacticalRow.PolicyNumber.ToString());
                                    col++;
                                }
                                else
                                {
                                    tacticalRow.ExpirationDate = GetDateNullable(ws.Cells["R" + row.ToString()].Value);
                                }

                                tacticalRow.InsuredSum = GetDoubleNullable(ws.Cells["S" + row.ToString()].Value);
                                tacticalRow.PremioLordo = GetDoubleNullable(ws.Cells["T" + row.ToString()].Value) ?? 0;
                                tacticalRow.Tassi = GetDoubleNullable(ws.Cells["U" + row.ToString()].Value);
                                tacticalRow.PremioNetto = GetDoubleNullable(ws.Cells["V" + row.ToString()].Value);
                                tacticalRow.Commissioni = GetDoubleNullable(ws.Cells["W" + row.ToString()].Value);
                                tacticalRow.DataLavorazione = DateTime.Now;

                                CheckTransationTypeErrors(logRows, tacticalRow, row);

                                if (tacticalRow.Id == 0)
                                {
                                    uow.TacticalRowRepository.Insert(tacticalRow);
                                }
                                else
                                {
                                    uow.TacticalRowRepository.Update(tacticalRow);
                                }

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        private int GetGenderFromFiscalCode(string fiscalCode)
        {
            var retVal = 0;

            if (fiscalCode.Length == 16)
            {
                var day = fiscalCode.Substring(9, 2);
                var dayInt = 0;
                if (int.TryParse(day, out dayInt))
                {
                    retVal = dayInt > 40 ? 2 : 1;
                }
            }

            return retVal;
        }

        private int ImportTactical(ExcelWorksheet ws, bool hasInsuredRate, List<string> logRows, bool leasing)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                var products = uow.ProductRepository.Get().ToList();

                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvanced(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells[row, 1].Value);
                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                tacticalRow = new TacticalRow();
                                tacticalRow.PolicyNumber = policyNumber;
                            }

                            var nonModificareStato = false;
                            var importTransactionType = GetString(ws.Cells["J" + row.ToString()].Value);
                            if (tacticalRow.TransactionType != importTransactionType
                                && IsRecessoDisdetta(tacticalRow))
                            {
                                //La pratica è presente nel database come Recesso/Disdetta ma sta per essere caricata come new business
                                //Viene quindi bloccato update dello stato e verrà aggiornata data di lavorazione
                                AddLogRow(logRows, "Si stava importando una riga NB su una pratica in recesso/disdetta", row, tacticalRow.PolicyNumber.ToString());
                                nonModificareStato = true;
                            }

                            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_CONTRATTO_CHIUSO)
                            {
                                //La pratica è stata chiusa
                                AddLogRow(logRows, "Si stava cercando di modificare una pratica chiusa", row, tacticalRow.PolicyNumber.ToString());
                            }
                            else
                            {
                                var col = 2;

                                tacticalRow.Surname = GetString(ws.Cells[row, col++].Value);
                                tacticalRow.FirstName = GetString(ws.Cells[row, col++].Value);
                                tacticalRow.Birth = GetDate(ws.Cells[row, col++].Value);
                                tacticalRow.CalculatedAge = GetInt(ws.Cells[row, col++].Value);
                                tacticalRow.Sex = GetInt(ws.Cells[row, col++].Value);
                                tacticalRow.ZipCode = GetString(ws.Cells[row, col++].Value);
                                tacticalRow.Location = GetString(ws.Cells[row, col++].Value);
                                tacticalRow.Street = GetString(ws.Cells[row, col++].Value);
                                if (!nonModificareStato)
                                {
                                    tacticalRow.TransactionType = GetString(ws.Cells[row, col++].Value);
                                }
                                else
                                {
                                    col++;
                                }

                                var newTariff = GetString(ws.Cells[row, col++].Value);
                                if (tacticalRow.Id > 0 && tacticalRow.Tariff != newTariff)
                                {
                                    //Vuol dire che sono in UPDATE e il codice prodotto è cambiato quindi devo segnalarlo
                                    AddLogRow(logRows, string.Format("Codice prodotto variato da {0} a {1}", tacticalRow.Tariff, newTariff), row, tacticalRow.PolicyNumber.ToString());
                                }

                                var productVita = products.Where(p => p.Code == newTariff && p.Tipologia == TIPOLOGIA_VITA).FirstOrDefault();
                                var productDanni = products.Where(p => p.Code == newTariff && p.Tipologia == TIPOLOGIA_DANNI).FirstOrDefault();

                                if (productVita == null || productDanni == null)
                                {
                                    AddLogRow(logRows, "Codice Tariff non trovato: " + newTariff, row, tacticalRow.PolicyNumber.ToString());
                                }

                                tacticalRow.Tariff = newTariff;

                                tacticalRow.Duration = GetIntNullable(ws.Cells[row, col++].Value);

                                var date = GetDateNullable(ws.Cells[row, col++].Value);
                                if (tacticalRow.DateOfCancellation.HasValue && tacticalRow.DateOfCancellation.Value != date)
                                {
                                    //È arrivata un'estinzione totale dopo una parziale
                                    tacticalRow.DateOfCancellation3 = date;
                                }
                                else
                                {
                                    tacticalRow.DateOfCancellation = date;
                                }

                                tacticalRow.DateOfSignature = GetDateNullable(ws.Cells[row, col++].Value);

                                if (tacticalRow.Id > 0 && tacticalRow.ExpirationDate.HasValue && tacticalRow.ExpirationDate.Value != GetDateNullable(ws.Cells[row, col].Value))
                                {
                                    AddLogRow(logRows, "ExpirationDate variata", row, tacticalRow.PolicyNumber.ToString());
                                    col++;
                                }
                                else
                                {
                                    tacticalRow.ExpirationDate = GetDateNullable(ws.Cells[row, col++].Value);
                                }

                                tacticalRow.InsuredSum = GetDoubleNullable(ws.Cells[row, col++].Value);
                                if (hasInsuredRate)
                                {
                                    tacticalRow.InsuredRate = GetDoubleNullable(ws.Cells[row, col++].Value);
                                }
                                tacticalRow.PremioLordo = GetDoubleNullable(ws.Cells[row, col++].Value) ?? 0;
                                tacticalRow.Tassi = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.PremioNetto = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.Commissioni = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.PremioVita = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.CommissioniVita = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.PremioLordoDanni = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.TassiDanni = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.PremioNettoDanni = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.CommissioniDanni = GetDoubleNullable(ws.Cells[row, col++].Value);
                                tacticalRow.DataLavorazione = DateTime.Now;

                                if (!leasing)
                                {
                                    var newValue = GetDoubleNullable(ws.Cells["AB" + row.ToString()].Value);
                                    if (!newValue.HasValue && tacticalRow.DebitoResiduo.HasValue)
                                    {
                                        AddLogRow(logRows, "Si stava cercando di svuotare il debito residuo", row, tacticalRow.PolicyNumber.ToString());
                                    }
                                    else if (newValue.HasValue)
                                    {
                                        if (!tacticalRow.DebitoResiduoIniziale.HasValue || tacticalRow.DebitoResiduoIniziale.Value == 0)
                                        {
                                            tacticalRow.DebitoResiduoIniziale = newValue;
                                        }

                                        tacticalRow.DebitoResiduo = newValue;
                                    }
                                }

                                CheckTransationTypeErrors(logRows, tacticalRow, row);

                                if (tacticalRow.Id == 0)
                                {
                                    uow.TacticalRowRepository.Insert(tacticalRow);
                                }
                                else
                                {
                                    uow.TacticalRowRepository.Update(tacticalRow);
                                }

                                if (row % 500 == 0)
                                {
                                    uow.Save();
                                }
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }

                uow.Save();
            }

            return row - 2;
        }

        private void CheckTransationTypeErrors(List<string> logRows, TacticalRow tacticalRow, int row)
        {
            if (tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_EA || tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO || tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
            {
                if (!tacticalRow.DateOfCancellation.HasValue)
                {
                    AddLogRow(logRows, "Data di cancellazione assente", row, tacticalRow.PolicyNumber.ToString());
                }
            }
        }

        private int ImportEAPWorksheet(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 1].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvancedEAP(ws.Name, row);

                            var policyNumber = GetInt(ws.Cells[row, 1].Value);

                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                            if (tacticalRow == null)
                            {
                                AddLogRow(logRows, "Numero polizza non trovato: " + policyNumber, row, tacticalRow.PolicyNumber.ToString());
                            }
                            else
                            {
                                var date = GetDate(ws.Cells[row, 3].Value);

                                if (tacticalRow.DateOfCancellation.HasValue && tacticalRow.DateOfCancellation.Value != date)
                                {
                                    //SONO IN UNA NUOVA ESTINZIONE PARZIALE
                                    if (tacticalRow.DateOfCancellation2.HasValue && tacticalRow.DateOfCancellation2.Value != date)
                                    {
                                        //Si tratta di un terza o più estinzione
                                        if (!tacticalRow.ImportoEstinzione2.HasValue)
                                        {
                                            tacticalRow.ImportoEstinzione2 = 0;
                                        }
                                    }
                                    else
                                    {
                                        //Si tratta di una seconda estinzione
                                        tacticalRow.ImportoEstinzione2 = 0;
                                    }

                                    tacticalRow.DateOfCancellation2 = date;
                                    tacticalRow.ImportoEstinzione2 += GetDoubleNullable(ws.Cells[row, 16].Value);
                                }
                                else
                                {
                                    tacticalRow.DateOfCancellation = date;
                                    tacticalRow.ImportoEstinzione = GetDoubleNullable(ws.Cells[row, 16].Value);
                                }

                                tacticalRow.CapitaleResiduo = GetDoubleNullable(ws.Cells[row, 17].Value);
                                if (!tacticalRow.DebitoResiduoIniziale.HasValue || tacticalRow.DebitoResiduoIniziale.Value == 0)
                                {
                                    tacticalRow.DebitoResiduoIniziale = GetDoubleNullable(ws.Cells["Z" + row.ToString()].Value);
                                }
                                tacticalRow.DebitoResiduo = GetDoubleNullable(ws.Cells["Z" + row.ToString()].Value);
                                tacticalRow.TransactionType = "91";

                                var newDuration = GetInt(ws.Cells["O" + row.ToString()].Value);
                                if (newDuration != tacticalRow.Duration)
                                {
                                    AddLogRow(logRows, "Durata cambiata. Prima era " + tacticalRow.Duration.Value.ToString() + " adesso è " + newDuration.ToString(), row, policyNumber.ToString());
                                    tacticalRow.ExpirationDate2 = tacticalRow.DateOfSignature.Value.AddMonths(newDuration).AddDays(-1);
                                }

                                tacticalRow.Duration = newDuration;

                                var iban = GetString(ws.Cells["S" + row.ToString()].Value);
                                if (!string.IsNullOrEmpty(iban))
                                {
                                    tacticalRow.Iban = iban;
                                }

                                uow.TacticalRowRepository.Update(tacticalRow);

                                uow.Save();
                            }

                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + ex.Message);
                    }
                }
            }

            return row - 2;
        }

        private int ImportRecessiWorksheet(ExcelWorksheet ws, List<string> logRows)
        {
            var row = 2;
            var exit = false;

            using (var uow = new UnitOfWork())
            {
                while (!exit)
                {
                    try
                    {
                        if (ws.Cells[row, 2].Value == null)
                        {
                            exit = true;
                        }
                        else
                        {
                            RaiseStepAdvancedRecessi(ws.Name, row);

                            if (GetString(ws.Cells[row, 8].Value) == "CBP")
                            {
                                var policyNumber = GetInt(ws.Cells[row, 4].Value);

                                var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                                if (tacticalRow == null)
                                {
                                    AddLogRow(logRows, "Numero polizza non trovato: " + policyNumber, row, policyNumber.ToString());
                                }
                                else
                                {
                                    tacticalRow.DateOfCancellation = GetDate(ws.Cells[row, 5].Value);
                                    tacticalRow.DataRimborso = GetDateNullable(ws.Cells[row, 17].Value);

                                    if (!tacticalRow.DebitoResiduoIniziale.HasValue || tacticalRow.DebitoResiduoIniziale.Value == 0)
                                    {
                                        tacticalRow.DebitoResiduoIniziale = GetDoubleNullable(ws.Cells["Z" + row.ToString()].Value);
                                    }
                                    tacticalRow.DebitoResiduo = GetDoubleNullable(ws.Cells["Z" + row.ToString()].Value);

                                    if (IsRecessoDisdetta(tacticalRow))
                                    {
                                        AddLogRow(logRows, "Polizza già caricata come recesso/disdetta: " + policyNumber, row, tacticalRow.PolicyNumber.ToString());
                                    }

                                    if (GetString(ws.Cells[row, 15].Value).ToLower() == "x")
                                    {
                                        if (tacticalRow.TransactionType == "91")
                                        {
                                            AddLogRow(logRows, "Polizza in estinzione parziale che diventa disdetta: " + policyNumber, row, tacticalRow.PolicyNumber.ToString());
                                        }

                                        tacticalRow.TransactionType = HelperClass.TRANSATION_TYPE_DISDETTA;
                                    }
                                    else
                                    {
                                        if (tacticalRow.TransactionType == "91")
                                        {
                                            AddLogRow(logRows, "Polizza in estinzione parziale che diventa recesso: " + policyNumber, row, tacticalRow.PolicyNumber.ToString());
                                        }

                                        tacticalRow.TransactionType = HelperClass.TRANSATION_TYPE_RECESSO;
                                    }

                                    uow.TacticalRowRepository.Update(tacticalRow);

                                    uow.Save();
                                }
                            }
                            row++;
                        }
                    }
                    catch (Exception ex)
                    {
                        var msg = ex.Message;
                        if (ex.InnerException != null)
                        {
                            msg += "\n" + ex.InnerException.Message;
                        }
                        throw new Exception("Errore importazione riga " + row.ToString() + ": " + msg);
                    }
                }
            }

            return row - 2;
        }

        private void RaiseStepAdvanced(string name, int row)
        {
            if (ImportStepAdvanced != null)
            {
                ImportStepAdvanced(name, row);
            }
        }

        private void RaiseStepAdvancedEAP(string name, int row)
        {
            if (ImportStepAdvancedEAP != null)
            {
                ImportStepAdvancedEAP(name, row);
            }
        }

        private void RaiseStepAdvancedRecessi(string name, int row)
        {
            if (ImportStepAdvancedRecessi != null)
            {
                ImportStepAdvancedRecessi(name, row);
            }
        }

        private void RaiseExportStepAdvanced(string header, int currentRowIndex, int countRow)
        {
            if (ExportStepAdvanced != null)
            {
                ExportStepAdvanced(header, currentRowIndex, countRow);
            }
        }

        private double? GetDoubleNullable(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is ExcelErrorValue)
            {
                return null;
            }

            return Convert.ToDouble(value);
        }

        private double? GetDoubleNullableFromString(object value)
        {
            if (value == null)
            {
                return null;
            }

            if (value is ExcelErrorValue)
            {
                return null;
            }

            if (value is string)
            {
                var text = value.ToString().Replace(".", ",");
                return Convert.ToDouble(text);
            }

            return Convert.ToDouble(value);
        }

        private int? GetIntNullable(object value)
        {
            if (value == null)
            {
                return null;
            }

            return Convert.ToInt32(value);
        }

        private DateTime? GetDateNullable(object value)
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

        private DateTime GetDate(object value)
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

        private string GetString(object value)
        {
            if (value == null)
            {
                return "";
            }

            return value.ToString();
        }

        private int GetInt(object value)
        {
            return Convert.ToInt32(value);
        }

        private void ImportProducts_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file delle convenzioni";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StartImportProducts(openFileDialog.FileName);
            }
        }

        private void OpenOutput_Click(object sender, EventArgs e)
        {
            Process.Start(OutputFolder.Text);
        }

        private void ImportFileEAP_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file EAP";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StartImportEAP(openFileDialog.FileName);
            }
        }

        private void ImportFileRecessi_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file recessi";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StartImportRecessi(openFileDialog.FileName);
            }
        }

        protected override bool ProcessCmdKey(ref Message msg, Keys keyData)
        {
            if (keyData == Keys.F12)
            {
                if (MessageBox.Show(this, "Aggiornare il database?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    var folder = Path.Combine(OutputFolder.Text, "Scripts");

                    if (Directory.Exists(folder))
                    {
                        var connString = "";
                        using (var uow = new UnitOfWork())
                        {
                            connString = uow.GetConnectionString();
                        }

                        var files = Directory.GetFiles(folder).ToList();
                        files.Sort();

                        try
                        {
                            foreach (var file in files)
                            {
                                var sql = File.ReadAllText(file);

                                using (var conn = new SqlConnection(connString))
                                {
                                    conn.Open();

                                    using (var cmd = new SqlCommand(sql, conn))
                                    {
                                        cmd.ExecuteNonQuery();
                                    }
                                }
                            }

                            if (files.Count == 0)
                            {
                                HelperClass.ShowWarning(this, "Nessuno script trovato");
                            }
                            else
                            {
                                HelperClass.ShowInfo(this, "Database aggiornato");
                            }
                        }
                        catch (Exception ex)
                        {
                            HelperClass.ShowError(this, ex.Message);
                        }
                    }
                    else
                    {
                        HelperClass.ShowError(this, "Cartella 'Scripts' non esiste");
                    }
                }

                return true;
            }
            else if (keyData == Keys.F11)
            {
                FontFactory.Register(GetFilePath("Modellini\\calibri.ttf"));
                FontFactory.Register(GetFilePath("Modellini\\calibriBold.ttf"));

                using (var uow = new UnitOfWork())
                {
                    //var tacticalRow = uow.TacticalRowRepository.Get(t => t.Id == 100517).First(); // TipoEstinzione = 1
                    var tacticalRow = uow.TacticalRowRepository.Get(t => t.Id == 119573).First(); // TipoEstinzione = 2

                    using (var modellinoParziali = new ExcelPackage(new FileInfo(GetFilePath("Modellini\\EstinzioniParzialiDefinitivo.xlsx"))))
                    {
                        var worksheetEstinzioniParziali = modellinoParziali.Workbook.Worksheets["Refund calculator (Other) - NEW"];

                        var productVita = uow.ProductRepository.Get(p => p.Code == tacticalRow.Tariff && p.Tipologia == TIPOLOGIA_VITA).FirstOrDefault();
                        var productDanni = uow.ProductRepository.Get(p => p.Code == tacticalRow.Tariff && p.Tipologia == TIPOLOGIA_DANNI).FirstOrDefault();

                        //Estizione parziale dopo parziale
                        worksheetEstinzioniParziali.Cells["C9"].Value = productVita.CodeCNPSI;
                        worksheetEstinzioniParziali.Cells["C10"].Value = tacticalRow.DateOfSignature;
                        worksheetEstinzioniParziali.Cells["C11"].Value = tacticalRow.ExpirationDate;
                        worksheetEstinzioniParziali.Cells["C12"].Value = tacticalRow.DateOfCancellation3;
                        worksheetEstinzioniParziali.Cells["C13"].Value = tacticalRow.InsuredSum.Value - tacticalRow.PremioLordo;
                        worksheetEstinzioniParziali.Cells["C14"].Value = (tacticalRow.CapitaleResiduo ?? 0) + Math.Abs(tacticalRow.ImportoEstinzione.Value);
                        worksheetEstinzioniParziali.Cells["C15"].Value = tacticalRow.DebitoResiduo ?? 0;
                        worksheetEstinzioniParziali.Cells["C16"].Value = Math.Abs(tacticalRow.ImportoEstinzione.Value);
                        worksheetEstinzioniParziali.Cells["C17"].Value = tacticalRow.CapitaleResiduo ?? 0;
                        worksheetEstinzioniParziali.Cells["C18"].Value = tacticalRow.ImportoEstinzione2 ?? 0;
                        worksheetEstinzioniParziali.Calculate();

                        CreaLetteraRecessoLeasing(OutputFolder.Text, tacticalRow, DateTime.Now);

                        ShowInfo("Ok");

                        OpenOutput_Click(null, null);
                    }
                }
            }

            return base.ProcessCmdKey(ref msg, keyData);
        }

        private void ButtonImportDatiAggiuntivi_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file lavorativo";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StartImportDatiAggiuntivi(openFileDialog.FileName);
            }
        }

        private void EsportaDatiBanca_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file lavorativo";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                EsportazioneDatiBanca(openFileDialog.FileName, PagamentoAccantonato.Checked);
            }
        }

        private void EsportazioneDatiBanca(string fileName, bool pagamentoAccantonato = false)
        {
            var logRows = new List<string>();

            try
            {
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                using (var monthly = new ExcelPackage(new FileInfo(fileName)))
                {
                    var wsMonthly = monthly.Workbook.Worksheets[1];

                    using (var datiBanca = new ExcelPackage())
                    {
                        var ws = datiBanca.Workbook.Worksheets.Add("Dati bonifici");

                        var inputRow = 1;
                        ws.Cells["A" + inputRow.ToString()].Value = "Tipologia";
                        ws.Cells["B" + inputRow.ToString()].Value = "Cognome";
                        ws.Cells["C" + inputRow.ToString()].Value = "Nome";
                        ws.Cells["D" + inputRow.ToString()].Value = "IBAN";
                        ws.Cells["E" + inputRow.ToString()].Value = "Valore";
                        ws.Cells["F" + inputRow.ToString()].Value = "Causale";
                        ws.Cells["G" + inputRow.ToString()].Value = "InternalId";
                        ws.Cells["H" + inputRow.ToString()].Value = "PagamentoAccantonato";
                        ws.Cells["I" + inputRow.ToString()].Value = "DataCancellazione";

                        ws.Column(5).Style.Numberformat.Format = "#0.00";
                        ws.Column(9).Style.Numberformat.Format = "dd/MM/yyyy";

                        inputRow++;

                        var exit = false;

                        var rowOutput = 2;

                        using (var uow = new UnitOfWork())
                        {
                            while (!exit)
                            {
                                if (wsMonthly.Cells[inputRow, 1].Value == null)
                                {
                                    exit = true;
                                }
                                else
                                {
                                    try
                                    {
                                        var importoVita = (double)wsMonthly.Cells["L" + inputRow.ToString()].Value;
                                        var importoDanni = (double)wsMonthly.Cells["L" + (inputRow + 1).ToString()].Value;

                                        if (GetString(wsMonthly.Cells["G" + inputRow.ToString()].Value).ToLower() != "revoca"
                                            && (importoVita + importoDanni > 0))
                                        {
                                            int policyNumber = Convert.ToInt32(wsMonthly.Cells["F" + inputRow.ToString()].Value);
                                            var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).First();
                                            var beneficiari = uow.BeneficiaryRepository.Get(b => b.TacticalRowId == tacticalRow.Id).ToList();

                                            if (beneficiari.Count > 0)
                                            {
                                                var sommaQuota = 0D;
                                                beneficiari.ForEach(b => sommaQuota += b.Share);
                                                if (sommaQuota != 100)
                                                {
                                                    logRows.Add(string.Format("Errore in riga monthly {0}: la somma delle quote dei beneficiari è diversa da 100", inputRow));
                                                }

                                                if (importoVita > 0)
                                                {
                                                    if (importoDanni < 0)
                                                    {
                                                        importoVita += importoDanni;
                                                    }


                                                    var values = GetQuotaValues(beneficiari, importoVita);

                                                    for (var i = 0; i < values.Count; i++)
                                                    {
                                                        var value = values[i];
                                                        var beneficiario = beneficiari[i];

                                                        ws.Cells["A" + rowOutput.ToString()].Value = 0;
                                                        ws.Cells["B" + rowOutput.ToString()].Value = beneficiario.LastName;
                                                        ws.Cells["C" + rowOutput.ToString()].Value = beneficiario.FirstName;
                                                        ws.Cells["D" + rowOutput.ToString()].Value = beneficiario.Iban;
                                                        ws.Cells["E" + rowOutput.ToString()].Value = value;
                                                        ws.Cells["F" + rowOutput.ToString()].Value = GetCausale(tacticalRow, GetString(wsMonthly.Cells["G" + inputRow.ToString()].Value));
                                                        ws.Cells["G" + rowOutput.ToString()].Value = tacticalRow.Id;
                                                        ws.Cells["H" + rowOutput.ToString()].Value = pagamentoAccantonato ? "SI" : "NO";
                                                        ws.Cells["I" + rowOutput.ToString()].Value = GetNormalDateString(GetDateNullable(wsMonthly.Cells["H" + inputRow.ToString()].Value));

                                                        rowOutput++;
                                                    }
                                                }

                                                if (importoDanni > 0)
                                                {
                                                    var values = GetQuotaValues(beneficiari, importoDanni);

                                                    for (var i = 0; i < values.Count; i++)
                                                    {
                                                        var value = values[i];
                                                        var beneficiario = beneficiari[i];

                                                        ws.Cells["A" + rowOutput.ToString()].Value = 1;
                                                        ws.Cells["B" + rowOutput.ToString()].Value = beneficiario.LastName;
                                                        ws.Cells["C" + rowOutput.ToString()].Value = beneficiario.FirstName;
                                                        ws.Cells["D" + rowOutput.ToString()].Value = beneficiario.Iban;
                                                        ws.Cells["E" + rowOutput.ToString()].Value = value;
                                                        ws.Cells["F" + rowOutput.ToString()].Value = GetCausale(tacticalRow, GetString(wsMonthly.Cells["G" + inputRow.ToString()].Value));
                                                        ws.Cells["G" + rowOutput.ToString()].Value = tacticalRow.Id;
                                                        ws.Cells["H" + rowOutput.ToString()].Value = pagamentoAccantonato ? "SI" : "NO";
                                                        ws.Cells["I" + rowOutput.ToString()].Value = GetNormalDateString(GetDateNullable(wsMonthly.Cells["H" + (inputRow + 1).ToString()].Value));

                                                        rowOutput++;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                                if (importoVita > 0)
                                                {
                                                    RiempiRigaImportoVita(wsMonthly, ws, inputRow, rowOutput, pagamentoAccantonato, logRows, uow);
                                                    rowOutput++;
                                                }

                                                if (importoDanni > 0)
                                                {
                                                    RiempiRigaImportoDanni(wsMonthly, ws, inputRow + 1, rowOutput, pagamentoAccantonato, logRows, uow);
                                                    rowOutput++;
                                                }
                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        logRows.Add(string.Format("Errore in riga monthly {0}: {1}", inputRow, ex.Message));
                                    }

                                    inputRow += 2;
                                }
                            }
                        }

                        datiBanca.SaveAs(new FileInfo(GetCurrentBankName(fileName)));
                    }
                }

                Cursor = Cursors.Default;
                ShowInfo("File esportati");

                Process.Start(OutputFolder.Text);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
            finally
            {
                WriteLog(logRows, "EsportazioneBanca");
            }
        }

        private List<double> GetQuotaValues(List<Beneficiary> beneficiari, double importoTotale)
        {
            var retVal = new List<double>();
            var somma = 0D;
            for (var i = 0; i < beneficiari.Count - 1; i++)
            {
                var value = Math.Round(importoTotale * beneficiari[i].Share / 100, 2);
                somma += value;
                retVal.Add(value);
            }

            retVal.Add(importoTotale - somma);

            return retVal;
        }

        private void RiempiRigaImportoVita(ExcelWorksheet wsInput, ExcelWorksheet wsOutput, int inputRow, int outputRow, bool pagamentoAccantonato, List<string> logRows, UnitOfWork uow)
        {
            wsOutput.Cells["A" + outputRow.ToString()].Value = 0;

            int policyNumber = Convert.ToInt32(wsInput.Cells["F" + inputRow.ToString()].Value);
            var tacticalRow = uow.TacticalRowRepository.Get(t => t.PolicyNumber == policyNumber).First();

            if (!string.IsNullOrEmpty(tacticalRow.CognomeCoIntestario) && !string.IsNullOrEmpty(tacticalRow.NomeCoIntestatario))
            {
                var cognome = wsInput.Cells["C" + inputRow.ToString()].Value + " " + wsInput.Cells["D" + inputRow.ToString()].Value;
                var nome = tacticalRow.CognomeCoIntestario + " " + tacticalRow.NomeCoIntestatario;
                wsOutput.Cells["B" + outputRow.ToString()].Value = cognome;
                wsOutput.Cells["C" + outputRow.ToString()].Value = nome;
            }
            else
            {
                wsOutput.Cells["B" + outputRow.ToString()].Value = wsInput.Cells["C" + inputRow.ToString()].Value;
                wsOutput.Cells["C" + outputRow.ToString()].Value = wsInput.Cells["D" + inputRow.ToString()].Value;
            }

            wsOutput.Cells["D" + outputRow.ToString()].Value = tacticalRow.Iban;

            if (string.IsNullOrEmpty(tacticalRow.Iban))
            {
                logRows.Add(string.Format("IBAN mancante in riga monthly {0}", inputRow));
            }

            var importoVita = (double)wsInput.Cells["L" + inputRow.ToString()].Value;
            var importoDanni = (double)wsInput.Cells["L" + (inputRow + 1).ToString()].Value;

            if (importoDanni < 0)
            {
                importoVita += importoDanni;
            }

            var causale = GetCausale(tacticalRow, GetString(wsInput.Cells["G" + inputRow.ToString()].Value));

            if (string.IsNullOrEmpty(causale))
            {
                logRows.Add(string.Format("Causale mancante in riga monthly {0}", inputRow));
            }

            wsOutput.Cells["E" + outputRow.ToString()].Value = importoVita;
            wsOutput.Cells["F" + outputRow.ToString()].Value = causale;
            wsOutput.Cells["G" + outputRow.ToString()].Value = tacticalRow.Id;
            wsOutput.Cells["H" + outputRow.ToString()].Value = pagamentoAccantonato ? "SI" : "NO";
            wsOutput.Cells["I" + outputRow.ToString()].Value = GetNormalDateString(GetDateNullable(wsInput.Cells["H" + inputRow.ToString()].Value));
        }

        private void RiempiRigaImportoDanni(ExcelWorksheet wsInput, ExcelWorksheet wsOutput, int inputRow, int outputRow, bool pagamentoAccantonato, List<string> logRows, UnitOfWork uow)
        {
            wsOutput.Cells["A" + outputRow.ToString()].Value = 1;

            int policyNumber = Convert.ToInt32(wsInput.Cells["F" + inputRow.ToString()].Value);
            var tacticalRow = uow.TacticalRowRepository.Get(t => t.PolicyNumber == policyNumber).First();

            if (!string.IsNullOrEmpty(tacticalRow.CognomeCoIntestario) && !string.IsNullOrEmpty(tacticalRow.NomeCoIntestatario))
            {
                var cognome = wsInput.Cells["C" + inputRow.ToString()].Value + " " + wsInput.Cells["D" + inputRow.ToString()].Value;
                var nome = tacticalRow.CognomeCoIntestario + " " + tacticalRow.NomeCoIntestatario;
                wsOutput.Cells["B" + outputRow.ToString()].Value = cognome;
                wsOutput.Cells["C" + outputRow.ToString()].Value = nome;
            }
            else
            {
                wsOutput.Cells["B" + outputRow.ToString()].Value = wsInput.Cells["C" + inputRow.ToString()].Value;
                wsOutput.Cells["C" + outputRow.ToString()].Value = wsInput.Cells["D" + inputRow.ToString()].Value;
            }

            wsOutput.Cells["D" + outputRow.ToString()].Value = tacticalRow.Iban;

            if (string.IsNullOrEmpty(tacticalRow.Iban))
            {
                logRows.Add(string.Format("IBAN mancante in riga monthly {0}", inputRow));
            }

            var importoVita = (double)wsInput.Cells["L" + (inputRow - 1).ToString()].Value;
            var importoDanni = (double)wsInput.Cells["L" + inputRow.ToString()].Value;

            if (importoVita < 0)
            {
                importoDanni += importoVita;
            }

            var causale = GetCausale(tacticalRow, GetString(wsInput.Cells["G" + inputRow.ToString()].Value));

            if (string.IsNullOrEmpty(causale))
            {
                logRows.Add(string.Format("Causale mancante in riga monthly {0}", inputRow));
            }

            wsOutput.Cells["E" + outputRow.ToString()].Value = importoDanni;
            wsOutput.Cells["F" + outputRow.ToString()].Value = causale;
            wsOutput.Cells["G" + outputRow.ToString()].Value = tacticalRow.Id;
            wsOutput.Cells["H" + outputRow.ToString()].Value = pagamentoAccantonato ? "SI" : "NO";
            wsOutput.Cells["I" + outputRow.ToString()].Value = GetNormalDateString(GetDateNullable(wsInput.Cells["H" + inputRow.ToString()].Value));
        }

        private string GetNormalDateString(DateTime? date)
        {
            if (!date.HasValue)
            {
                return "";
            }

            return date.Value.ToString("dd/MM/yyyy");
        }

        private string GetCausale(TacticalRow tr, string tipologiaCancellazione)
        {
            var retVal = "";

            retVal = string.Format("{0} {1} {2} Rimborso Premio per {3} CNP Santander Insurance",
                                    tr.PolicyNumber,
                                    tr.FirstName,
                                    tr.Surname,
                                    tipologiaCancellazione.ToLower());

            //switch (tipologiaCancellazione.ToLower())
            //{
            //    case "estinzione anticipata":
            //        retVal = string.Format("Rimborso Premio pagato e non goduto per Estinzione Anticipata CNP Santander Insurance - {0} {1} {2}", tr.PolicyNumber, tr.Surname, tr.FirstName);
            //        break;
            //    case "estinzione parziale":
            //        retVal = string.Format("Rimborso Premio pagato e non goduto per Estinzione Anticipata Parziale CNP Santander Insurance - {0} {1} {2}", tr.PolicyNumber, tr.Surname, tr.FirstName);
            //        break;
            //    case "disdetta":
            //        retVal = string.Format("Rimborso Premio per disdetta CNP Santander Insurance - {0} {1} {2}", tr.PolicyNumber, tr.Surname, tr.FirstName);
            //        break;
            //    case "recesso":
            //        retVal = string.Format("Rimborso Premio per recesso CNP Santander Insurance - {0} {1} {2}", tr.PolicyNumber, tr.Surname, tr.FirstName);
            //        break;
            //    default:
            //        break;
            //}

            return retVal;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var bankForm = new Bank();
            bankForm.ShowDialog(this);
        }

        private void Main_Load(object sender, EventArgs e)
        {

        }

        #region Creazione Lettere

        private void CreaLetteraRecesso(string baseOutputFolder, TacticalRow tacticalRow, DateTime date)
        {
            var pdfTemplate = string.IsNullOrEmpty(Program.TEMPLATE_RECESSO) ?
                                  GetFilePath(@"Modellini\" + "Recesso.pdf")
                                : GetFilePath(@"Modellini\" + Program.TEMPLATE_RECESSO);
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            int numberOfPages;
            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                PdfReader.unethicalreading = true;
                numberOfPages = pdfReader.NumberOfPages;

                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var pdfFormFields = pdfStamper.AcroFields;

                    var tasse = tacticalRow.PremioLordo * 0.025;
                    var premioNetto = tacticalRow.PremioLordo - tasse;
                    var importoRimborso = premioNetto;

                    pdfFormFields.SetField("NumeroAdesione", tacticalRow.PolicyNumber.ToString());
                    pdfFormFields.SetField("NomeCognome", GetTitoloFromGender(tacticalRow.Sex) + " " + tacticalRow.FirstName + " " + tacticalRow.Surname);
                    pdfFormFields.SetField("Indirizzo", tacticalRow.Street);
                    pdfFormFields.SetField("DettagliIndirizzo", tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia);
                    pdfFormFields.SetField("Data", GetLetterDate(date));
                    pdfFormFields.SetField("Iban", tacticalRow.Iban);
                    pdfFormFields.SetField("PremioLordo", WriteCurrencyString(tacticalRow.PremioLordo));
                    pdfFormFields.SetField("Tasse", WriteCurrencyString(tasse));
                    pdfFormFields.SetField("PremioNetto", WriteCurrencyString(premioNetto));
                    pdfFormFields.SetField("ImportoRimborso", WriteCurrencyString(importoRimborso));
                    pdfStamper.AcroFields.GenerateAppearances = true;
                    pdfStamper.FormFlattening = true;

                    pdfStamper.Close();
                }
            }
        }

        private string GetTitoloFromGender(int sex)
        {
            var retVal = "";

            if (sex == 0)
            {
                retVal = "Spett.le";
            }
            else if (sex == 1)
            {
                retVal = "Sig.";
            }
            else if (sex == 2)
            {
                retVal = "Sig.ra";
            }

            return retVal;
        }

        private void CreaLetteraRecessoLeasing(string baseOutputFolder, TacticalRow tacticalRow, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\RecessoLeasing.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 678, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 165, 468, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy"), 130, 443, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraDisdetta(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\Disdetta.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 703, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy"), 310, 533.5F, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 103, 522, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 220, 360, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 360, 360, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 485, 360, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 220, 340, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 360, 340, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 485, 340, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 220, 320, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 360, 320, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C20"].Value), 485, 320, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 220, 300, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 360, 300, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C46"].Value), 485, 300, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, 233, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, 220, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C26"].Value), 420, 207, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C27"].Value), 420, 194, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    var imposte = (double)worksheetEstinzioni.Cells["D42"].Value - (double)worksheetEstinzioni.Cells["D29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D42"].Value), 315, 640, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 315, 628, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D29"].Value), 315, 616, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 315, 592, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 315, 580, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D43"].Value), 315, 419, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 315, 395, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 315, 383, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 315, 349, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 315, 315, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    imposte = (double)worksheetEstinzioni.Cells["E42"].Value - (double)worksheetEstinzioni.Cells["E29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E42"].Value), 315, 241, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 315, 229, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E29"].Value), 315, 217, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 315, 195, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 315, 181, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA TERZA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(3);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E43"].Value), 315, 606, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 315, 582, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 315, 570, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 315, 514, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 315, 479, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipate(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipata.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 698, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Iban ?? ""), 351, 560, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 220, 380, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 360, 380, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 485, 380, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 220, 360, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 360, 360, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 485, 360, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 220, 340, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 360, 340, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C20"].Value), 485, 340, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 220, 320, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 360, 320, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C46"].Value), 485, 320, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 385, 237, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 385, 224, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C26"].Value), 385, 211, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C27"].Value), 385, 198, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    var imposte = (double)worksheetEstinzioni.Cells["D42"].Value - (double)worksheetEstinzioni.Cells["D29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D42"].Value), 315, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 315, 672, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D29"].Value), 315, 658, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 315, 630, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 315, 617, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D43"].Value), 350, 441, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 350, 414, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 350, 401, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 350, 361, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 350, 320, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    imposte = (double)worksheetEstinzioni.Cells["E42"].Value - (double)worksheetEstinzioni.Cells["E29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E42"].Value), 350, 238, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 350, 225, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E29"].Value), 350, 212, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 185, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 171, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA TERZA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(3);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E43"].Value), 350, 605, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 577, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 564, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 350, 497, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 350, 457, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipataZero(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipataZero.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 699, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 220, 377, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 360, 377, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 485, 377, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 220, 357, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 360, 357, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 485, 357, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 220, 336, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 360, 336, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C20"].Value), 485, 336, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 220, 316, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 360, 316, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C46"].Value), 485, 316, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 385, 236, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 385, 223, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C26"].Value), 385, 210, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C27"].Value), 385, 197, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    var imposte = (double)worksheetEstinzioni.Cells["D42"].Value - (double)worksheetEstinzioni.Cells["D29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D42"].Value), 315, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 315, 672, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D29"].Value), 315, 658, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 315, 630, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 315, 617, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D43"].Value), 350, 441, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 350, 414, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 350, 401, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 350, 361, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 350, 320, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    imposte = (double)worksheetEstinzioni.Cells["E42"].Value - (double)worksheetEstinzioni.Cells["E29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E42"].Value), 350, 238, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 350, 225, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E29"].Value), 350, 212, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 185, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 171, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA TERZA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(3);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E43"].Value), 350, 605, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 577, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 564, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 350, 497, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 350, 457, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipataSingolo(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipataSingolo.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 705, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Iban ?? ""), 351, 552, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 220, 377, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 360, 377, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 485, 377, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 220, 357, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 360, 357, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 485, 357, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 220, 336, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 360, 336, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C20"].Value), 485, 336, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 220, 316, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 360, 316, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C46"].Value), 485, 316, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 385, 230, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 385, 217, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C26"].Value), 385, 204, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C27"].Value), 385, 191, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    var imposte = (double)worksheetEstinzioni.Cells["D42"].Value - (double)worksheetEstinzioni.Cells["D29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D42"].Value), 315, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 315, 672, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D29"].Value), 315, 658, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 315, 630, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 315, 617, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D43"].Value), 350, 441, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 350, 414, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 350, 401, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D20"].Value), 350, 361, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 350, 320, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    imposte = (double)worksheetEstinzioni.Cells["E42"].Value - (double)worksheetEstinzioni.Cells["E29"].Value;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E42"].Value), 350, 238, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(imposte), 350, 225, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E29"].Value), 350, 212, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 185, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 171, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA TERZA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(3);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E43"].Value), 350, 605, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 350, 577, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 350, 564, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E20"].Value), 350, 497, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 350, 457, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipateNoIban(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneNoIban.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 145, 726, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 620, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Egregio Signore," : "Gentile Signora,"), 65, 520, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 470, 498, 0);

                    var cancellationDate = tacticalRow.DateOfCancellation;
                    if (tacticalRow.DateOfCancellation3.HasValue)
                    {
                        cancellationDate = tacticalRow.DateOfCancellation3;
                    }
                    else if (tacticalRow.DateOfCancellation2.HasValue)
                    {
                        cancellationDate = tacticalRow.DateOfCancellation2;
                    }

                    if (cancellationDate.HasValue)
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, cancellationDate.Value.ToString("dd/MM/yyyy"), 65, 443, 0);
                    }

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D45"].Value), 205, 243, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E45"].Value), 340, 243, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 457, 243, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D44"].Value), 205, 223, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E44"].Value), 340, 223, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 457, 223, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D46"].Value), 205, 183, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E46"].Value), 340, 183, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C46"].Value), 457, 183, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA TERZA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(3);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 70, 520, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 70, 510, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 70, 500, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetLetterDate(date), 395, 404, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 135, 337, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniParziali(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneParziale.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 316, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 316, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 316, 680, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 316, 620, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 390, 465, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 377, 448, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C11"].Value).ToString("dd/MM/yyyy"), 377, 431, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C28"].Value), 377, 414, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 390, 397, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 377, 380, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C29"].Value), 377, 363, 0);
                    //pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C44"].Value), 390, 346, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C45"].Value), 390, 329, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 267, 287, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 333, 287, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyStringPercentual(worksheetEstinzioni.Cells["C50"].Value), 157, 274, 0);

                    //TABELLA
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyStringPercentual(worksheetEstinzioni.Cells["C50"].Value), 375, 233, 0);
                    var c45 = GetDoubleNullable(worksheetEstinzioni.Cells["C45"].Value);
                    var c50 = GetDoubleNullable(worksheetEstinzioni.Cells["C50"].Value);
                    if (c45.HasValue && c50.HasValue)
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (c45.Value * c50.Value).ToString("#,##0.00"), 375, 216, 0);
                    }
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C54"].Value), 375, 199, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C55"].Value), 375, 182, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniParzialiUltimaVersione(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneParziale.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 316, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 316, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 316, 680, 0);
                    if (!string.IsNullOrEmpty(tacticalRow.Iban))
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 316, 620, 0);
                    }

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 390, 365, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 390, 266, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 10);

                    var baseHeight = 227F;
                    var increment = 12.5F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C40"].Value), 320, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 320, baseHeight - increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 320, baseHeight - 2 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 320, baseHeight - 3 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 320, baseHeight - 4 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(50F), 320, baseHeight - 6 * increment, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C15"].Value), 270, 673.5F, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 113, 659, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C65"].Value), 290, 645F, 0);

                    baseHeight = 577F;
                    increment = 18F;

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C65"].Value) + "%", 380, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C66"].Value), 380, baseHeight - increment, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C74"].Value), 380, baseHeight - 3 * increment, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        /// <summary>
        /// Uguale a CreaLetteraEstinzioniParzialiUltimaVersione ma usa modellino diverso
        /// </summary>
        /// <param name="baseOutputFolder"></param>
        /// <param name="tacticalRow"></param>
        /// <param name="worksheetEstinzioni"></param>
        /// <param name="date"></param>
        private void CreaLetteraEstinzioniParzialiDopoParzialeUltimaVersione(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneParziale.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 685, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 316, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 316, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 316, 680, 0);
                    if (!string.IsNullOrEmpty(tacticalRow.Iban))
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 316, 620, 0);
                    }

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 390, 365, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 390, 266, 0);

                    pdfContentByte.SetFontAndSize(font.BaseFont, 10);

                    var baseHeight = 227F;
                    var increment = 12.5F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C40"].Value), 320, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 320, baseHeight - increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 320, baseHeight - 2 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 320, baseHeight - 3 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 320, baseHeight - 4 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(50F), 320, baseHeight - 6 * increment, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C16"].Value), 270, 673.5F, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 113, 659, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C65"].Value), 290, 645F, 0);

                    baseHeight = 577F;
                    increment = 18F;

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C65"].Value) + "%", 380, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C66"].Value), 380, baseHeight - increment, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C74"].Value), 380, baseHeight - 3 * increment, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipateDopoParzialeUltimaVersione(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipataDopoParziale.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 695, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 316, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 316, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 316, 680, 0);
                    if (!string.IsNullOrEmpty(tacticalRow.Iban))
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 316, 620, 0);
                    }

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 67, 545, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Iban ?? ""), 67, 533, 0);

                    var baseHeight = 435F;
                    var increment = 12.5F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, baseHeight - increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 420, baseHeight - 2 * increment, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 420, baseHeight - 3 * increment, 0);

                    baseHeight = 299F;
                    increment = 12.5F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D84"].Value), 420, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D91"].Value), 420, baseHeight - increment, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D92"].Value), 420, 229, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    baseHeight = 172F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E84"].Value), 420, baseHeight, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E91"].Value), 420, baseHeight - increment, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();


                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E92"].Value), 420, 706, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniTotali(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneTotale.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 705, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 320, 717, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 320, 707, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 320, 697, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 320, 641, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 70, 557, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Iban ?? ""), 108, 545, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O53"].Value), 220, 410, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P53"].Value), 360, 410, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["N53"].Value), 485, 410, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O52"].Value), 220, 390, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P52"].Value), 360, 390, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["N52"].Value), 485, 390, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O54"].Value), 220, 370, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P54"].Value), 360, 370, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["N54"].Value), 485, 370, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O55"].Value), 220, 350, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P55"].Value), 360, 350, 0);
                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["N55"].Value), 485, 350, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 400, 276, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["N12"].Value).ToString("dd/MM/yyyy"), 400, 264, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C11"].Value).ToString("dd/MM/yyyy"), 400, 252, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetInt(worksheetEstinzioni.Cells["N28"].Value).ToString(), 400, 240, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetInt(worksheetEstinzioni.Cells["N29"].Value).ToString(), 400, 228, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O51"].Value), 381, 692, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O54"].Value), 348, 680.5F, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["O55"].Value), 310, 623, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P51"].Value), 377, 565.5F, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P54"].Value), 350, 553, 0);

                    pdfContentByte.SetFontAndSize(fontBold.BaseFont, 11);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["P54"].Value), 315, 495.5F, 0);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraDisdettaNuova(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\Disdetta.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 713, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 710, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 630, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.DateOfCancellation.Value.ToString("dd/MM/yyyy"), 310, 587.5F, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 110, 576.5F, 0);

                    var baseValue = 450F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 39, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 26, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 420, baseValue + 13, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 420, baseValue, 0);

                    baseValue = 170;
                    var multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D25"].Value), 360, baseValue + multiplier * 8, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 7, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D30"].Value), 360, baseValue + multiplier * 2, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    baseValue = 559;
                    multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E25"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E30"].Value), 360, baseValue + multiplier * 3, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipataZeroNuova(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipataZero.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 703, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 710, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 700, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 690, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 630, 0);

                    var baseValue = 472F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 39, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 26, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 420, baseValue + 13, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 420, baseValue, 0);

                    baseValue = 193.2F;
                    var multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D25"].Value), 360, baseValue + multiplier * 8, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 7, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D30"].Value), 360, baseValue + multiplier * 2, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    baseValue = 559;
                    multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E25"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E30"].Value), 360, baseValue + multiplier * 3, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipataSingoloNuova(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipataSingolo.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 723, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 730, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 720, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 710, 0);
                    if (!string.IsNullOrEmpty(tacticalRow.Iban))
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 650, 0);
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Iban, 360, 611F, 0);
                    }

                    var baseValue = 495F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 39, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 26, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 420, baseValue + 13, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 420, baseValue, 0);

                    baseValue = 215;
                    var multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D25"].Value), 360, baseValue + multiplier * 8, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 7, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D30"].Value), 360, baseValue + multiplier * 2, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    baseValue = 559;
                    multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E25"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E30"].Value), 360, baseValue + multiplier * 3, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void CreaLetteraEstinzioniAnticipateNoIbanNuova(string baseOutputFolder, TacticalRow tacticalRow, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneNoIban.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow, "_noiban");

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 145, 726, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 740, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 730, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 720, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 660, 0);
                    //pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Egregio Signore," : "Gentile Signora,"), 65, 520, 0);

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 470, 565, 0);

                    var cancellationDate = tacticalRow.DateOfCancellation;
                    if (tacticalRow.DateOfCancellation3.HasValue)
                    {
                        cancellationDate = tacticalRow.DateOfCancellation3;
                    }
                    else if (tacticalRow.DateOfCancellation2.HasValue)
                    {
                        cancellationDate = tacticalRow.DateOfCancellation2;
                    }

                    if (cancellationDate.HasValue)
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, cancellationDate.Value.ToString("dd/MM/yyyy"), 65, 512, 0);
                    }

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 70, 655, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 70, 645, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 70, 635, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetLetterDate(date), 395, 538, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 140, 471, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        private void _CreaLetteraEstinzioniAnticipataNuova(string baseOutputFolder, TacticalRow tacticalRow, ExcelWorksheet worksheetEstinzioni, DateTime date)
        {
            string pdfTemplate = GetFilePath("Modellini\\EstinzioneAnticipata.pdf");
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var font = FontFactory.GetFont("Calibri", 28, BaseColor.BLACK);
                    var fontBold = FontFactory.GetFont("Calibri Bold", 28, BaseColor.BLACK);

                    //STAMPA PRIMA PAGINA
                    var pdfContentByte = pdfStamper.GetOverContent(1);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.PolicyNumber.ToString(), 150, 695, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname, 340, 730, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Street, 340, 720, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia, 340, 710, 0);
                    if (!string.IsNullOrEmpty(tacticalRow.Iban))
                    {
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, "Milano, " + GetLetterDate(date), 340, 650, 0);
                        pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, tacticalRow.Iban, 370, 571F, 0);
                    }

                    var baseValue = 450F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C10"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 39, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, GetDate(worksheetEstinzioni.Cells["C12"].Value).ToString("dd/MM/yyyy"), 420, baseValue + 26, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 420, baseValue + 13, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 420, baseValue, 0);

                    baseValue = 170;
                    var multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D25"].Value), 360, baseValue + multiplier * 8, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 7, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C14"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C13"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D30"].Value), 360, baseValue + multiplier * 2, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["D60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    //STAMPA SECONDA PAGINA
                    pdfContentByte = pdfStamper.GetOverContent(2);
                    pdfContentByte.SetFontAndSize(font.BaseFont, 11);
                    pdfContentByte.SetColorFill(BaseColor.BLACK);
                    pdfContentByte.BeginText();

                    baseValue = 533.5F;
                    multiplier = 11.6F;
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E25"].Value), 360, baseValue + multiplier * 6, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C37"].Value), 360, baseValue + multiplier * 5, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["C38"].Value), 360, baseValue + multiplier * 4, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E30"].Value), 360, baseValue + multiplier * 3, 0);
                    pdfContentByte.ShowTextAligned(Element.ALIGN_LEFT, WriteCurrencyString(worksheetEstinzioni.Cells["E60"].Value), 360, baseValue, 0);

                    pdfContentByte.EndText();
                    pdfContentByte.Stroke();

                    pdfStamper.Close();
                }
            }
        }

        public string CreaLetteraEstinzioniAnticipataNuova(string baseOutputFolder, TacticalRow tacticalRow, Product product, DateTime date)
        {
            var pdfTemplate = string.IsNullOrEmpty(Program.TEMPLATE_ESTINZIONE_ANTICIPATA) ?
                                  GetFilePath(@"Modellini\" + "EstinzioneAnticipata.pdf")
                                : GetFilePath(@"Modellini\" + Program.TEMPLATE_ESTINZIONE_ANTICIPATA);
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            int numberOfPages;
            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                PdfReader.unethicalreading = true;
                numberOfPages = pdfReader.NumberOfPages;

                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var pdfFormFields = pdfStamper.AcroFields;

                    var premioDaRimborsare = GetPremioDaRimborsare(tacticalRow, product);
                    var durataGiorni = GetDurataGiorni(tacticalRow);
                    var durataGoduta = GetGiorniGoduti(tacticalRow);

                    pdfFormFields.SetField("NumeroAdesione", tacticalRow.PolicyNumber.ToString());
                    pdfFormFields.SetField("NomeCognome", (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname);
                    pdfFormFields.SetField("Indirizzo", tacticalRow.Street);
                    pdfFormFields.SetField("DettagliIndirizzo", tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia);
                    pdfFormFields.SetField("Data", GetLetterDate(date));
                    pdfFormFields.SetField("Iban", tacticalRow.Iban);
                    pdfFormFields.SetField("DataDecorrenza", tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy"));
                    pdfFormFields.SetField("DataEstinzione", tacticalRow.ExpirationDate.Value.ToString("dd/MM/yyyy"));
                    pdfFormFields.SetField("DurataMesi", GetMesi(durataGiorni).ToString());
                    pdfFormFields.SetField("DurataGodutaMesi", GetMesi(durataGoduta).ToString());
                    pdfFormFields.SetField("PremioNetto", WriteCurrencyString(premioDaRimborsare));
                    pdfFormFields.SetField("DurataGiorni", durataGiorni.ToString());
                    pdfFormFields.SetField("DurataGiorniGoduta", durataGoduta.ToString());
                    pdfFormFields.SetField("DurataResidua", (durataGiorni - durataGoduta).ToString());
                    pdfFormFields.SetField("ImportoRimborso", WriteCurrencyString(premioDaRimborsare));
                    pdfStamper.AcroFields.GenerateAppearances = true;
                    pdfStamper.FormFlattening = true;

                    pdfStamper.Close();
                }
            }
            return newFile;
        }

        public string CreaLetteraDisdettaGAP(string baseOutputFolder, TacticalRow tacticalRow, Product product, DateTime date)
        {
            var pdfTemplate = string.IsNullOrEmpty(Program.TEMPLATE_DISDETTA) ?
                                  GetFilePath(@"Modellini\" + "Disdetta.pdf")
                                : GetFilePath(@"Modellini\" + Program.TEMPLATE_DISDETTA);
            var newFile = GetLetterFilePath(baseOutputFolder, tacticalRow);

            int numberOfPages;
            using (var pdfReader = new PdfReader(pdfTemplate))
            {
                PdfReader.unethicalreading = true;
                numberOfPages = pdfReader.NumberOfPages;

                using (var pdfStamper = new PdfStamper(pdfReader, new FileStream(newFile, FileMode.Create)))
                {
                    var pdfFormFields = pdfStamper.AcroFields;

                    var premioDaRimborsare = GetPremioDaRimborsare(tacticalRow, product);
                    var durataGiorni = GetDurataGiorni(tacticalRow);
                    var durataGoduta = GetGiorniGoduti(tacticalRow);

                    pdfFormFields.SetField("NumeroAdesione", tacticalRow.PolicyNumber.ToString());
                    pdfFormFields.SetField("NomeCognome", (tacticalRow.Sex == 1 ? "Sig. " : "Sig.ra ") + tacticalRow.FirstName + " " + tacticalRow.Surname);
                    pdfFormFields.SetField("Indirizzo", tacticalRow.Street);
                    pdfFormFields.SetField("DettagliIndirizzo", tacticalRow.ZipCode + " - " + tacticalRow.Location + " - " + tacticalRow.Provincia);
                    pdfFormFields.SetField("Data", GetLetterDate(date));
                    pdfFormFields.SetField("Iban", tacticalRow.Iban);
                    pdfFormFields.SetField("DataDecorrenza", tacticalRow.DateOfSignature.Value.ToString("dd/MM/yyyy"));
                    pdfFormFields.SetField("DataEstinzione", tacticalRow.ExpirationDate.Value.ToString("dd/MM/yyyy"));
                    pdfFormFields.SetField("DurataMesi", GetMesi(durataGiorni).ToString());
                    pdfFormFields.SetField("DurataGodutaMesi", GetMesi(durataGoduta).ToString());
                    pdfFormFields.SetField("PremioNetto", WriteCurrencyString(premioDaRimborsare));
                    pdfFormFields.SetField("DurataGiorni", durataGiorni.ToString());
                    pdfFormFields.SetField("DurataGiorniGoduta", durataGoduta.ToString());
                    pdfFormFields.SetField("DurataResidua", (durataGiorni - durataGoduta).ToString());
                    pdfFormFields.SetField("ImportoRimborso", WriteCurrencyString(premioDaRimborsare));
                    pdfStamper.AcroFields.GenerateAppearances = true;
                    pdfStamper.FormFlattening = true;

                    pdfStamper.Close();
                }
            }
            return newFile;
        }

        private int GetMesi(int giorni)
        {
            return giorni / 30;
        }


        #endregion

        private string GetLetterFilePath(string baseOutputFolder, TacticalRow tacticalRow, string suffix = "")
        {
            var codiceHE = "NO_HE";
            if (!string.IsNullOrEmpty(tacticalRow.HE))
            {
                codiceHE = tacticalRow.HE;
            }

            return Path.Combine(baseOutputFolder, string.Format("{0}-{1}-{2} {3}{4}.pdf", codiceHE, "10017", tacticalRow.FirstName, tacticalRow.Surname.ToUpper(), suffix));
        }

        private void ImportHE_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file codici HE";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                StartImportCodiciHE(openFileDialog.FileName);
            }
        }

        private void SelectFolderForRenameOldLetters_Click(object sender, EventArgs e)
        {
            var dialog = new FolderBrowserDialog();
            if (dialog.ShowDialog(this) == DialogResult.OK)
            {
                OldLettersFolder.Text = dialog.SelectedPath;
            }
        }

        private void RenameOldLetters_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(OldLettersFolder.Text))
            {
                ShowError("Nessuna cartella di input selezionata");
                return;
            }

            if (!Directory.Exists(OldLettersFolder.Text))
            {
                ShowError("La cartella selezionata non esiste o non è accessibile");
                return;
            }

            try
            {
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                var logRows = new List<string>();
                using (var uow = new UnitOfWork())
                {
                    var baseOutputFolder = Path.Combine(OutputFolder.Text, "LettereRinominate");
                    if (Directory.Exists(baseOutputFolder))
                    {
                        Directory.Delete(baseOutputFolder, true);
                    }

                    System.Threading.Thread.Sleep(1000);

                    Directory.CreateDirectory(baseOutputFolder);

                    int row = 1;
                    foreach (var file in Directory.GetFiles(OldLettersFolder.Text, "*", SearchOption.TopDirectoryOnly))
                    {
                        var fileName = Path.GetFileNameWithoutExtension(file);
                        var tokens = fileName.Split(new char[] { '-' }, StringSplitOptions.RemoveEmptyEntries);

                        if (tokens.Length == 2)
                        {
                            int policyNumber;
                            if (!int.TryParse(tokens[0].Trim(), out policyNumber))
                            {
                                AddLogRow(logRows, "Numero polizza non corretto. Nome lettera " + fileName, row, "");
                            }
                            else
                            {
                                var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == policyNumber).FirstOrDefault();

                                if (tacticalRow == null)
                                {
                                    AddLogRow(logRows, "Polizza non trovata nel database. Nome lettera " + fileName, row, policyNumber.ToString());
                                }
                                else
                                {
                                    if (string.IsNullOrEmpty(tacticalRow.HE))
                                    {
                                        AddLogRow(logRows, "Codice HE non trovato per la polizza corrente. Nome lettera " + fileName, row, policyNumber.ToString());
                                    }
                                    else
                                    {
                                        File.Copy(file, GetLetterFilePath(baseOutputFolder, tacticalRow));
                                    }
                                }
                            }
                        }

                        row++;
                    }
                }

                WriteLog(logRows, "LogRinominaLettere");

                Cursor = Cursors.Default;
                ShowInfo("Migrazione lettere effettuato");
                Process.Start(OutputFolder.Text);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void ExportRows_Click(object sender, EventArgs e)
        {
            var policyNumbers = HelperClass.PromptString(this, "Inserisci numero polizze separati da ;", "Esporta polizze");

            if (!string.IsNullOrEmpty(policyNumbers))
            {
                try
                {
                    using (var excel = new ExcelPackage())
                    {
                        var ws = excel.Workbook.Worksheets.Add("Export");

                        using (var uow = new UnitOfWork())
                        {
                            var policies = policyNumbers.Split(new char[] { ';' }, StringSplitOptions.RemoveEmptyEntries);
                            var row = 1;
                            var col = 1;

                            ws.Cells[row, col++].Value = "Birth";
                            ws.Cells[row, col++].Value = "CalculatedAge";
                            ws.Cells[row, col++].Value = "CapitaleResiduo";
                            ws.Cells[row, col++].Value = "Commissioni";
                            ws.Cells[row, col++].Value = "CommissioniDanni";
                            ws.Cells[row, col++].Value = "CommissioniVita";
                            ws.Cells[row, col++].Value = "DataRimborso";
                            ws.Cells[row, col++].Value = "DateOfCancellation";
                            ws.Cells[row, col++].Value = "DateOfCancellation2";
                            ws.Cells[row, col++].Value = "DateOfCancellation3";
                            ws.Cells[row, col++].Value = "DateOfSignature";
                            ws.Cells[row, col++].Value = "DebitoResiduo";
                            ws.Cells[row, col++].Value = "ExpirationDate";
                            ws.Cells[row, col++].Value = "ExpirationDate2";
                            ws.Cells[row, col++].Value = "FirstName";
                            ws.Cells[row, col++].Value = "HE";
                            ws.Cells[row, col++].Value = "Iban";
                            ws.Cells[row, col++].Value = "Id";
                            ws.Cells[row, col++].Value = "ImportoEstinzione";
                            ws.Cells[row, col++].Value = "ImportoEstinzione2";
                            ws.Cells[row, col++].Value = "InsuredRate";
                            ws.Cells[row, col++].Value = "InsuredSum";
                            ws.Cells[row, col++].Value = "Location";
                            ws.Cells[row, col++].Value = "PolicyNumber";
                            ws.Cells[row, col++].Value = "PremioLordo";
                            ws.Cells[row, col++].Value = "PremioLordoDanni";
                            ws.Cells[row, col++].Value = "PremioNetto";
                            ws.Cells[row, col++].Value = "PremioNettoDanni";
                            ws.Cells[row, col++].Value = "PremioVita";
                            ws.Cells[row, col++].Value = "Provincia";
                            ws.Cells[row, col++].Value = "Sex";
                            ws.Cells[row, col++].Value = "Street";
                            ws.Cells[row, col++].Value = "Surname";
                            ws.Cells[row, col++].Value = "Tariff";
                            ws.Cells[row, col++].Value = "Tassi";
                            ws.Cells[row, col++].Value = "TassiDanni";
                            ws.Cells[row, col++].Value = "TransactionType";
                            ws.Cells[row, col++].Value = "ZipCode";
                            ws.Cells[row, col++].Value = "DebitoResiduoIniziale";

                            foreach (var policyNumber in policies)
                            {
                                row++;
                                col = 1;

                                var policyInt = 0;

                                if (!int.TryParse(policyNumber, out policyInt))
                                {
                                    throw new Exception("Errore nel formato del numero polizza: " + policyNumber);
                                }

                                var policy = uow.TacticalRowRepository.Get(t => t.PolicyNumber == policyInt).FirstOrDefault();

                                if (policy == null)
                                {
                                    throw new Exception("Polizza non trovata: " + policyNumber);
                                }

                                ws.Cells[row, col++].Value = policy.Birth;
                                ws.Cells[row, col++].Value = policy.CalculatedAge;
                                ws.Cells[row, col++].Value = policy.CapitaleResiduo;
                                ws.Cells[row, col++].Value = policy.Commissioni;
                                ws.Cells[row, col++].Value = policy.CommissioniDanni;
                                ws.Cells[row, col++].Value = policy.CommissioniVita;
                                ws.Cells[row, col++].Value = policy.DataRimborso;
                                ws.Cells[row, col++].Value = policy.DateOfCancellation;
                                ws.Cells[row, col++].Value = policy.DateOfCancellation2;
                                ws.Cells[row, col++].Value = policy.DateOfCancellation3;
                                ws.Cells[row, col++].Value = policy.DateOfSignature;
                                ws.Cells[row, col++].Value = policy.DebitoResiduo;
                                ws.Cells[row, col++].Value = policy.ExpirationDate;
                                ws.Cells[row, col++].Value = policy.ExpirationDate2;
                                ws.Cells[row, col++].Value = policy.FirstName;
                                ws.Cells[row, col++].Value = policy.HE;
                                ws.Cells[row, col++].Value = policy.Iban;
                                ws.Cells[row, col++].Value = policy.Id;
                                ws.Cells[row, col++].Value = policy.ImportoEstinzione;
                                ws.Cells[row, col++].Value = policy.ImportoEstinzione2;
                                ws.Cells[row, col++].Value = policy.InsuredRate;
                                ws.Cells[row, col++].Value = policy.InsuredSum;
                                ws.Cells[row, col++].Value = policy.Location;
                                ws.Cells[row, col++].Value = policy.PolicyNumber;
                                ws.Cells[row, col++].Value = policy.PremioLordo;
                                ws.Cells[row, col++].Value = policy.PremioLordoDanni;
                                ws.Cells[row, col++].Value = policy.PremioNetto;
                                ws.Cells[row, col++].Value = policy.PremioNettoDanni;
                                ws.Cells[row, col++].Value = policy.PremioVita;
                                ws.Cells[row, col++].Value = policy.Provincia;
                                ws.Cells[row, col++].Value = policy.Sex;
                                ws.Cells[row, col++].Value = policy.Street;
                                ws.Cells[row, col++].Value = policy.Surname;
                                ws.Cells[row, col++].Value = policy.Tariff;
                                ws.Cells[row, col++].Value = policy.Tassi;
                                ws.Cells[row, col++].Value = policy.TassiDanni;
                                ws.Cells[row, col++].Value = policy.TransactionType;
                                ws.Cells[row, col++].Value = policy.ZipCode;
                                ws.Cells[row, col++].Value = policy.DebitoResiduoIniziale;
                            }
                        }

                        excel.SaveAs(new FileInfo(GetCurrentExportName()));
                    }

                    ShowInfo("Esportazione eseguita");

                    Process.Start(OutputFolder.Text);
                }
                catch (Exception ex)
                {
                    ShowError(ex.Message);
                }
            }
        }

        private void groupBox3_Enter(object sender, EventArgs e)
        {

        }

        private void CheckChiusura_Click(object sender, EventArgs e)
        {
            try
            {
                var inputDate = HelperClass.PromptDate(this, "Inserisci la data massima delle scadenze da verificare", "Data scadenza");
                if (inputDate.HasValue)
                {
                    var oggi = new DateTime(inputDate.Value.Year, inputDate.Value.Month, inputDate.Value.Day, 23, 59, 59);
                    Cursor = Cursors.WaitCursor;
                    Application.DoEvents();

                    var logRows = new List<string>();
                    using (var uow = new UnitOfWork())
                    {
                        var tacticalRows = uow.TacticalRowRepository.Get(tr => (tr.TransactionType == "40" || tr.TransactionType == "NC") && tr.ExpirationDate < oggi && !tr.ExpirationDate2.HasValue).ToList();
                        tacticalRows.AddRange(uow.TacticalRowRepository.Get(tr => (tr.TransactionType == "40" || tr.TransactionType == "NC") && tr.ExpirationDate2.HasValue && tr.ExpirationDate2.Value < oggi).ToList());

                        foreach (var tacticalRow in tacticalRows)
                        {
                            tacticalRow.TransactionType = HelperClass.TRANSATION_TYPE_CONTRATTO_CHIUSO;
                            uow.TacticalRowRepository.Update(tacticalRow);
                            uow.Save();

                            var dataScadenza = tacticalRow.ExpirationDate2.HasValue ? tacticalRow.ExpirationDate2.Value : tacticalRow.ExpirationDate.Value;

                            logRows.Add(string.Format("Pratica chiusa: {0} - Data scadenza: {1}", tacticalRow.PolicyNumber, dataScadenza.ToString("dd/MM/yyyy")));
                        }
                    }

                    WriteLog(logRows, "LogChiusuraPratiche");

                    Cursor = Cursors.Default;
                    ShowInfo("Verifiche eseguito con successo");
                    Process.Start(OutputFolder.Text);
                }
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void ExportMonthlyAML_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file lavorativo";
            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            var logRows = new List<string>();

            try
            {
                if (!Directory.Exists(OutputFolder.Text))
                {
                    Directory.CreateDirectory(OutputFolder.Text);
                }

                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                var modellino = new FileInfo(GetFilePath("Modellini\\AML-ControlFileDescription.xlsx"));

                using (var monthly = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                {
                    var wsMonthly = monthly.Workbook.Worksheets[1];

                    using (var targetPackage = new ExcelPackage(modellino))
                    {
                        var wsTarget = targetPackage.Workbook.Worksheets[1];

                        var inputRow = 2;
                        var outputRow = 2;

                        var exit = false;


                        using (var uow = new UnitOfWork())
                        {
                            while (!exit)
                            {
                                if (wsMonthly.Cells[inputRow, 1].Value == null)
                                {
                                    exit = true;
                                }
                                else
                                {
                                    try
                                    {
                                        var pratica = int.Parse(wsMonthly.Cells["F" + inputRow.ToString()].Value.ToString());
                                        var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == pratica).First();

                                        var cognome = "";
                                        var nome = "";
                                        var ragioneSociale = "";
                                        var dataDiNascita = "";

                                        if (!string.IsNullOrEmpty(tacticalRow.NominativoAderente))
                                        {
                                            if (!string.IsNullOrEmpty(tacticalRow.CodiceFiscaleAderente) && tacticalRow.CodiceFiscaleAderente.Length == 16)
                                            {
                                                cognome = tacticalRow.NominativoAderente;
                                                dataDiNascita = GetDateFromFiscalCode(tacticalRow.CodiceFiscaleAderente).ToString("yyyy/MM/dd");
                                            }
                                            else
                                            {
                                                ragioneSociale = tacticalRow.NominativoAderente;
                                            }
                                        }
                                        else
                                        {
                                            if (string.IsNullOrEmpty(tacticalRow.FirstName))
                                            {
                                                ragioneSociale = tacticalRow.Surname;
                                            }
                                            else
                                            {
                                                cognome = tacticalRow.Surname;
                                                nome = tacticalRow.FirstName;
                                                dataDiNascita = tacticalRow.Birth.ToString("yyyy/MM/dd");
                                            }
                                        }

                                        wsTarget.Cells["A" + outputRow].Value = cognome;
                                        wsTarget.Cells["B" + outputRow].Value = ragioneSociale;
                                        wsTarget.Cells["C" + outputRow].Value = nome;
                                        wsTarget.Cells["E" + outputRow].Value = dataDiNascita;

                                        outputRow++;

                                        if (!string.IsNullOrEmpty(tacticalRow.CognomeCoIntestario))
                                        {
                                            wsTarget.Cells["A" + outputRow].Value = tacticalRow.CognomeCoIntestario;
                                            wsTarget.Cells["B" + outputRow].Value = "";
                                            wsTarget.Cells["C" + outputRow].Value = tacticalRow.NomeCoIntestatario;
                                            wsTarget.Cells["E" + outputRow].Value = tacticalRow.CointestatarioDataNascita.HasValue ? tacticalRow.CointestatarioDataNascita.Value.ToString("yyyy/MM/dd") : "";

                                            outputRow++;
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        logRows.Add(string.Format("Errore in riga monthly {0}: {1}", inputRow, ex.Message));
                                    }

                                    inputRow += 2;
                                }
                            }
                        }

                        targetPackage.SaveAs(new FileInfo(GetCurrentAmlName(openFileDialog.FileName)));
                    }
                }

                Cursor = Cursors.Default;

                ShowInfo("File esportato");

                Process.Start(OutputFolder.Text);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
            finally
            {
                WriteLog(logRows, "EsportazioneBanca");
            }
        }

        private void OpenSearch_Click(object sender, EventArgs e)
        {
            var form = new Search();
            form.ShowDialog(this);
        }

        private void ImportazioneMonthlyTotale_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file MONTHLY_TOTALE";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var targetPath = GetMonthlyTotalePath();

                    using (var excelPackage = new ExcelPackage(new FileInfo(openFileDialog.FileName)))
                    {
                        var ws = excelPackage.Workbook.Worksheets[1];

                        var row = 2;
                        var exit = false;

                        while (!exit)
                        {
                            if (ws.Cells[row, 2].Value == null)
                            {
                                exit = true;
                            }
                            else
                            {
                                RaiseExportStepAdvanced("Importazione riga", row, 0);

                                var convenzione = GetString(ws.Cells["A" + row].Value);
                                var codicePratica = int.Parse(GetString(ws.Cells["F" + row].Value));
                                var premioAccantonato = GetString(ws.Cells["N" + row].Value);
                                var premioRimborsato = GetString(ws.Cells["O" + row].Value);
                                var dataCancellazione = GetDateNullable(ws.Cells["H" + row].Value);
                                var dataAccantomento = GetDateNullable(ws.Cells["Q" + row].Value);
                                var isDanni = !convenzione.Contains("_28");

                                if (!dataCancellazione.HasValue)
                                {
                                    throw new Exception("Errore data cancellazione in riga: " + row.ToString());
                                }

                                using (var uow = new UnitOfWork())
                                {
                                    var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == codicePratica).FirstOrDefault();
                                    if (tacticalRow != null)
                                    {
                                        var cancellation = uow.CancellationRepository.Get(c => c.TacticalRowId == tacticalRow.Id && c.Date == dataCancellazione.Value).FirstOrDefault();
                                        if (cancellation == null)
                                        {
                                            cancellation = new Cancellation();
                                            cancellation.TacticalRowId = tacticalRow.Id;
                                            cancellation.Date = dataCancellazione.Value;
                                        }

                                        if (isDanni)
                                        {
                                            cancellation.DataAccontonatoDanni = dataAccantomento;
                                            cancellation.PremioAccantonatoDanni = premioAccantonato.Trim().ToUpper() == "S";
                                            cancellation.PremioRimborsatoDanni = premioRimborsato.Trim().ToUpper() == "S";
                                            cancellation.DataDistintaDanni = GetDateNullable(ws.Cells["U" + row].Value);
                                            cancellation.DataValutaDanni = GetDateNullable(ws.Cells["V" + row].Value);
                                        }
                                        else
                                        {
                                            cancellation.DataAccantonatoVita = dataAccantomento;
                                            cancellation.PremioAccantonatoVita = premioAccantonato.Trim().ToUpper() == "S";
                                            cancellation.PremioRimborsatoVita = premioRimborsato.Trim().ToUpper() == "S";
                                            cancellation.DataDistintaVita = GetDateNullable(ws.Cells["S" + row].Value);
                                            cancellation.DataValutaVita = GetDateNullable(ws.Cells["T" + row].Value);
                                        }

                                        if (cancellation.Id == 0)
                                        {
                                            uow.CancellationRepository.Insert(cancellation);
                                        }
                                        else
                                        {
                                            uow.CancellationRepository.Update(cancellation);
                                        }

                                        uow.Save();
                                    }

                                    row++;
                                }
                            }
                        }

                        File.Copy(openFileDialog.FileName, targetPath, true);

                        ExportProgress.Text = "";
                        ShowInfo("Importazione completata");
                    }
                }
                catch (Exception ex)
                {
                    ShowError(ex.Message);
                }
            }
        }

        private string GetMonthlyTotalePath()
        {
            return GetFilePath(@"Modellini\MONTHLY_TOTALE.xlsx");
        }

        private string GetMonthlyTotaleExportPath()
        {
            return Path.Combine(OutputFolder.Text, "MONTHLY_TOTALE.xlsx");
        }

        private void EsportazioneMonthlyTotale_Click(object sender, EventArgs e)
        {
            if (MessageBox.Show(this, "Avviare l'esportazione del MONTHLY_TOTALE ?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
            {
                var monthlyTotalePath = GetMonthlyTotalePath();

                if (!File.Exists(monthlyTotalePath))
                {
                    ShowError("Non hai ancora importato il MONTHLY_TOTALE la prima volta");
                    return;
                }

                try
                {
                    using (var excelPackage = new ExcelPackage(new FileInfo(monthlyTotalePath)))
                    {
                        var ws = excelPackage.Workbook.Worksheets[1];

                        var row = 2;
                        var exit = false;

                        using (var uow = new UnitOfWork())
                        {
                            while (!exit)
                            {
                                if (ws.Cells[row, 2].Value == null)
                                {
                                    exit = true;
                                }
                                else
                                {
                                    RaiseExportStepAdvanced("Esportazione riga", row, 0);

                                    try
                                    {
                                        var convenzione = GetString(ws.Cells["A" + row].Value);
                                        var codicePratica = int.Parse(GetString(ws.Cells["F" + row].Value));
                                        var isDanni = !convenzione.Contains("_28");

                                        var tacticalRow = uow.TacticalRowRepository.Get(tr => tr.PolicyNumber == codicePratica).FirstOrDefault();
                                        if (tacticalRow != null)
                                        {
                                            ScriviValoriRimborsatoAccantonatoNelMonthly(ws, row, isDanni, tacticalRow.Id, uow);
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        throw ex;
                                    }

                                    row++;
                                }
                            }
                        }

                        excelPackage.SaveAs(new FileInfo(GetMonthlyTotaleExportPath()));

                        ExportProgress.Text = "";
                        ShowInfo("Esportazione completata");
                        OpenOutput_Click(null, null);
                    }
                }
                catch (Exception ex)
                {
                    ShowError(ex.Message);
                }
            }
        }

        private void ScriviValoriRimborsatoAccantonatoNelMonthly(ExcelWorksheet ws, int row, bool isDanni, int tacticalRowId, UnitOfWork uow)
        {
            var dataCancellazione = GetDateNullable(ws.Cells["H" + row].Value);
            var cancellation = uow.CancellationRepository.Get(c => c.TacticalRowId == tacticalRowId && c.Date == dataCancellazione).FirstOrDefault();

            if (cancellation != null)
            {
                if (isDanni)
                {
                    if (cancellation.PremioAccantonatoDanni.HasValue)
                    {
                        ws.Cells["N" + row].Value = cancellation.PremioAccantonatoDanni.Value ? "S" : "N";
                    }
                    if (cancellation.PremioRimborsatoDanni.HasValue)
                    {
                        ws.Cells["O" + row].Value = cancellation.PremioRimborsatoDanni.Value ? "S" : "N";
                    }

                    ws.Cells["Q" + row].Value = cancellation.DataAccontonatoDanni;
                    ws.Cells["U" + row].Value = GetNormalDateString(cancellation.DataDistintaDanni);
                    ws.Cells["V" + row].Value = GetNormalDateString(cancellation.DataValutaDanni);
                }
                else
                {
                    if (cancellation.PremioAccantonatoVita.HasValue)
                    {
                        ws.Cells["N" + row].Value = cancellation.PremioAccantonatoVita.Value ? "S" : "N";
                    }
                    if (cancellation.PremioRimborsatoVita.HasValue)
                    {
                        ws.Cells["O" + row].Value = cancellation.PremioRimborsatoVita.Value ? "S" : "N";
                    }

                    ws.Cells["Q" + row].Value = cancellation.DataAccantonatoVita;
                    ws.Cells["S" + row].Value = GetNormalDateString(cancellation.DataDistintaVita);
                    ws.Cells["T" + row].Value = GetNormalDateString(cancellation.DataValutaVita);
                }
            }
        }

        private void ImpostazioneMonthlyTotale_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file MONTHLY_TOTALE";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                if (MessageBox.Show(this, "Sicuro di voler sostituire il MONTHLY TOTALE con questo nuovo file ?", "Conferma", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                {
                    File.Copy(openFileDialog.FileName, GetMonthlyTotalePath(), true);

                    ShowInfo("Nuovo MONTHLY TOTALE caricato");
                }
            }
        }

        private void OpenDistinte_Click(object sender, EventArgs e)
        {
            var bankTransactionSearch = new SearchBankTransactions();
            bankTransactionSearch.ShowDialog(this);
        }

        private void ExportRecessi_Click(object sender, EventArgs e)
        {
            EsportaMonthly(true);
        }

        private void EsportaMonthly(bool recessiDisdette)
        {
            try
            {
                if (string.IsNullOrEmpty(OutputFolder.Text))
                {
                    ShowWarning("Selezionare una cartella di output");
                    return;
                }

                if (MonthList.SelectedIndex == -1)
                {
                    ShowWarning("Selezionare un mese");
                    return;
                }

                int year;
                if (!int.TryParse(ExportYear.Text, out year))
                {
                    ShowWarning("Inserire un anno valido");
                    return;
                }

                if (!Directory.Exists(OutputFolder.Text))
                {
                    Directory.CreateDirectory(OutputFolder.Text);
                }

                Cursor = Cursors.WaitCursor;
                ExportProgress.Text = "Recupero informazioni in corso...";
                Application.DoEvents();

                using (var excelPackage = new ExcelPackage())
                {
                    var ws = excelPackage.Workbook.Worksheets.Add("Monthly Cancellation File");

                    var row = 1;
                    ws.Cells["A" + row.ToString()].Value = "Convenzione";
                    ws.Cells["B" + row.ToString()].Value = "Prodotto";
                    ws.Cells["C" + row.ToString()].Value = "Cognome";
                    ws.Cells["D" + row.ToString()].Value = "Nome";
                    ws.Cells["E" + row.ToString()].Value = "Data di nascita";
                    ws.Cells["F" + row.ToString()].Value = "Codice contratto";
                    ws.Cells["G" + row.ToString()].Value = "Tipologia di cancellazione";
                    ws.Cells["H" + row.ToString()].Value = "Data di cancellazione";
                    ws.Cells["I" + row.ToString()].Value = "Data di decorrenza contratto";
                    ws.Cells["J" + row.ToString()].Value = "ET Fees CNPSI";
                    ws.Cells["K" + row.ToString()].Value = "ET Fees SCB ";
                    ws.Cells["L" + row.ToString()].Value = "Importo Rimborso";
                    ws.Cells["M" + row.ToString()].Value = "Commissioni";
                    ws.Cells["N" + row.ToString()].Value = "Premio Accantonato (S/N)";
                    ws.Cells["O" + row.ToString()].Value = "Premio Rimborsato al Cliente (S/N)";
                    ws.Cells["P" + row.ToString()].Value = "Data ricezione documentazione";
                    ws.Cells["Q" + row.ToString()].Value = "Data Accantonamento";
                    ws.Cells["R" + row.ToString()].Value = "Codice Tariffa";

                    ws.Column(12).Style.Numberformat.Format = "#0.00";
                    ws.Column(13).Style.Numberformat.Format = "#0.00";

                    ws.Column(16).Style.Numberformat.Format = "dd/MM/yyyy";
                    ws.Column(17).Style.Numberformat.Format = "dd/MM/yyyy";

                    row++;

                    var logRows = new List<string>();

                    EsportaEstinzioni(ws, MonthList.SelectedIndex + 1, year, recessiDisdette, OutputFolder.Text, ref row, logRows);
                    if (!recessiDisdette)
                    {
                        EsportaRevoche(ws, MonthList.SelectedIndex + 1, year, OutputFolder.Text, ref row, logRows);
                    }
                    EsportaRecessi(ws, MonthList.SelectedIndex + 1, year, recessiDisdette, OutputFolder.Text, ref row, logRows);

                    excelPackage.SaveAs(new FileInfo(GetCurrentMonthlyName(recessiDisdette)));

                    WriteLog(logRows, "ExportLog");
                }

                Cursor = Cursors.Default;
                ExportProgress.Text = "Esportazione completata";

                ShowInfo("File esportati");

                Process.Start(OutputFolder.Text);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }

        private void RecessiBanca_Click(object sender, EventArgs e)
        {
            try
            {
                if (string.IsNullOrEmpty(OutputFolder.Text))
                {
                    ShowWarning("Selezionare una cartella di output");
                    return;
                }

                if (MonthList.SelectedIndex == -1)
                {
                    ShowWarning("Selezionare un mese");
                    return;
                }

                int year;
                if (!int.TryParse(ExportYear.Text, out year))
                {
                    ShowWarning("Inserire un anno valido");
                    return;
                }

                if (!Directory.Exists(OutputFolder.Text))
                {
                    Directory.CreateDirectory(OutputFolder.Text);
                }

                Cursor = Cursors.WaitCursor;
                ExportProgress.Text = "Recupero informazioni in corso...";
                Application.DoEvents();

                using (var excelPackage = new ExcelPackage())
                {
                    var ws = excelPackage.Workbook.Worksheets.Add("Monthly Cancellation File");

                    var row = 1;
                    ws.Cells["A" + row.ToString()].Value = "Cognome";
                    ws.Cells["B" + row.ToString()].Value = "Nome";
                    ws.Cells["C" + row.ToString()].Value = "Pratica";
                    ws.Cells["D" + row.ToString()].Value = "Data Cancellazione";
                    ws.Cells["E" + row.ToString()].Value = "Gestione CBP";
                    ws.Cells["F" + row.ToString()].Value = "Disdetta";

                    ws.Column(4).Style.Numberformat.Format = "dd/MM/yyyy";

                    row++;

                    var logRows = new List<string>();
                    var dtFrom = new DateTime(year, MonthList.SelectedIndex + 1, 1);

                    using (var uow = new UnitOfWork())
                    {
                        var query = uow.Context.Set<TacticalRow>().AsQueryable();

                        query = query.Where(tr => tr.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO
                                                || tr.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA);

                        var dtTo = dtFrom.AddMonths(1).AddSeconds(-1);

                        query = query.Where(tr => tr.DateOfCancellation >= dtFrom && tr.DateOfCancellation <= dtTo);


                        foreach (var tacticalRow in query.ToList())
                        {
                            ws.Cells["A" + row.ToString()].Value = tacticalRow.Surname;
                            ws.Cells["B" + row.ToString()].Value = tacticalRow.FirstName;
                            ws.Cells["C" + row.ToString()].Value = tacticalRow.PolicyNumber;
                            ws.Cells["D" + row.ToString()].Value = tacticalRow.DateOfCancellation;
                            ws.Cells["E" + row.ToString()].Value = tacticalRow.Santander ? "NO" : "SI";
                            ws.Cells["F" + row.ToString()].Value = tacticalRow.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA ? "X" : "";

                            row++;
                        }

                    }

                    excelPackage.SaveAs(new FileInfo(GetCurrentRecessiBancaName()));

                    WriteLog(logRows, "ExportLog");
                }

                Cursor = Cursors.Default;
                ExportProgress.Text = "Esportazione completata";

                ShowInfo("File esportati");

                Process.Start(OutputFolder.Text);
            }
            catch (Exception ex)
            {
                Cursor = Cursors.Default;
                ShowError(ex.Message);
            }
        }
    }
}
