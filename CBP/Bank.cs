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
using System.Data.SqlClient;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace CBP
{
    public partial class Bank : Form
    {
        private double MAX_VALUE = 100000D;
        private bool UfficioPaghe { get; set; }

        public Bank()
        {
            InitializeComponent();

            DataBonifico.Value = DateTime.Now;

            UfficioPaghe = (ConfigurationManager.AppSettings["UfficioPaghe"] ?? "") == "true";
            OutputFolder.Text = ConfigurationManager.AppSettings["OutputFolder"] ?? "";
            if (string.IsNullOrEmpty(OutputFolder.Text))
            {
                OutputFolder.Text = @"C:\Develop\CBP\Dati\Output";
            }

            CreaExcelDaReportBase.Visible = UfficioPaghe;
        }

        private void ExportMonthly_Click(object sender, EventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.CheckFileExists = true;
            openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
            openFileDialog.Title = "Seleziona il file input banca";
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                CreaXmlFile(openFileDialog.FileName);
            }
        }

        private void CreaXmlFile(string fileName)
        {
            var logRows = new List<string>();

            try
            {
                Cursor = Cursors.WaitCursor;
                Application.DoEvents();

                List<PaymentInfo> pagamentiVita = null, pagamentiDanni = null, pagamentiDanni2 = null;

                using (var bankInput = new ExcelPackage(new FileInfo(fileName)))
                {
                    var ws = bankInput.Workbook.Worksheets[1];

                    var payments = GetPayments(ws, logRows);

                    pagamentiVita = payments.Where(p => p.Tipologia == 0).ToList();
                    pagamentiDanni = payments.Where(p => p.Tipologia == 1).ToList();
                    pagamentiDanni2 = payments.Where(p => p.Tipologia == 2).ToList();
                }

                List<List<PaymentInfo>> splitted = null;
                var index = 1;
                var executionIndex = 0;

                if (pagamentiVita.Count > 0)
                {
                    splitted = SplitPayments(pagamentiVita);
                    index = 1;
                    foreach (var payments in splitted)
                    {
                        Application.DoEvents();

                        var nomeDistinta = CreaXmlFile(payments, fileName, 0, index++, executionIndex++);
                        HelperClass.CreaDistinta(payments, false, nomeDistinta);
                    }
                }

                if (pagamentiDanni.Count > 0)
                {
                    splitted = SplitPayments(pagamentiDanni);
                    index = 1;
                    foreach (var payments in splitted)
                    {
                        Application.DoEvents();

                        var nomeDistinta = CreaXmlFile(payments, fileName, 1, index++, executionIndex++);
                        HelperClass.CreaDistinta(payments, true, nomeDistinta);
                    }
                }

                if (pagamentiDanni2.Count > 0)
                {
                    splitted = SplitPayments(pagamentiDanni2);
                    index = 1;
                    foreach (var payments in splitted)
                    {
                        Application.DoEvents();

                        var nomeDistinta = CreaXmlFile(payments, fileName, 2, index++, executionIndex++);
                        HelperClass.CreaDistinta(payments, true, nomeDistinta);
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

        private List<List<PaymentInfo>> SplitPayments(List<PaymentInfo> payments)
        {
            var retVal = new List<List<PaymentInfo>>();

            var lastTotal = 0D;

            var lastList = new List<PaymentInfo>();
            retVal.Add(lastList);

            foreach (var payment in payments)
            {
                if (payment.Value + lastTotal > MAX_VALUE)
                {
                    lastList = new List<PaymentInfo>();
                    lastList.Add(payment);
                    retVal.Add(lastList);
                    lastTotal = payment.Value;
                }
                else
                {
                    lastTotal += payment.Value;
                    lastList.Add(payment);
                }
            }

            return retVal;
        }

        private string CreaXmlFile(List<PaymentInfo> payments, string fileName, int tipologia, int index, int executionIndex)
        {
            var xmlDoc = new XmlDocument();

            var instrId = 1;
            var endToEndId = 1;
            var nomeCC = "";
            var codiceCUC = "";
            var codiceIban = "";

            xmlDoc.AppendChild(xmlDoc.CreateXmlDeclaration("1.0", "UTF-8", null));

            var root = xmlDoc.AppendChild(xmlDoc.CreateElement("CBIPaymentRequest"));

            var attr = xmlDoc.CreateAttribute("xmlns");
            attr.InnerText = "urn:CBI:xsd:CBIPaymentRequest.00.04.00";
            root.Attributes.Append(attr);

            attr = xmlDoc.CreateAttribute("xmlns:xsi");
            attr.InnerText = "http://www.w3.org/2001/XMLSchema-instance";
            root.Attributes.Append(attr);

            attr = xmlDoc.CreateAttribute("xsi:schemaLocation");
            attr.InnerText = "urn:CBI:xsd:CBIPaymentRequest.00.04.00 file:/CBIPaymentRequest.00.04.00.xsd";
            root.Attributes.Append(attr);

            var grpHdr = xmlDoc.CreateElement("GrpHdr");
            root.AppendChild(grpHdr);

            var pmtInf = xmlDoc.CreateElement("PmtInf");
            root.AppendChild(pmtInf);

            var msgIdValue = DateTime.Now.ToString("yyyyMMddHHmmssfff") + executionIndex.ToString().PadLeft(2, '0');

            var msgId = xmlDoc.CreateElement("MsgId");
            msgId.InnerText = msgIdValue;
            grpHdr.AppendChild(msgId);

            var creDtTm = xmlDoc.CreateElement("CreDtTm");

            creDtTm.InnerText = DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss");
            grpHdr.AppendChild(creDtTm);

            var nbOfTxs = xmlDoc.CreateElement("NbOfTxs");
            nbOfTxs.InnerText = payments.Count().ToString();
            grpHdr.AppendChild(nbOfTxs);

            var ctrlSum = xmlDoc.CreateElement("CtrlSum");
            ctrlSum.InnerText = WriteDoubleInXml(Math.Round(payments.Sum(p => p.Value), 2));
            grpHdr.AppendChild(ctrlSum);

            var initgPty = xmlDoc.CreateElement("InitgPty");
            grpHdr.AppendChild(initgPty);

            var nm = xmlDoc.CreateElement("Nm");
            initgPty.AppendChild(nm);

            var id = xmlDoc.CreateElement("Id");
            initgPty.AppendChild(id);

            var orgId = xmlDoc.CreateElement("OrgId");
            id.AppendChild(orgId);

            var othr = xmlDoc.CreateElement("Othr");
            orgId.AppendChild(othr);

            var idFinal = xmlDoc.CreateElement("Id");
            othr.AppendChild(idFinal);

            var issr = xmlDoc.CreateElement("Issr");
            issr.InnerText = "CBI";
            othr.AppendChild(issr);

            if (tipologia == 0)
            {
                nomeCC = "CNP SANTANDER INSURANCE LIFE DAC";
                codiceCUC = "XXXBF708";
                codiceIban = "IT19I0100501634000000000372";
            }
            else if (tipologia == 1)
            {
                nomeCC = "CNP SANTANDER INSURANCE EUROPE DAC";
                codiceCUC = "XXXBF70V";
                codiceIban = "IT65G0100501634000000000370";
            }
            else if (tipologia == 2)
            {
                nomeCC = "CNP SANTANDER INSURANCE EUROPE DAC";
                codiceCUC = "XXXBF70V";
                codiceIban = "IT94H0100501634000000007777";
            }
            else
            {
                throw new Exception("Tipologia non valida");
            }

            nm.InnerText = nomeCC;
            idFinal.InnerText = codiceCUC;

            var pmtInfId = xmlDoc.CreateElement("PmtInfId");
            pmtInfId.InnerText = msgIdValue;
            pmtInf.AppendChild(pmtInfId);

            var pmtMtd = xmlDoc.CreateElement("PmtMtd");
            pmtMtd.InnerText = "TRA";
            pmtInf.AppendChild(pmtMtd);

            var btchBookg = xmlDoc.CreateElement("BtchBookg");
            btchBookg.InnerText = "true";
            pmtInf.AppendChild(btchBookg);

            var pmtTpInf = xmlDoc.CreateElement("PmtTpInf");
            pmtInf.AppendChild(pmtTpInf);

            var instrPrty = xmlDoc.CreateElement("InstrPrty");
            instrPrty.InnerText = "NORM";
            pmtTpInf.AppendChild(instrPrty);

            var svcLvl = xmlDoc.CreateElement("SvcLvl");
            pmtTpInf.AppendChild(svcLvl);

            var cd = xmlDoc.CreateElement("Cd");
            cd.InnerText = "SEPA";
            svcLvl.AppendChild(cd);

            var reqdExctnDt = xmlDoc.CreateElement("ReqdExctnDt");
            reqdExctnDt.InnerText = DataBonifico.Value.ToString("yyyy-MM-dd");
            pmtInf.AppendChild(reqdExctnDt);

            var dbtr = xmlDoc.CreateElement("Dbtr");
            pmtInf.AppendChild(dbtr);

            var nmDbtr = xmlDoc.CreateElement("Nm");
            nmDbtr.InnerText = nomeCC;
            dbtr.AppendChild(nmDbtr);

            var pstlAdrDbtr = xmlDoc.CreateElement("PstlAdr");
            dbtr.AppendChild(pstlAdrDbtr);

            var ctryDbtr = xmlDoc.CreateElement("Ctry");
            ctryDbtr.InnerText = "IT";
            pstlAdrDbtr.AppendChild(ctryDbtr);

            var adrLineDbtr = xmlDoc.CreateElement("AdrLine");
            adrLineDbtr.InnerText = "VIA";
            pstlAdrDbtr.AppendChild(adrLineDbtr);

            var idDbtr = xmlDoc.CreateElement("Id");
            dbtr.AppendChild(idDbtr);

            var orgIdDbtr = xmlDoc.CreateElement("OrgId");
            idDbtr.AppendChild(orgIdDbtr);

            var othrDbtr = xmlDoc.CreateElement("Othr");
            orgIdDbtr.AppendChild(othrDbtr);

            var idOrgDbtr = xmlDoc.CreateElement("Id");
            idOrgDbtr.InnerText = codiceCUC;
            othrDbtr.AppendChild(idOrgDbtr);

            var dbtrAcct = xmlDoc.CreateElement("DbtrAcct");
            pmtInf.AppendChild(dbtrAcct);

            var idDbtrAcct = xmlDoc.CreateElement("Id");
            dbtrAcct.AppendChild(idDbtrAcct);

            var ibanDbtrAcct = xmlDoc.CreateElement("IBAN");
            ibanDbtrAcct.InnerText = codiceIban;
            idDbtrAcct.AppendChild(ibanDbtrAcct);

            var dbtrAgt = xmlDoc.CreateElement("DbtrAgt");
            pmtInf.AppendChild(dbtrAgt);

            var finInstnId = xmlDoc.CreateElement("FinInstnId");
            dbtrAgt.AppendChild(finInstnId);

            var bic = xmlDoc.CreateElement("BIC");
            bic.InnerText = "BNLIITRR";
            finInstnId.AppendChild(bic);

            var clrSysMmbId = xmlDoc.CreateElement("ClrSysMmbId");
            finInstnId.AppendChild(clrSysMmbId);

            var mmbId = xmlDoc.CreateElement("MmbId");
            mmbId.InnerText = "01005";
            clrSysMmbId.AppendChild(mmbId);

            var chrgBr = xmlDoc.CreateElement("ChrgBr");
            chrgBr.InnerText = "SLEV";
            pmtInf.AppendChild(chrgBr);

            foreach (var payment in payments)
            {
                var cdtTrfTxInf = xmlDoc.CreateElement("CdtTrfTxInf");
                pmtInf.AppendChild(cdtTrfTxInf);

                var pmtId = xmlDoc.CreateElement("PmtId");
                cdtTrfTxInf.AppendChild(pmtId);

                var instrIdNode = xmlDoc.CreateElement("InstrId");
                instrIdNode.InnerText = (instrId++).ToString();
                pmtId.AppendChild(instrIdNode);

                var endToEndIdNode = xmlDoc.CreateElement("EndToEndId");
                endToEndIdNode.InnerText = msgIdValue + (endToEndId++).ToString().PadLeft(5, '0');
                pmtId.AppendChild(endToEndIdNode);

                var pmtTpInfCdtTrfTxInf = xmlDoc.CreateElement("PmtTpInf");
                cdtTrfTxInf.AppendChild(pmtTpInfCdtTrfTxInf);

                var ctgyPurp = xmlDoc.CreateElement("CtgyPurp");
                pmtTpInfCdtTrfTxInf.AppendChild(ctgyPurp);

                var cdCtgyPurp = xmlDoc.CreateElement("Cd");
                cdCtgyPurp.InnerText = "OTHR";
                ctgyPurp.AppendChild(cdCtgyPurp);

                var amt = xmlDoc.CreateElement("Amt");
                cdtTrfTxInf.AppendChild(amt);

                var instdAmt = xmlDoc.CreateElement("InstdAmt");

                attr = xmlDoc.CreateAttribute("Ccy");
                attr.InnerText = "EUR";
                instdAmt.Attributes.Append(attr);

                instdAmt.InnerText = WriteDoubleInXml(payment.Value);
                amt.AppendChild(instdAmt);

                var cdtr = xmlDoc.CreateElement("Cdtr");
                cdtTrfTxInf.AppendChild(cdtr);

                var nmCdtr = xmlDoc.CreateElement("Nm");
                nmCdtr.InnerText = payment.Nome + " " + payment.Cognome;
                cdtr.AppendChild(nmCdtr);

                var cdtrAcct = xmlDoc.CreateElement("CdtrAcct");
                cdtTrfTxInf.AppendChild(cdtrAcct);

                var idCdtrAcct = xmlDoc.CreateElement("Id");
                cdtrAcct.AppendChild(idCdtrAcct);

                var ibanNode = xmlDoc.CreateElement("IBAN");
                ibanNode.InnerText = payment.Iban;
                idCdtrAcct.AppendChild(ibanNode);

                var rmtInf = xmlDoc.CreateElement("RmtInf");
                cdtTrfTxInf.AppendChild(rmtInf);

                var ustrd = xmlDoc.CreateElement("Ustrd");
                ustrd.InnerText = payment.Causale;
                rmtInf.AppendChild(ustrd);
            }

            var percorsoFileDistinta = GetCurrentBankName(fileName, tipologia, index);
            xmlDoc.Save(percorsoFileDistinta);

            return Path.GetFileName(percorsoFileDistinta);
        }

        private string WriteDoubleInXml(double val)
        {
            return val.ToString().Replace(",", ".");
        }

        private List<PaymentInfo> GetPayments(ExcelWorksheet ws, List<string> logRows)
        {
            var retVal = new List<PaymentInfo>();
            var row = 2;
            var exit = false;

            while (!exit)
            {
                if (ws.Cells[row, 1].Value == null)
                {
                    exit = true;
                }
                else
                {
                    try
                    {
                        var info = new PaymentInfo();
                        info.Causale = GetString(ws.Cells["F" + row.ToString()].Value);
                        info.Cognome = GetString(ws.Cells["B" + row.ToString()].Value);
                        info.Iban = GetString(ws.Cells["D" + row.ToString()].Value);
                        info.Nome = GetString(ws.Cells["C" + row.ToString()].Value);
                        info.Tipologia = GetInt(ws.Cells["A" + row.ToString()].Value);
                        info.Value = GetDouble(ws.Cells["E" + row.ToString()].Value);
                        info.TacticalRowId = GetInt(ws.Cells["G" + row.ToString()].Value);
                        info.PagamentoAccantonato = GetString(ws.Cells["H" + row.ToString()].Value).ToUpper() == "SI";
                        info.DataCancellazione = HelperClass.GetDateNullable(ws.Cells["I" + row.ToString()].Value);

                        if (string.IsNullOrEmpty(info.Iban))
                        {
                            logRows.Add(string.Format("Errore in riga {0}. IBAN mancante", row));
                        }
                        else
                        {
                            retVal.Add(info);
                        }
                    }
                    catch (Exception ex)
                    {
                        logRows.Add(string.Format("Errore in riga {0}: {1}", row, ex.Message));
                    }

                    row++;
                }
            }

            return retVal;
        }

        private int? GetIntNullable(object value)
        {
            if (value == null)
            {
                return null;
            }

            return Convert.ToInt32(value);
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

        private double GetDouble(object value)
        {
            if (value == null)
            {
                throw new Exception("Valore non definito");
            }

            if (value is ExcelErrorValue)
            {
                throw new Exception("Valore non definito");
            }

            return Convert.ToDouble(value);
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

        private void WriteLog(List<string> logRows, string name)
        {
            var logName = string.Format("{0}{1}.txt", name, DateTime.Now.ToString("yyyyMMdd_HHmm"));
            using (var writer = new StreamWriter(Path.Combine(OutputFolder.Text, logName), false))
            {
                foreach (var logRow in logRows)
                {
                    writer.WriteLine(logRow);
                }
            }
        }

        private string GetCurrentBankName(string inputFile, int tipologia, int index)
        {
            var suffix = "";

            switch (tipologia)
            {
                case 0:
                    suffix = "_VITA";
                    break;
                case 1:
                    suffix = "_DANNI";
                    break;
                case 2:
                    suffix = "_DANNI777";
                    break;
                default:
                    break;
            }
            suffix += "_" + index.ToString().PadLeft(2, '0');
            return Path.Combine(OutputFolder.Text, Path.GetFileNameWithoutExtension(inputFile) + suffix + ".xml");
        }

        private string GetInputBankName(string inputFileName)
        {
            return Path.Combine(OutputFolder.Text, "TRADUZIONE_" + inputFileName);
        }

        private void OpenOutput_Click(object sender, EventArgs e)
        {
            Process.Start(OutputFolder.Text);
        }

        private void CreaExcelDaReportBase_Click(object sender, EventArgs e)
        {
            try
            {
                var conversioni = LeggiTabellaConversione();

                var openFileDialog = new OpenFileDialog();
                openFileDialog.CheckFileExists = true;
                openFileDialog.Filter = "Excel files (*.xlsx)|*.xlsx|All files (*.*)|*.*";
                openFileDialog.Title = "Seleziona il file report per pagamenti automatici";
                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    CreazioneExcelDaReportBase(openFileDialog.FileName, conversioni);

                    ShowInfo("File creato");
                }
            }
            catch (Exception ex)
            {
                ShowError(ex.Message);
            }
        }

        private void CreazioneExcelDaReportBase(string filePath, List<RigaConversionePagamenti> conversionList)
        {
            using (var excel = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = excel.Workbook.Worksheets[1];

                using (var excelTarget = new ExcelPackage())
                {
                    var wsTarget = excelTarget.Workbook.Worksheets.Add("InputBanca");

                    wsTarget.Cells[1, 1].Value = "Tipologia";
                    wsTarget.Cells[1, 2].Value = "Cognome";
                    wsTarget.Cells[1, 3].Value = "Nome";
                    wsTarget.Cells[1, 4].Value = "IBAN";
                    wsTarget.Cells[1, 5].Value = "Valore";
                    wsTarget.Cells[1, 6].Value = "Causale";

                    var exit = false;
                    var row = 2;
                    var targetRow = 2;

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
                                if (ws.Cells["D" + row.ToString()].Value == null)
                                {
                                    throw new Exception("Cognome colonna D non impostato in riga " + row.ToString());
                                }

                                if (ws.Cells["E" + row.ToString()].Value == null)
                                {
                                    throw new Exception("Nome colonna E non impostato in riga " + row.ToString());
                                }

                                if (ws.Cells["G" + row.ToString()].Value == null)
                                {
                                    throw new Exception("Policy name non impostato in riga " + row.ToString());
                                }

                                if (ws.Cells["H" + row.ToString()].Value == null)
                                {
                                    throw new Exception("Warranty non impostata in riga " + row.ToString());
                                }

                                var policyName = ws.Cells["G" + row.ToString()].Value.ToString();
                                var garanzia = ws.Cells["H" + row.ToString()].Value.ToString();

                                var rigaConvesione = conversionList.Where(c => c.ProductCode == policyName && c.Warranty == garanzia).FirstOrDefault();
                                if (rigaConvesione == null)
                                {
                                    throw new Exception(string.Format("Nella tabella conversioni prodotti non esiste la riga con Codice Prodotto = {0} e Garanzia = {1}", policyName, garanzia));
                                }

                                wsTarget.Cells[targetRow, 1].Value = rigaConvesione.Tipologia;

                                var nominativo = ws.Cells["D" + row.ToString()].Value.ToString() + " " + ws.Cells["E" + row.ToString()].Value.ToString();
                                if (ws.Cells["N" + row.ToString()].Value == null || ws.Cells["N" + row.ToString()].Value.ToString().Trim() == "")
                                {
                                    wsTarget.Cells[targetRow, 2].Value = GetString(ws.Cells["M" + row.ToString()].Value);
                                    wsTarget.Cells[targetRow, 3].Value = GetString(ws.Cells["L" + row.ToString()].Value);
                                }
                                else
                                {
                                    wsTarget.Cells[targetRow, 2].Value = ws.Cells["N" + row.ToString()].Value.ToString();
                                    wsTarget.Cells[targetRow, 3].Value = "";
                                }

                                if (ws.Cells["J" + row.ToString()].Value == null || ws.Cells["J" + row.ToString()].Value.ToString().Trim() == "")
                                {
                                    throw new Exception("IBAN non impostato in riga " + row.ToString());
                                }
                                wsTarget.Cells[targetRow, 4].Value = ws.Cells["J" + row.ToString()].Value.ToString();

                                if (ws.Cells["B" + row.ToString()].Value == null)
                                {
                                    throw new Exception("Amount non impostato in riga " + row.ToString());
                                }
                                wsTarget.Cells[targetRow, 5].Value = Convert.ToDouble(ws.Cells["B" + row.ToString()].Value);

                                var sottoscrizione = ws.Cells["F" + row.ToString()].Value.ToString();
                                var inizio = GetDate(ws, "P", row);
                                var fine = GetDate(ws, "Q", row);
                                var causale = string.Format("{0} {1} {2} - Liquidazione sx {3} {4} - {5}", sottoscrizione, nominativo, rigaConvesione.SiglaProdotto, garanzia, inizio, fine);
                                wsTarget.Cells[targetRow, 6].Value = causale;

                                row++;
                                targetRow++;
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                    }

                    excelTarget.SaveAs(new FileInfo(GetInputBankName(Path.GetFileName(filePath))));
                }
            }
        }

        private string GetDate(ExcelWorksheet ws, string col, int row)
        {
            var value = ws.Cells[col + row.ToString()].Value;

            if (value == null)
            {
                throw new Exception("Data mancante nella riga " + row.ToString());
            }

            if (value is int || value is double)
            {
                return FromExcelSerialDate(Convert.ToInt32(value)).ToString("dd/MM/yyyy");
            }

            if (value is DateTime)
            {
                return ((DateTime)value).ToString("dd/MM/yyyy");
            }

            if (value is string)
            {
                return value.ToString();
            }

            throw new Exception("Data non riconosciuta nella riga " + row.ToString());
        }

        public static DateTime FromExcelSerialDate(int SerialDate)
        {
            if (SerialDate > 59) SerialDate -= 1; //Excel/Lotus 2/29/1900 bug   
            return new DateTime(1899, 12, 31).AddDays(SerialDate);
        }

        private List<RigaConversionePagamenti> LeggiTabellaConversione()
        {
            var retVal = new List<RigaConversionePagamenti>();

            var filePath = GetFilePath("Modellini\\TabellaConversionePagamenti.xlsx");
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("File TabellaConversionePagamenti.xlsx non trovato nella cartella Modellini");
            }

            using (var excel = new ExcelPackage(new FileInfo(filePath)))
            {
                var ws = excel.Workbook.Worksheets[1];
                var exit = false;
                var row = 2;
                RigaConversionePagamenti lastRiga = null;

                while (!exit)
                {
                    if (ws.Cells[row, 1].Value == null || ws.Cells[row, 1].Value.ToString() == "")
                    {
                        exit = true;
                    }
                    else
                    {
                        var riga = new RigaConversionePagamenti();
                        riga.Insurer = ws.Cells[row, 1].Value.ToString();
                        if (ws.Cells[row, 2].Value == null)
                        {
                            if (lastRiga == null)
                            {
                                throw new Exception("Errore nella definizione del prodotto in riga " + row.ToString());
                            }

                            riga.ProductCode = lastRiga.ProductCode;
                        }
                        else
                        {
                            riga.ProductCode = ws.Cells[row, 2].Value.ToString();
                        }

                        riga.SiglaProdotto = ws.Cells[row, 3].Value.ToString();
                        riga.Warranty = ws.Cells[row, 4].Value.ToString();
                        riga.Tipologia = Convert.ToInt32(ws.Cells[row, 5].Value);
                        retVal.Add(riga);
                        row++;

                        lastRiga = riga;
                    }
                }
            }

            return retVal;
        }

        private string GetFilePath(string filePath)
        {
            var exePath = new FileInfo(Application.ExecutablePath);
            var retVal = Path.Combine(exePath.DirectoryName, filePath);

            return retVal;
        }

        private void Bank_Load(object sender, EventArgs e)
        {

        }
    }
}
