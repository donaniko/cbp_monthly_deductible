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
    public partial class SearchBankTransactions : Form
    {
        private int countPolicies;
        private int pageIndex;
        private int pageCount;
        private int[] pageSizes = new int[] { 20, 40, 100 };

        public SearchBankTransactions()
        {
            InitializeComponent();

            foreach (var size in pageSizes)
            {
                ElencoSize.Items.Add(size + " per pagina");
            }

            ElencoSize.SelectedIndex = 0;

            NomeDistinta.Focus();
        }

        private void StartSearch_Click(object sender, EventArgs e)
        {
            pageIndex = 0;
            DoQuery();
        }

        private void DoQuery()
        {
            using (var uow = new UnitOfWork())
            {
                var query = uow.Context.BankTransfers.AsQueryable();

                var intVal = 0;
                if (int.TryParse(NumeroAdesione.Text, out intVal))
                {
                    query = query.Where(bt => bt.BankTransferPolicies.Any(btp => btp.TacticalRow.PolicyNumber == intVal));
                }

                if (!string.IsNullOrEmpty(NomeDistinta.Text))
                {
                    query = query.Where(bt => bt.Name.Contains(NomeDistinta.Text));
                }

                countPolicies = query.Count();
                pageCount = countPolicies / pageSizes[ElencoSize.SelectedIndex];
                if (countPolicies % pageSizes[ElencoSize.SelectedIndex] > 0)
                {
                    pageCount++;
                }

                PaginationLabel.Text = string.Format("Pagina {0} / {1}", (pageIndex + 1), pageCount);

                PrevAll.Enabled = Prev.Enabled = pageIndex > 0;
                Next.Enabled = NextAll.Enabled = pageIndex < pageCount - 1;

                try
                {
                    var dataSource = new List<BankTransferInfo>();
                    var list = query.OrderBy(bt => bt.DataDistinta).Skip(pageIndex * pageSizes[ElencoSize.SelectedIndex]).Take(pageSizes[ElencoSize.SelectedIndex]).ToList();
                    foreach (var bankTransfer in list)
                    {
                        dataSource.Add(new BankTransferInfo(bankTransfer, uow));
                    }
                    BankTransactions.DataSource = dataSource;
                }
                catch (Exception)
                {
                    ShowError("Errore esecuzione ricerca. Riprovare");
                }
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

        private void Next_Click(object sender, EventArgs e)
        {
            if (pageIndex < pageCount - 1)
            {
                pageIndex++;
                DoQuery();
            }
        }

        private void NextAll_Click(object sender, EventArgs e)
        {
            pageIndex = pageCount - 1;
            DoQuery();
        }

        private void Prev_Click(object sender, EventArgs e)
        {
            if (pageIndex > 0)
            {
                pageIndex--;
                DoQuery();
            }
        }

        private void PrevAll_Click(object sender, EventArgs e)
        {
            pageIndex = 0;
            DoQuery();
        }

        private void ElencoSize_SelectedIndexChanged(object sender, EventArgs e)
        {
            pageIndex = 0;
            DoQuery();
        }

        private void BankTransactions_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (BankTransactions.SelectedRows.Count > 0)
            {
                var row = (BankTransferInfo)BankTransactions.SelectedRows[0].DataBoundItem;

                if (e.ColumnIndex == 1)
                {
                    var dataDistinta = HelperClass.PromptDate(this, "Imposta la data distinta", "Data Distinta");
                    if (dataDistinta.HasValue)
                    {
                        Cursor = Cursors.WaitCursor;
                        Application.DoEvents();

                        HelperClass.ImpostaDataDistinta(row.Id, dataDistinta.Value);

                        StartSearch_Click(null, null);

                        Cursor = Cursors.Default;
                    }
                }
                else if (e.ColumnIndex == 2)
                {
                    var dataValuta = HelperClass.PromptDate(this, "Imposta la data di valuta", "Data Valuta");
                    if (dataValuta.HasValue)
                    {
                        Cursor = Cursors.WaitCursor;
                        Application.DoEvents();

                        HelperClass.ImpostaDataValuta(row.Id, dataValuta.Value);

                        StartSearch_Click(null, null);

                        Cursor = Cursors.Default;
                    }
                }
                else
                {
                    ShowInfo("Selezionare una distinta");
                }
            }
        }

        private void SearchText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                StartSearch_Click(sender, e);
            }
        }

        private void BankTransactions_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
