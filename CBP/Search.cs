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
    public partial class Search : Form
    {
        private int countPolicies;
        private int pageIndex;
        private int pageCount;
        private int[] pageSizes = new int[] { 20, 40, 100 };

        public Search()
        {
            InitializeComponent();

            foreach (var size in pageSizes)
            {
                ElencoSize.Items.Add(size + " per pagina");
            }
            
            ElencoSize.SelectedIndex = 0;
            Tipologia.SelectedIndex = 0;

            SearchText.Focus();
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
                var query = uow.Context.TacticalRows.AsQueryable();

                if (Tipologia.SelectedIndex > 0)
                {
                    query = query.Where(t => t.TransactionType == Helpers.HelperClass.TRANSATION_TYPE_RECESSO
                                             || t.TransactionType == Helpers.HelperClass.TRANSATION_TYPE_DISDETTA);

                    query = query.Where(t => t.RecessoDisdettaCompleto == (Tipologia.SelectedIndex == 2));
                }

                var intVal = 0;
                if (int.TryParse(SearchText.Text, out intVal))
                {
                    query = query.Where(t => t.PolicyNumber == intVal);
                }
                else if (!string.IsNullOrEmpty(SearchText.Text))
                {
                    query = query.Where(t => t.Surname.StartsWith(SearchText.Text) || t.FirstName.StartsWith(SearchText.Text));
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
                    var dataSource = new List<TacticalRowInfo>();
                    foreach (var tacticalRow in query.OrderBy(tr => tr.PolicyNumber).Skip(pageIndex * pageSizes[ElencoSize.SelectedIndex]).Take(pageSizes[ElencoSize.SelectedIndex]).ToList())
                    {
                        dataSource.Add(new TacticalRowInfo(tacticalRow, uow));
                    }
                    Policies.DataSource = dataSource;
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

        private void Policies_SelectionChanged(object sender, EventArgs e)
        {
            try
            {
                NumeroPolizza.Text = "";
                StatoPolizza.Text = "";
                LastName.Text = "";
                FirstName.Text = "";

                foreach (DataGridViewRow row in Policies.SelectedRows)
                {
                    NumeroPolizza.Text = row.Cells["PolicyNumber"].Value.ToString();
                    StatoPolizza.Text = row.Cells["TransactionType"].Value.ToString();
                    LastName.Text = row.Cells["Surname"].Value.ToString();
                    FirstName.Text = row.Cells["FirstName"].Value.ToString();
                }
            }
            catch (Exception)
            {

            }
        }

        private void NuovoRecesso_Click(object sender, EventArgs e)
        {
            var form = new PolicyEditor();
            if (form.ShowDialog(this) == DialogResult.OK)
            {
                ShowInfo("Polizza aggiunta. Numero adesione: " + form.Policy.PolicyNumber);

                SearchText.Text = form.Policy.PolicyNumber.ToString();
                StartSearch_Click(null, null);
            }
        }

        private void Policies_CellDoubleClick(object sender, DataGridViewCellEventArgs e)
        {
            if (Policies.SelectedRows.Count > 0)
            {
                var row = (TacticalRowInfo)Policies.SelectedRows[0].DataBoundItem;
                var policyEditor = new PolicyEditor(row);
                if (policyEditor.ShowDialog(this) == DialogResult.OK)
                {
                    ShowInfo("Polizza modificata");
                    StartSearch_Click(null, null);
                }
            }
            else
            {
                ShowInfo("Selezionare una polizza");
            }
        }

        private void SearchText_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                StartSearch_Click(sender, e);
            }
        }

        private void Recesso_Click(object sender, EventArgs e)
        {
            if (Policies.SelectedRows.Count > 0)
            {
                var row = (TacticalRowInfo)Policies.SelectedRows[0].DataBoundItem;

                var canModify = true;

                if (row.TransactionType != HelperClass.TRANSATION_TYPE_NEWBUSINESS && row.TransactionType != HelperClass.TRANSATION_TYPE_EA)
                {
                    canModify = false;
                }

                if (   (row.TransactionType == HelperClass.TRANSATION_TYPE_RECESSO
                    || row.TransactionType == HelperClass.TRANSATION_TYPE_DISDETTA)
                    && !row.RecessoDisdettaCompleto)
                {
                    canModify = true;
                }

                if (!canModify)
                {
                    ShowInfo("Impossibile impostare come recesso/disdetta questa pratica");
                    return;
                }

                var form = new RecessoDisdetta(row.Id);

                if (form.ShowDialog(this) == DialogResult.OK)
                {
                    ShowInfo("Polizza modificata");
                    StartSearch_Click(null, null);
                }
            }
            else
            {
                ShowInfo("Selezionare una polizza");
            }
        }

        private void Policies_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {

        }
    }
}
