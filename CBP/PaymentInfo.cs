using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP
{
    public class PaymentInfo
    {
        public int Tipologia { get; set; }
        public string Cognome { get; set; }
        public string Nome { get; set; }
        public string Iban { get; set; }
        public double Value { get; set; }
        public string Causale { get; set; }
        public int TacticalRowId { get; set; }
        public bool PagamentoAccantonato { get; set; }
        public DateTime? DataCancellazione { get; set; }

        public PaymentInfo() { }
    }
}
