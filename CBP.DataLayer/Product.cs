//------------------------------------------------------------------------------
// <auto-generated>
//     Codice generato da un modello.
//
//     Le modifiche manuali a questo file potrebbero causare un comportamento imprevisto dell'applicazione.
//     Se il codice viene rigenerato, le modifiche manuali al file verranno sovrascritte.
// </auto-generated>
//------------------------------------------------------------------------------

namespace CBP.DataLayer
{
    using System;
    using System.Collections.Generic;
    
    public partial class Product
    {
        public int Id { get; set; }
        public string Code { get; set; }
        public string Prodotto { get; set; }
        public string Convenzione { get; set; }
        public Nullable<double> FeesA { get; set; }
        public Nullable<double> FeesB { get; set; }
        public string Tipologia { get; set; }
        public string CodeCNPSI { get; set; }
    }
}
