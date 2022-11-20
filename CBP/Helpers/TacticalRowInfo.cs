using CBP.DataLayer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP.Helpers
{
    public class TacticalRowInfo
    {
        public int Id { get; set; }
        public int PolicyNumber { get; set; }
        public string Surname { get; set; }
        public string FirstName { get; set; }
        public System.DateTime Birth { get; set; }
        public int CalculatedAge { get; set; }
        public int Sex { get; set; }
        public string ZipCode { get; set; }
        public string Location { get; set; }
        public string TransactionType { get; set; }
        public string Tariff { get; set; }
        public Nullable<int> Duration { get; set; }
        public Nullable<System.DateTime> DateOfCancellation { get; set; }
        public Nullable<System.DateTime> DateOfSignature { get; set; }
        public Nullable<System.DateTime> ExpirationDate { get; set; }
        public Nullable<double> InsuredSum { get; set; }
        public double PremioLordo { get; set; }
        public Nullable<double> Tassi { get; set; }
        public Nullable<double> PremioNetto { get; set; }
        public Nullable<double> Commissioni { get; set; }
        public Nullable<double> PremioVita { get; set; }
        public Nullable<double> CommissioniVita { get; set; }
        public Nullable<double> PremioLordoDanni { get; set; }
        public Nullable<double> TassiDanni { get; set; }
        public Nullable<double> PremioNettoDanni { get; set; }
        public Nullable<double> CommissioniDanni { get; set; }
        public string Street { get; set; }
        public Nullable<double> InsuredRate { get; set; }
        public Nullable<double> ImportoEstinzione { get; set; }
        public Nullable<double> CapitaleResiduo { get; set; }
        public Nullable<System.DateTime> DataRimborso { get; set; }
        public string Provincia { get; set; }
        public string Iban { get; set; }
        public Nullable<System.DateTime> DateOfCancellation2 { get; set; }
        public Nullable<System.DateTime> DateOfCancellation3 { get; set; }
        public Nullable<double> ImportoEstinzione2 { get; set; }
        public Nullable<System.DateTime> ExpirationDate2 { get; set; }
        public Nullable<double> DebitoResiduo { get; set; }
        public string HE { get; set; }
        public string CognomeCoIntestario { get; set; }
        public string NomeCoIntestatario { get; set; }
        public Nullable<bool> RichiestaSottoscrittore { get; set; }
        public Nullable<bool> CartaIdentita { get; set; }
        public Nullable<bool> CodiceFiscale { get; set; }
        public Nullable<System.DateTime> DataLavorazione { get; set; }
        public Nullable<bool> Cointestatario { get; set; }
        public Nullable<bool> CointestatarioCartaIdentita { get; set; }
        public string NominativoAderente { get; set; }
        public string CodiceFiscaleAderente { get; set; }
        public Nullable<double> DebitoResiduoIniziale { get; set; }
        public Nullable<System.DateTime> CointestatarioDataNascita { get; set; }
        public bool RecessoDisdettaCompleto { get; set; }
        public bool Santander { get; set; }
        public string Targa { get; set; }
        public string Telaio { get; set; }

        public TacticalRowInfo() { }

        public TacticalRowInfo(TacticalRow entity, UnitOfWork uow)
        {
            Birth = entity.Birth;
            CalculatedAge = entity.CalculatedAge;
            CapitaleResiduo = entity.CapitaleResiduo;
            CartaIdentita = entity.CartaIdentita;
            CodiceFiscale = entity.CodiceFiscale;
            CodiceFiscaleAderente = entity.CodiceFiscaleAderente;
            CognomeCoIntestario = entity.CognomeCoIntestario;
            Cointestatario = entity.Cointestatario;
            CointestatarioCartaIdentita = entity.CointestatarioCartaIdentita;
            CointestatarioDataNascita = entity.CointestatarioDataNascita;
            Commissioni = entity.Commissioni;
            CommissioniDanni = entity.CommissioniDanni;
            CommissioniVita = entity.CommissioniVita;
            DataLavorazione = entity.DataLavorazione;
            DataRimborso = entity.DataRimborso;
            DateOfCancellation = entity.DateOfCancellation;
            DateOfCancellation2 = entity.DateOfCancellation2;
            DateOfCancellation3 = entity.DateOfCancellation3;
            DateOfSignature = entity.DateOfSignature;
            DebitoResiduo = entity.DebitoResiduo;
            DebitoResiduoIniziale = entity.DebitoResiduoIniziale;
            Duration = entity.Duration;
            ExpirationDate = entity.ExpirationDate;
            ExpirationDate2 = entity.ExpirationDate2;
            FirstName = entity.FirstName;
            HE = entity.HE;
            Iban = entity.Iban;
            Id = entity.Id;
            ImportoEstinzione = entity.ImportoEstinzione;
            ImportoEstinzione2 = entity.ImportoEstinzione2;
            InsuredRate = entity.InsuredRate;
            InsuredSum = entity.InsuredSum;
            Location = entity.Location;
            NomeCoIntestatario = entity.NomeCoIntestatario;
            NominativoAderente = entity.NominativoAderente;
            PolicyNumber = entity.PolicyNumber;
            PremioLordo = entity.PremioLordo;
            PremioLordoDanni = entity.PremioLordoDanni;
            PremioNetto = entity.PremioNetto;
            PremioNettoDanni = entity.PremioNettoDanni;
            PremioVita = entity.PremioVita;
            Provincia = entity.Provincia;
            RecessoDisdettaCompleto = entity.RecessoDisdettaCompleto;
            RichiestaSottoscrittore = entity.RichiestaSottoscrittore;
            Sex = entity.Sex;
            Street = entity.Street;
            Surname = entity.Surname;
            Tariff = entity.Tariff;
            Tassi = entity.Tassi;
            TassiDanni = entity.TassiDanni;
            TransactionType = entity.TransactionType;
            ZipCode = entity.ZipCode;
            Santander = entity.Santander;
            Targa = entity.Targa;
            Telaio = entity.Telaio;
        }
    }
}
