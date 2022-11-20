using CBP.DataLayer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP.Helpers
{
    public class BeneficiaryInfo
    {
        public int Id { get; set; }
        public DateTime? BirthDate { get; set; }
        public string FirstName { get; set; }
        public string Iban { get; set; }
        public string LastName { get; set; }
        public double Share { get; set; }
        public int TacticalRowId { get; set; }

        public BeneficiaryInfo() { }

        public BeneficiaryInfo(Beneficiary entity, UnitOfWork uow)
        {
            Id = entity.Id;
            BirthDate= entity.BirthDate;
            FirstName = entity.FirstName;
            Iban = entity.Iban;
            LastName = entity.LastName;
            Share = entity.Share;
            TacticalRowId = entity.TacticalRowId;
        }
    }
}
