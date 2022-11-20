using CBP.DataLayer;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP.Helpers
{
    public class BankTransferInfo
    {
        public int Id { get; set; }
        public DateTime DataDistinta { get; set; }
        public DateTime? DataValuta { get; set; }
        public string Name { get; set; }
        public int ConteggioBonifici { get; set; }
        public bool RimborsoDanni { get; set; }

        public BankTransferInfo(BankTransfer entity, UnitOfWork uow)
        {
            Id = entity.Id;
            DataDistinta = entity.DataDistinta;
            DataValuta = entity.DataValuta;
            Name = entity.Name;
            RimborsoDanni = entity.RimborsoDanni;
            ConteggioBonifici = uow.BankTransferPolicyRepository.Get(btp => btp.BankTransferId == Id).Count();
        }
    }
}
