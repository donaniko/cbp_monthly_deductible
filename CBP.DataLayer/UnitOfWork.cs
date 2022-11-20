using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CBP.DataLayer
{
    public class UnitOfWork : IDisposable
    {
        private bool disposed = false;
        private CBPEntities context = new CBPEntities();
        private GenericRepository<TacticalRow> tacticalRowsRepository;
        private GenericRepository<Product> productRepository;
        private GenericRepository<GeneralInfo> generalInfoRepository;
        private GenericRepository<BankTransfer> bankTransferRepository;
        private GenericRepository<BankTransferPolicy> bankTransferPolicyRepository;
        private GenericRepository<Beneficiary> beneficiaryRepository;
        private GenericRepository<Cancellation> cancellationRepository;

        public CBPEntities Context
        {
            get { return context; }
            set { context = value; }
        }

        public UnitOfWork()
        {
        }

        public string GetConnectionString()
        {
            return Context.Database.Connection.ConnectionString;
        }

        #region IUnitOfWork

        public GenericRepository<Cancellation> CancellationRepository
        {
            get
            {
                if (cancellationRepository == null)
                {
                    cancellationRepository = new GenericRepository<Cancellation>(context);
                }
                return cancellationRepository;
            }

        }

        public GenericRepository<Beneficiary> BeneficiaryRepository
        {
            get
            {
                if (beneficiaryRepository == null)
                {
                    beneficiaryRepository = new GenericRepository<Beneficiary>(context);
                }
                return beneficiaryRepository;
            }

        }

        public GenericRepository<BankTransferPolicy> BankTransferPolicyRepository
        {
            get
            {
                if (bankTransferPolicyRepository == null)
                {
                    bankTransferPolicyRepository = new GenericRepository<BankTransferPolicy>(context);
                }
                return bankTransferPolicyRepository;
            }

        }

        public GenericRepository<BankTransfer> BankTransferRepository
        {
            get
            {
                if (bankTransferRepository == null)
                {
                    bankTransferRepository = new GenericRepository<BankTransfer>(context);
                }
                return bankTransferRepository;
            }

        }

        public GenericRepository<TacticalRow> TacticalRowRepository
        {
            get
            {
                if (tacticalRowsRepository == null)
                {
                    tacticalRowsRepository = new GenericRepository<TacticalRow>(context);
                }
                return tacticalRowsRepository;
            }

        }

        public GenericRepository<Product> ProductRepository
        {
            get
            {
                if (productRepository == null)
                {
                    productRepository = new GenericRepository<Product>(context);
                }
                return productRepository;
            }

        }

        public GenericRepository<GeneralInfo> GeneralInfoRepository
        {
            get
            {
                if (generalInfoRepository == null)
                {
                    generalInfoRepository = new GenericRepository<GeneralInfo>(context);
                }
                return generalInfoRepository;
            }

        }
        public void Save()
        {
            context.SaveChanges();
        }

        #endregion

        #region IDisposable

        public void Dispose()
        {
            if (!this.disposed)
            {
                context.Dispose();
            }
            this.disposed = true;

            GC.SuppressFinalize(this);
        }

        #endregion
    }
}
