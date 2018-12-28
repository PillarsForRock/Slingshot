using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioFinancialTransactionDetail
    {
        public static FinancialTransactionDetail Translate( GivingCSV importTransaction, Dictionary<int, string> accountLookups )
        {
            var transactionDetail = new FinancialTransactionDetail();
            transactionDetail.Id = importTransaction.Id;
            transactionDetail.TransactionId = importTransaction.Id;
            transactionDetail.Amount = importTransaction.Amount;
            transactionDetail.AccountId = accountLookups.Where( a => a.Value == importTransaction.Category ).Select( a => a.Key ).FirstOrDefault();

            return transactionDetail;
        }
    }
}
