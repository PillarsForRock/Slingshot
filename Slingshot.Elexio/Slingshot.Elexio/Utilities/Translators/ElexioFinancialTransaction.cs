using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioFinancialTransaction
    {
        public static FinancialTransaction Translate( GivingCSV importTransaction )
        {
            var transaction = new FinancialTransaction();
            transaction.Id = importTransaction.Id;
            transaction.AuthorizedPersonId = importTransaction.UserId;
            transaction.TransactionCode = importTransaction.CheckNumber;
            transaction.TransactionDate = importTransaction.Date;
            transaction.Summary = importTransaction.Note;
            transaction.TransactionSource = TransactionSource.OnsiteCollection;
            transaction.TransactionType = TransactionType.Contribution;

            return transaction;
        }
    }
}
