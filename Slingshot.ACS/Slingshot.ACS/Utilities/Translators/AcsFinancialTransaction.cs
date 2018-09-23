using System;
using System.Data;
using System.Security.Cryptography;
using System.Text;

using Slingshot.Core.Model;

namespace Slingshot.ACS.Utilities.Translators
{
    public static class AcsFinancialTransaction
    {
        public static FinancialTransaction Translate( DataRow row )
        {
            var giftDate = row.Field<DateTime?>( "GiftDate" );
            if ( !giftDate.HasValue )
            {
                return null;
            }

            var financialTransaction = new FinancialTransaction();

            financialTransaction.Id = row.Field<int>( "TransactionID" );
            financialTransaction.TransactionCode = row.Field<string>( "CheckNumber" );
            financialTransaction.TransactionDate = giftDate.Value;
            financialTransaction.AuthorizedPersonId = row.Field<int>( "IndividualId" );

            // payment types can vary from Church to Church, so using the most popular ones here
            var source = row.Field<string>( "PaymentType" );
            switch ( source )
            {
                case "Online":
                    financialTransaction.TransactionSource = TransactionSource.Website;
                    break;
                case "Check":
                    financialTransaction.TransactionSource = TransactionSource.BankChecks;
                    financialTransaction.CurrencyType = CurrencyType.Check;
                    break;
                case "Credit Card":
                    financialTransaction.CurrencyType = CurrencyType.CreditCard;
                    break;
                case "Cash":
                    financialTransaction.TransactionSource = TransactionSource.OnsiteCollection;
                    financialTransaction.CurrencyType = CurrencyType.Cash;
                    break;
                default:
                    financialTransaction.TransactionSource = TransactionSource.OnsiteCollection;
                    financialTransaction.CurrencyType = CurrencyType.Unknown;
                    break;
            }

            // adding the original ACS payment type to the transaction summary for reference
            financialTransaction.Summary = "ACS PaymentType: " + source;

            string key = giftDate.Value.ToString();
            MD5 md5Hasher = MD5.Create();
            var hashed = md5Hasher.ComputeHash( Encoding.UTF8.GetBytes( key ) );
            var batchId = Math.Abs( BitConverter.ToInt32( hashed, 0 ) ); // used abs to ensure positive number
            if ( batchId > 0 )
            {
                financialTransaction.BatchId = batchId;
            }

            financialTransaction.TransactionType = TransactionType.Contribution;

            financialTransaction.CreatedDateTime = giftDate;

            return financialTransaction;
        }
    }
}
