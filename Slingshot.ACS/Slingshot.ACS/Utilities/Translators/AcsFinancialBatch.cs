using System;
using System.Data;
using System.Security.Cryptography;
using System.Text;

using Slingshot.Core.Model;

namespace Slingshot.ACS.Utilities.Translators
{
    public static class AcsFinancialBatch
    {
        public static FinancialBatch Translate( DataRow row )
        {
            var financialBatch = new FinancialBatch();

            var giftDate = row.Field<DateTime?>( "GiftDate" );
            if ( !giftDate.HasValue )
            {
                return null;
            }

            string key = giftDate.Value.ToString();
            MD5 md5Hasher = MD5.Create();
            var hashed = md5Hasher.ComputeHash( Encoding.UTF8.GetBytes( key ) );
            var batchId = Math.Abs( BitConverter.ToInt32( hashed, 0 ) ); // used abs to ensure positive number
            if ( batchId > 0 )
            {
                financialBatch.Id = batchId;
            }

            financialBatch.Name = $"Imported Batch: {giftDate.Value.ToShortDateString()}";
            financialBatch.StartDate = giftDate.Value;
            financialBatch.EndDate = giftDate.Value;
            financialBatch.Status = BatchStatus.Closed;

            return financialBatch;
        }
    }
}
