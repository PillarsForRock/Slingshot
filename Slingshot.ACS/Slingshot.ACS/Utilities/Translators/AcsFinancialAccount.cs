using System;
using System.Data;

using Slingshot.Core.Model;

namespace Slingshot.ACS.Utilities.Translators
{
    public static class AcsFinancialAccount
    {
        public static FinancialAccount Translate( DataRow row )
        {
            var financialAccount = new FinancialAccount();

            financialAccount.Id = AcsApi.ImportSource == ImportSource.CSVFiles ? ( row.Field<int?>( "FundNumber" ) ?? 0 ): row.Field<Int16>( "FundNumber" );
            financialAccount.GlCode = row.Field<string>( "FundCode" );
            financialAccount.Name = row.Field<string>( "FundDescription" );

            return financialAccount;
        }
    }
}
