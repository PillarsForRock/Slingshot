using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioFinancialPledge
    {
        public static FinancialPledge Translate( dynamic importFinancialPledge )
        {
            var pledge = new FinancialPledge();

            pledge.Id = importFinancialPledge.pledgeId;
            pledge.PersonId = importFinancialPledge.uid;
            pledge.AccountId = importFinancialPledge.categoryId;
            pledge.StartDate = importFinancialPledge.startDate;
            pledge.EndDate = importFinancialPledge.endDate;
            pledge.TotalAmount = importFinancialPledge.totalAmount;

            string pledgeFreq = importFinancialPledge.frequency;
            if ( pledgeFreq == "yearly" )
            {
                pledge.PledgeFrequency = PledgeFrequency.Yearly;
            }
            else
            {
                pledge.PledgeFrequency = PledgeFrequency.OneTime;
            }
            

            return pledge;
        }
    }
}
