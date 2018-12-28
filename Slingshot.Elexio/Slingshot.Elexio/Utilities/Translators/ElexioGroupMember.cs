using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioGroupMember
    {
        public static GroupMember Translate( dynamic importGroupMember, int groupId )
        {
            var groupMember = new GroupMember();

            groupMember.PersonId = importGroupMember.uid;
            groupMember.GroupId = groupId;
            groupMember.Role = "Member";

            return groupMember;
        }
    }
}
