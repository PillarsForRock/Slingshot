using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioGroup
    {
        public static Group Translate( dynamic importGroup )
        {
            var group = new Group();

            group.Id = importGroup.gid;
            group.GroupTypeId = 9999;
            group.Name = importGroup.name;
            group.Description = importGroup.description;

            string active = importGroup.active;
            group.IsActive = active.AsBoolean();

            group.MeetingDay = importGroup.meetingDay;
            group.MeetingTime = importGroup.meetingTime;

            group.IsPublic = true;

            return group;
        }
    }
}
