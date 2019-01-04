﻿using System;
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

            // address
            string street1 = importGroup.address;
            string city = importGroup.city;
            string state = importGroup.state;
            string postalcode = importGroup.zipcode;

            if ( street1.IsNotNullOrWhitespace() && city.IsNotNullOrWhitespace() &&
                 state.IsNotNullOrWhitespace() && postalcode.IsNotNullOrWhitespace() )
            {
                var address = new GroupAddress();
                address.GroupId = group.Id;
                address.Street1 = street1;
                address.City = city;
                address.State = state;
                address.PostalCode = postalcode;
                address.AddressType = AddressType.Home;

                group.Addresses.Add( address );
            }

            // meeting day and time
            group.MeetingDay = importGroup.meetingDay;
            group.MeetingTime = importGroup.meetingTime;

            group.IsPublic = true;

            return group;
        }
    }
}
