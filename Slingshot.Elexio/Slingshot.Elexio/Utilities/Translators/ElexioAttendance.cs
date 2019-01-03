﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Slingshot.Core.Model;

namespace Slingshot.Elexio.Utilities.Translators
{
    public static class ElexioAttendance
    {
        public static Attendance Translate( dynamic importAttendance )
        {
            var attendance = new Attendance();

            attendance.PersonId = importAttendance.uid;
            attendance.StartDateTime = importAttendance.date;
            attendance.GroupId = importAttendance.gid;
            attendance.Note = importAttendance.reason;

            return attendance;
        }
    }
}