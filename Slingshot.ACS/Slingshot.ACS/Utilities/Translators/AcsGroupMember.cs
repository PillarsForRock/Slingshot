using System;
using System.Data;
using System.Security.Cryptography;
using System.Text;

using Slingshot.Core.Model;

namespace Slingshot.ACS.Utilities.Translators
{
    public static class AcsGroupMember
    {
        public static GroupMember Translate( DataRow row )
        {
            var groupMember = new GroupMember();

            groupMember.PersonId = row.Field<int>( "IndividualId" );

            var groupPath = row.Field<string>( "GroupPath" );
            var groupName = row.Field<string>( "GroupName" );

            // generate a unique group id
            string key = groupPath + groupName;
            MD5 md5Hasher = MD5.Create();
            var hashed = md5Hasher.ComputeHash( Encoding.UTF8.GetBytes( key ) );
            var groupId = Math.Abs( BitConverter.ToInt32( hashed, 0 ) ); // used abs to ensure positive number
            if ( groupId > 0 )
            {
                groupMember.GroupId = groupId;
            }

            groupMember.Role = row.Field<string>( "Position" );

            return groupMember;
        }
    }
}
