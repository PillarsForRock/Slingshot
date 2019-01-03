using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Net;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using System.Xml;
using System.Xml.Linq;

using CsvHelper;
using CsvHelper.Configuration;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using RestSharp;
using Slingshot.Core;
using Slingshot.Core.Model;
using Slingshot.Core.Utilities;
using Slingshot.Elexio.Utilities.Translators;

namespace Slingshot.Elexio.Utilities
{
    /// <summary>
    /// API F1 Status
    /// </summary>
    public static class ElexioApi
    {
        private static RestClient _client;
        private static RestRequest _request;
        private static List<int> _uids = new List<int>();
        private static Dictionary<int, string> _accountLookups = new Dictionary<int, string>();

        public static MetaData MetaData { get; set; }

        /// <summary>
        ///  Set F1Api.DumpResponseToXmlFile to true to save all API Responses
        ///   to XML files and include them in the slingshot package
        /// </summary>
        /// <value>
        /// <c>true</c> if the response should get dumped to XML; otherwise, <c>false</c>.
        /// </value>
        public static bool DumpResponseToXmlFile { get; set; }

        /// <summary>
        /// Gets or sets the last run date.
        /// </summary>
        /// <value>
        /// The last run date.
        /// </value>
        public static DateTime LastRunDate { get; set; } = DateTime.MinValue;

        /// <summary>
        /// Gets or sets the api counter.
        /// </summary>
        /// <value>
        /// The api counter.
        /// </value>
        public static int ApiCounter { get; set; } = 0;

        /// <summary>
        /// Gets or sets the domain.
        /// </summary>
        /// <value>
        /// The domain.
        /// </value>
        public static string Hostname { get; set; }

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        /// <value>
        /// The error message.
        /// </value>
        public static string ErrorMessage { get; set; }

        /// <summary>
        /// Gets the API URL.
        /// </summary>
        /// <value>
        /// The API URL.
        /// </value>
        public static string ApiUrl
        {
            get
            {
                return $"https://{Hostname}.elexiochms.com/";
            }
        }

        /// <summary>
        /// Gets or sets the API username.
        /// </summary>
        /// <value>
        /// The API username.
        /// </value>
        public static string ApiUsername { get; set; }

        /// <summary>
        /// Gets or sets the API password.
        /// </summary>
        /// <value>
        /// The API password.
        /// </value>
        public static string ApiPassword { get; set; }

        /// <summary>
        /// Gets or sets the Session Id.
        /// </summary>
        /// <value>
        /// The Session Id.
        /// </value>
        public static string SessionId { get; set; }

        /// <summary>
        /// Gets a value indicating whether this instance is connected.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is connected; otherwise, <c>false</c>.
        /// </value>
        public static bool IsConnected { get; private set; } = false;

        #region API Call Paths

        private const string API_LOGIN = "api/user/login";
        private const string API_META_DATA = "api/user/get_meta_data";
        private const string API_INDIVIDUALS = "api/people/all";
        public const string API_INDIVIDUAL = "api/people/";
        private const string API_GIVING_CATEGORIES = "api/giving/categories";
        private const string API_PLEDGE_CAMPAIGNS = "api/pledges/campaigns";
        private const string API_PLEDGE_PEOPLE = "api/pledges/pledgers_for_campaign/";
        private const string API_GROUPS = "api/groups/sync";
        private const string API_GROUP_MEMBERS = "api/groups/";
        private const string API_ATTENDANCE = "api/attendance/for_person/";

        #endregion

        /// <summary>
        /// Initializes the export.
        /// </summary>
        public static void InitializeExport()
        {
            ImportPackage.InitalizePackageFolder();
            _uids = new List<int>();
            _accountLookups = new Dictionary<int, string>();
        }

        /// <summary>
        /// Connects the specified host name.
        /// </summary>
        /// <param name="hostName">Name of the host.</param>
        /// <param name="apiUsername">The API username.</param>
        /// <param name="apiPassword">The API password.</param>
        public static void Connect( string hostName, string apiUsername, string apiPassword )
        {
            Hostname = hostName;
            ApiUsername = apiUsername;
            ApiPassword = apiPassword;

            _client = new RestClient( ApiUrl );
            ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;

            _request = new RestRequest( API_LOGIN, Method.POST );
            _request.AddParameter( "username", ApiUsername );
            _request.AddParameter( "password", ApiPassword );

            // getting the api status sets the IsConnect flag
            UpdateApiStatus();
        }

        /// <summary>
        /// Updates the API status.
        /// </summary>
        public static void UpdateApiStatus()
        {
            // execute the request to get a oauth token and secret
            var response = _client.Execute( _request );
            LoginResponse responseObject = JsonConvert.DeserializeObject<LoginResponse>( response.Content );

            if ( response.StatusCode == System.Net.HttpStatusCode.OK && responseObject != null )
            {
                IsConnected = true;
                SessionId = responseObject.Data.SessionId;
            }
            else
            {
                ErrorMessage = response.Content;
            }
        }

        public static MetaData GetMetaData()
        {
            MetaData metaData = new MetaData();

            _request = new RestRequest( API_META_DATA, Method.GET );
            _request.AddQueryParameter( "session_id", SessionId );
            var response = _client.Execute( _request );

            metaData = JsonConvert.DeserializeObject<MetaData>( response.Content );

            return metaData;
        }

        /// <summary>
        /// Exports the individuals.
        /// </summary>
        public static void ExportIndividuals()
        {
            WritePersonAttributes();

            try
            {
                _request = new RestRequest( API_INDIVIDUALS, Method.GET );
                _request.AddQueryParameter( "session_id", SessionId );
                var response = _client.Execute( _request );
                ApiCounter++;

                dynamic data = JsonConvert.DeserializeObject( response.Content );

                var records = data.data;

                if ( records != null )
                {
                    foreach ( var letterGroup in records )
                    {
                        foreach ( var person in letterGroup.Value )
                        {
                            Person importPerson = ElexioPerson.Translate( person );
                            _uids.Add( importPerson.Id );

                            if ( importPerson != null )
                            {
                                ImportPackage.WriteToPackage( importPerson );
                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Exports the financial accounts.
        /// </summary>
        public static void ExportFinancialAccounts()
        {
            try
            {
                _request = new RestRequest( API_GIVING_CATEGORIES, Method.GET );
                _request.AddQueryParameter( "session_id", SessionId );
                var response = _client.Execute( _request );
                ApiCounter++;

                dynamic data = JsonConvert.DeserializeObject( response.Content );

                JArray records = data.data;

                if ( records.Count > 0 )
                {
                    foreach ( var account in records )
                    {
                        FinancialAccount importAccount = ElexioFinancialAccount.Translate( account );
                        if ( importAccount != null )
                        {
                            ImportPackage.WriteToPackage( importAccount );
                        }

                        _accountLookups.Add( (int)account["id"], (string)account["name"] );
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Exports the financial pledges.
        /// </summary>
        public static void ExportFinancialPledges()
        {
            try
            {
                _request = new RestRequest( API_PLEDGE_CAMPAIGNS, Method.GET );
                _request.AddQueryParameter( "session_id", SessionId );
                var response = _client.Execute( _request );
                ApiCounter++;

                dynamic data = JsonConvert.DeserializeObject( response.Content );

                JArray records = data.data;

                if ( records.Count > 0 )
                {
                    foreach ( var pledge in records )
                    {
                        string pledgeId = pledge.Value<string>( "campaignId" );

                        _request = new RestRequest( API_PLEDGE_PEOPLE + pledgeId.ToString(), Method.GET );
                        _request.AddQueryParameter( "session_id", SessionId );
                        response = _client.Execute( _request );
                        ApiCounter++;

                        data = JsonConvert.DeserializeObject( response.Content );

                        JArray personPledgeRecords = data.data;

                        if ( personPledgeRecords != null )
                        {
                            foreach ( var personPledge in personPledgeRecords )
                            {
                                FinancialPledge importPledge = ElexioFinancialPledge.Translate( personPledge );
                                if ( importPledge != null )
                                {
                                    ImportPackage.WriteToPackage( importPledge );
                                }
                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Exports the financial pledges.
        /// </summary>
        public static void ExportFinancialTransactions( string filename )
        {
            try
            {
                using ( TextReader reader = File.OpenText( filename ) )
                {
                    using ( var csv = new CsvReader( reader ) )
                    {
                        csv.Configuration.RegisterClassMap<GivingCSVMap>();

                        var records = csv.GetRecords<GivingCSV>().ToList();

                        // create a batch for each day
                        var batches = records.GroupBy( g => g.Date.Date,
                            ( key, g ) => new FinancialBatch
                            {
                                Id = ( key.ToString( "MMddyyyy" ) ).AsInteger(),
                                Name = key.ToShortDateString(),
                                StartDate = key.Date,
                                EndDate = key.Date,
                                Status = BatchStatus.Closed,
                            } ).ToList();

                        foreach ( FinancialBatch batch in batches )
                        {
                            var transactions = records.Where( r => r.Date == batch.StartDate.Value )
                              .Select( s => new FinancialTransaction
                              {
                                  Id = s.Id,
                                  BatchId = batch.Id,
                                  AuthorizedPersonId = s.UserId,
                                  TransactionCode = s.CheckNumber,
                                  TransactionDate = s.Date,
                                  Summary = s.Note,
                                  TransactionSource = TransactionSource.OnsiteCollection,
                                  TransactionType = TransactionType.Contribution
                              } ).ToList();

                            foreach ( FinancialTransaction transaction in transactions )
                            {
                                // check to see if the account already exists.  If not, it is likely the giving category is inactive and can't be exported via the API.
                                var transactionRecord = records.Where( r => r.Id == transaction.Id ).FirstOrDefault();
                                int accountId = _accountLookups.Where( a => a.Value == transactionRecord.Category ).Select( a => a.Key ).FirstOrDefault();
                                if ( accountId < 1 )
                                {
                                    MD5 md5Hasher = MD5.Create();
                                    var hashed = md5Hasher.ComputeHash( Encoding.UTF8.GetBytes( transactionRecord.Category ) );
                                    accountId = Math.Abs( BitConverter.ToInt32( hashed, 0 ) ); // used abs to ensure positive number

                                    // export this new financial account and then update the lookups
                                    ImportPackage.WriteToPackage( new FinancialAccount()
                                    {
                                        Id = accountId,
                                        Name = transactionRecord.Category,
                                        IsTaxDeductible = true
                                    } );

                                    _accountLookups.Add( accountId, transactionRecord.Category );
                                }

                                transaction.FinancialTransactionDetails.Add( new FinancialTransactionDetail
                                {
                                    Id = transaction.Id,
                                    TransactionId = transaction.Id,
                                    Amount = records.Where( r => r.Id == transaction.Id ).Select( r => r.Amount ).FirstOrDefault(),
                                    AccountId = accountId
                                } );
                            }

                            batch.FinancialTransactions.AddRange( transactions );

                            ImportPackage.WriteToPackage( batch );
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }


        /// <summary>
        /// Exports the groups.
        /// </summary>
        public static void ExportGroups()
        {
            // create Elexio group type
            ImportPackage.WriteToPackage( new GroupType()
            {
                Id = 9999,
                Name = "Elexio Group"
            } );

            try
            {
                _request = new RestRequest( API_GROUPS, Method.GET );
                _request.AddQueryParameter( "session_id", SessionId );
                var response = _client.Execute( _request );
                ApiCounter++;

                dynamic data = JsonConvert.DeserializeObject( response.Content );

                JArray records = data.data;

                if ( records.Count > 0 )
                {
                    foreach ( var group in records )
                    {
                        Slingshot.Core.Model.Group importGroup = ElexioGroup.Translate( group );
                        if ( importGroup != null )
                        {
                            ImportPackage.WriteToPackage( importGroup );
                        }

                        _request = new RestRequest( API_GROUP_MEMBERS + importGroup.Id.ToString() + "/people", Method.GET );
                        _request.AddQueryParameter( "session_id", SessionId );
                        response = _client.Execute( _request );
                        ApiCounter++;

                        data = JsonConvert.DeserializeObject( response.Content );

                        dynamic groupMemberRecords = data.data;

                        if ( groupMemberRecords != null )
                        {
                            foreach ( var letterGroup in groupMemberRecords )
                            {
                                foreach ( var groupMember in letterGroup.Value )
                                {
                                    GroupMember importGroupMember = ElexioGroupMember.Translate( groupMember, importGroup.Id );
                                    if ( importGroupMember != null )
                                    {
                                        ImportPackage.WriteToPackage( importGroupMember );
                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Exports the attendance.
        /// </summary>
        public static void ExportAttendance()
        {
            try
            {
                foreach ( var uid in _uids )
                {
                    _request = new RestRequest( API_ATTENDANCE + uid.ToString(), Method.GET );
                    _request.AddQueryParameter( "session_id", SessionId );
                    _request.AddQueryParameter( "count", "10000" );

                    var response = _client.Execute( _request );
                    ApiCounter++;

                    dynamic data = JsonConvert.DeserializeObject( response.Content );

                    JArray records = data.data.items;

                    if ( records.Count > 0 )
                    {
                        foreach ( var attendance in records )
                        {
                            Attendance importAttendance = ElexioAttendance.Translate( attendance );

                            if ( importAttendance != null )
                            {
                                ImportPackage.WriteToPackage( importAttendance );
                            }
                        }
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }
        }

        /// <summary>
        /// Writes the person attributes.
        /// </summary>
        public static void WritePersonAttributes()
        {
            MetaData = GetMetaData();

            //// dates 1 - 10
            // date 1
            if ( MetaData.data.dateFieldLabels.date1 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date1,
                    Key = MetaData.data.dateFieldLabels.date1.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 2
            if ( MetaData.data.dateFieldLabels.date2 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date2,
                    Key = MetaData.data.dateFieldLabels.date2.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 3
            if ( MetaData.data.dateFieldLabels.date3 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date3,
                    Key = MetaData.data.dateFieldLabels.date3.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 4
            if ( MetaData.data.dateFieldLabels.date4 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date4,
                    Key = MetaData.data.dateFieldLabels.date4.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 5
            if ( MetaData.data.dateFieldLabels.date5 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date5,
                    Key = MetaData.data.dateFieldLabels.date5.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 6
            if ( MetaData.data.dateFieldLabels.date6 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date6,
                    Key = MetaData.data.dateFieldLabels.date6.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 7
            if ( MetaData.data.dateFieldLabels.date7 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date7,
                    Key = MetaData.data.dateFieldLabels.date7.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 8
            if ( MetaData.data.dateFieldLabels.date8 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date8,
                    Key = MetaData.data.dateFieldLabels.date8.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 9
            if ( MetaData.data.dateFieldLabels.date9 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date9,
                    Key = MetaData.data.dateFieldLabels.date9.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            // date 10
            if ( MetaData.data.dateFieldLabels.date10 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.dateFieldLabels.date10,
                    Key = MetaData.data.dateFieldLabels.date10.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.DateFieldType"
                } );
            }

            //// text 1 - 15
            // text 1
            if ( MetaData.data.textFieldLabels.text1 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text1,
                    Key = MetaData.data.textFieldLabels.text1.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 2
            if ( MetaData.data.textFieldLabels.text2 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text2,
                    Key = MetaData.data.textFieldLabels.text2.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 3
            if ( MetaData.data.textFieldLabels.text3 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text3,
                    Key = MetaData.data.textFieldLabels.text3.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 4
            if ( MetaData.data.textFieldLabels.text4 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text4,
                    Key = MetaData.data.textFieldLabels.text4.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 5
            if ( MetaData.data.textFieldLabels.text5 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text5,
                    Key = MetaData.data.textFieldLabels.text5.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 6
            if ( MetaData.data.textFieldLabels.text6 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text6,
                    Key = MetaData.data.textFieldLabels.text6.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 7
            if ( MetaData.data.textFieldLabels.text7 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text7,
                    Key = MetaData.data.textFieldLabels.text7.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 8
            if ( MetaData.data.textFieldLabels.text8 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text8,
                    Key = MetaData.data.textFieldLabels.text8.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 9
            if ( MetaData.data.textFieldLabels.text9 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text9,
                    Key = MetaData.data.textFieldLabels.text9.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 10
            if ( MetaData.data.textFieldLabels.text10 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text10,
                    Key = MetaData.data.textFieldLabels.text10.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 11
            if ( MetaData.data.textFieldLabels.text11 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text11,
                    Key = MetaData.data.textFieldLabels.text11.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 12
            if ( MetaData.data.textFieldLabels.text12 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text12,
                    Key = MetaData.data.textFieldLabels.text12.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 13
            if ( MetaData.data.textFieldLabels.text13 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text13,
                    Key = MetaData.data.textFieldLabels.text13.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 14
            if ( MetaData.data.textFieldLabels.text14 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text14,
                    Key = MetaData.data.textFieldLabels.text14.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // text 15
            if ( MetaData.data.textFieldLabels.text15 != null )
            {
                ImportPackage.WriteToPackage( new PersonAttribute()
                {
                    Name = MetaData.data.textFieldLabels.text15,
                    Key = MetaData.data.textFieldLabels.text15.RemoveSpaces().RemoveSpecialCharacters(),
                    Category = "Elexio Attributes",
                    FieldType = "Rock.Field.Types.TextFieldType"
                } );
            }

            // death date
            ImportPackage.WriteToPackage( new PersonAttribute()
            {
                Name = "Death Date",
                Key = "DeathDate",
                Category = "Elexio Attributes",
                FieldType = "Rock.Field.Types.DateFieldType"
            } );
        }
    }

    public class LoginResponse
    {
        [JsonProperty( "sucess" )]
        public string Success { get; set; }

        [JsonProperty( "statusCode" )]
        public string StatusCode { get; set; }

        [JsonProperty( "data" )]
        public Data Data { get; set; }
    }

    public class Data
    {
        [JsonProperty( "fname" )]
        public string FirstName { get; set; }

        [JsonProperty( "lname" )]
        public string LastName { get; set; }

        [JsonProperty( "uid" )]
        public string uid { get; set; }

        [JsonProperty( "session_id" )]
        public string SessionId { get; set; }

        [JsonProperty( "org_name" )]
        public string OrgName { get; set; }
    }

    public class MetaData
    {
        public Data data { get; set; }

        public class Data
        {
            public DateFields dateFieldLabels { get; set; }

            public TextFields textFieldLabels { get; set; }

            public class DateFields
            {
                public string date1 { get; set; }
                public string date2 { get; set; }
                public string date3 { get; set; }
                public string date4 { get; set; }
                public string date5 { get; set; }
                public string date6 { get; set; }
                public string date7 { get; set; }
                public string date8 { get; set; }
                public string date9 { get; set; }
                public string date10 { get; set; }
            }

            public class TextFields
            {
                public string text1 { get; set; }
                public string text2 { get; set; }
                public string text3 { get; set; }
                public string text4 { get; set; }
                public string text5 { get; set; }
                public string text6 { get; set; }
                public string text7 { get; set; }
                public string text8 { get; set; }
                public string text9 { get; set; }
                public string text10 { get; set; }
                public string text11 { get; set; }
                public string text12 { get; set; }
                public string text13 { get; set; }
                public string text14 { get; set; }
                public string text15 { get; set; }
            }
        }
    }

    public class GivingCSV
    {
        public int Id { get; set; }

        public int UserId { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public decimal Amount { get; set; }

        public DateTime Date { get; set; }

        public string Category { get; set; }

        public bool TaxDeductible { get; set; }

        public string CheckNumber { get; set; }

        public string Note { get; set; }
    }

    public sealed class GivingCSVMap : CsvClassMap<GivingCSV>
    {
        public GivingCSVMap()
        {
            AutoMap();
            Map( m => m.UserId ).Name( "User Id" );
            Map( m => m.FirstName ).Name( "First Name" );
            Map( m => m.LastName ).Name( "Last Name" );
            Map( m => m.TaxDeductible ).Name( "Tax Deductible" );
            Map( m => m.CheckNumber ).Name( "Check Number" );
            Map( m => m.Amount ).ConvertUsing( m =>
                 {
                     return decimal.Parse( Regex.Replace( m.GetField( "Amount" ), @"[^\d.]", "" ) );
                 } );
        }
    }

}