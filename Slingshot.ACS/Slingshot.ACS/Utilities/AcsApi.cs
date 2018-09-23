using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;

using Slingshot.ACS.Utilities.Translators;
using Slingshot.Core;
using Slingshot.Core.Model;
using Slingshot.Core.Utilities;

namespace Slingshot.ACS.Utilities
{
    public static class AcsApi
    {
        private static OleDbConnection _dbConnection;
        private static SqlConnection _sqlConnection;
        private static string _emailType;
        private static DateTime _modifiedSince = new DateTime( 1800, 1, 1 );

        /// <summary>
        /// Gets or sets the import source.
        /// </summary>
        /// <value>
        /// The import source.
        /// </value>
        public static ImportSource ImportSource { get; set; }

        /// <summary>
        /// Gets or sets the file name.
        /// </summary>
        /// <value>
        /// The file name.
        /// </value>
        public static string FileName { get; set; }

        /// <summary>
        /// Gets or sets the error message.
        /// </summary>
        /// <value>
        /// The error message.
        /// </value>
        public static string ErrorMessage { get; set; }

        /// <summary>
        /// Gets the local database file.
        /// </summary>
        /// <value>
        /// The local database file.
        /// </value>
        public static string LocalDatabaseFile
        {
            get
            {
                string directory = Directory.GetCurrentDirectory().ToLower().Replace( "\\bin\\debug", "" );
                return Path.Combine( directory, "ACSData.mdf" );
            }
        }

        /// <summary>
        /// Gets the connection string.
        /// </summary>
        /// <value>
        /// The connection string.
        /// </value>
        public static string ConnectionString
        {
            get
            {
                if ( ImportSource == ImportSource.AccessDb )
                {
                    return $"Provider=Microsoft.Jet.OLEDB.4.0;Data Source={FileName}";
                }
                else
                {
                    return $@"Data Source=(LocalDB)\MSSQLLocalDB;AttachDbFilename=""{LocalDatabaseFile}"";Integrated Security=True";
                }
            }
        }

        /// <summary>
        /// Gets a value indicating whether this instance is connected.
        /// </summary>
        /// <value>
        ///   <c>true</c> if this instance is connected; otherwise, <c>false</c>.
        /// </value>
        public static bool IsConnected { get; private set; } = false;

        /// <summary>
        /// Gets or sets the person attributes
        /// </summary>
        public static Dictionary<string, string> PersonAttributes { get; set; }

        #region SQL Queries

        private static string DateDesignator
        {
            get
            {
                return ImportSource == ImportSource.AccessDb ? "#" : "'";
            }
        }

        public static string SQL_GROUPS
        {
            get
            {
                return $@"
SELECT '' AS [Path], 'Activities' AS [GroupName], 0 AS [GroupSort]  
UNION
SELECT DISTINCT 'Activities', [ActGroup], 0 FROM [Activity] WHERE ISNULL([ActGroup],'') <> ''
UNION
SELECT DISTINCT CONCAT('Activities',[ActGroup]), [Category], 0 FROM [Activity] WHERE ISNULL([Category],'') <> ''
UNION
SELECT DISTINCT CONCAT('Activities',[ActGroup],[Category]), [List1], [List1Sort] FROM [Activity] WHERE ISNULL([List1],'') <> ''
UNION
SELECT DISTINCT CONCAT('Activities',[ActGroup],[Category],[List1]), [List2], [List2Sort] FROM [Activity] WHERE ISNULL([List2],'') <> ''
UNION
SELECT DISTINCT CONCAT('Activities',[ActGroup],[Category],[List1],[List2]), [List3], [List3Sort] FROM [Activity] WHERE ISNULL([List3],'') <> ''
UNION

SELECT '' AS [Path], 'Classes', 0 AS [GroupName] 
UNION
SELECT DISTINCT 'Classes', [Group], 0 FROM [ATClass] WHERE ISNULL([Group],'') <> '' AND ISNULL([Class],'') <> ''
UNION
SELECT DISTINCT CONCAT('Classes',[Group]), [Department], 0 FROM [ATClass] WHERE ISNULL([Department],'') <> '' AND ISNULL([Class],'') <> ''
UNION
SELECT DISTINCT CONCAT('Classes',[Group],[Department]), [Class], 0 FROM [ATClass] WHERE ISNULL([Class],'') <> ''
UNION

SELECT '' AS [Path], 'Small Groups' AS [GroupName], 0 AS [GroupSort]  
UNION
SELECT DISTINCT 'Small Groups', [Level1], [SortLevel1] FROM [SGRoster] WHERE ISNULL([Level1],'') <> ''
UNION
SELECT DISTINCT CONCAT('Small Groups',[Level1]), [Level2], [SortLevel2] FROM [SGRoster] WHERE ISNULL([Level2],'') <> ''
UNION
SELECT DISTINCT CONCAT('Small Groups',[Level1],[Level2]), [Level3], [SortLevel3] FROM [SGRoster] WHERE ISNULL([Level3],'') <> ''
UNION
SELECT DISTINCT CONCAT('Small Groups',[Level1],[Level2],[Level3]), [GroupName], [GroupSort] FROM [SGRoster] 
ORDER BY [Path]
";
            }
        }

        public static string SQL_GROUPMEMBERS
        {
            get
            {
                return $@"
SELECT
	P.[IndividualId]
	,A.[RosterID]
	,CASE 
        WHEN A.[List3] <> '' THEN CONCAT('Activities',A.[ActGroup],A.[Category],A.[List1],A.[List2])
        ELSE
            CASE 
                WHEN A.[List2] <> '' THEN CONCAT('Activities',A.[ActGroup],A.[Category],A.[List1])
                ELSE CONCAT('Activities',A.[ActGroup],A.[Category])
            END
    END AS [GroupPath]
	,CASE 
        WHEN A.[List3] <> '' THEN A.[List3] 
        ELSE
            CASE 
                WHEN A.[List2] <> '' THEN A.[List2]
                ELSE A.[List1] 
            END
    END AS [GroupName]
	,'Member' AS [Position]
FROM Activity A
        INNER JOIN people P 
            ON P.[FamilyNumber] = A.[FamilyNumber] 
                AND P.[IndividualNumber] = A.[IndividualNumber]

UNION ALL

SELECT
	P.[IndividualId]
	,A.[RosterID]
    ,CONCAT('Classes',[Group],[Department]) AS [GroupPath]
	,A.[Class] AS [GroupName]
	,'Member' AS [Position]
FROM ATClass A
        INNER JOIN people P 
            ON P.[FamilyNumber] = A.[FamilyNumber] 
                AND P.[IndividualNumber] = A.[IndividualNumber]
WHERE ISNULL(A.[Class],'') <> ''

UNION ALL

SELECT
	P.[IndividualId]
	,SG.[RosterID]
    ,CONCAT('Small Groups',SG.[Level1],SG.[Level2],SG.[Level3]) AS [GroupPath]
	,SG.[GroupName]
	,SG.[Position]
FROM SGRoster SG
        INNER JOIN people P 
            ON P.[FamilyNumber] = SG.[FamilyNumber] 
                AND P.[IndividualNumber] = SG.[IndividualNumber]";
            }
        }

        public static string SQL_PEOPLE
        {
            get
            {
                return $@"
SELECT P.*, E.[EmailAddr]
FROM people P
LEFT JOIN emails E ON (P.[FamilyNumber] = E.[FamilyNumber]
	AND P.[IndividualNumber] = E.[IndividualNumber]
	AND E.[Description] = '{ _emailType }')
WHERE P.[DateLastChanged] >= { DateDesignator }{ _modifiedSince.ToShortDateString() }{ DateDesignator }";
            }
        }

        public static string SQL_PHONES
        {
            get
            {
                return $@"
SELECT PH.*, 
       P.[IndividualId] 
FROM   phones PH 
       INNER JOIN people P 
               ON P.[FamilyNumber] = PH.[FamilyNumber] 
                  AND P.[IndividualNumber] = PH.[IndividualNumber]
WHERE P.[DateLastChanged] >= { DateDesignator }{ _modifiedSince.ToShortDateString() }{ DateDesignator }";
            }
        }

        public static string SQL_PEOPLE_NOTES
        {
            get
            {
                return $@"
SELECT P.[IndividualId]
      ,IC.[ComtDate]
      ,IC.[ComtType]
      ,IC.[Comment]
  FROM [IComment] IC
          INNER JOIN people P 
            ON P.[FamilyNumber] = IC.[FamilyNumber] 
                AND P.[IndividualNumber] = IC.[IndividualNumber]
GROUP BY P.[IndividualId], IC.[ComtDate], IC.[ComtType], IC.[Comment]";
            }
        }

        public static string SQL_FAMILY_NOTES
        {
            get
            {
                return $@"
SELECT FC.[FamilyNumber]
    ,FC.[ComtDate]
    ,FC.[ComtType]
    ,FC.[Comment]
 FROM [FComment] FC
 GROUP BY FC.[FamilyNumber], FC.[ComtDate], FC.[ComtType], FC.[Comment]";
            }
        }

        private const string SQL_FINANCIAL_ACCOUNTS = @"
SELECT [FundDescription], 
       [FundNumber], 
       [FundCode] 
FROM   cbgifts 
GROUP  BY [FundDescription], 
          [FundNumber], 
          [FundCode] ";

        public static string SQL_FINANCIAL_BATCHES
        {
            get
            {
                return $@"
SELECT 
    [GiftDate], 
    SUM([Amount]) AS [ControlAmount]
FROM [CBGifts]
GROUP BY [GiftDate]";
            }
        }

        public static string SQL_FINANCIAL_TRANSACTIONS
        {
            get
            {
                return $@"
SELECT DISTINCT G.[TransactionID], 
                G.[CheckNumber], 
                G.[GiftDate], 
                G.[PaymentType], 
                P.[IndividualId] 
FROM   cbgifts G 
       INNER JOIN people P 
               ON P.[FamilyNumber] = G.[FamilyNumber] 
                  AND P.[IndividualNumber] = G.[IndividualNumber]
WHERE G.[GiftDate] >= { DateDesignator }{ _modifiedSince.ToShortDateString() }{ DateDesignator }";
            }
        }

        public static string SQL_FINANCIAL_TRANSACTIONDETAILS
        {
            get
            {
                return $@"
SELECT [FundNumber], 
       [Amount], 
       [TransactionID], 
       [GiftDescription], 
       [GiftDate] 
FROM   cbgifts G 
       INNER JOIN people P 
               ON P.[FamilyNumber] = G.[FamilyNumber] 
                  AND P.[IndividualNumber] = G.[IndividualNumber]
WHERE G.[GiftDate] >= { DateDesignator }{ _modifiedSince.ToShortDateString() }{ DateDesignator }";
            }
        }

        private static string SQL_FINANCIAL_PLEDGES
        {
            get
            {
                return $@"
SELECT PL.[PledgeID],
				PL.[StartDate],
				PL.[StopDate],
				PL.[TotalPled],
				PL.[Freq],
				PL.[FundNumber],
                PL.[EntryDate],
                P.[IndividualId] 
FROM   cbPledge PL 
       INNER JOIN people P 
               ON P.[FamilyNumber] = PL.[FamilyNumber] 
                  AND P.[IndividualNumber] = PL.[IndividualNumber]
WHERE PL.[EntryDate] >= { DateDesignator }{ _modifiedSince.ToShortDateString() }{ DateDesignator }";
            }
        }

        public const string SQL_EMAIL_TYPES = @"
SELECT [Description]
FROM [Emails]
GROUP BY [Description]";

        #endregion region


        /// <summary>
        /// Initializes the export.
        /// </summary>
        public static void InitializeExport()
        {
            ImportPackage.InitalizePackageFolder();
        }

        /// <summary>
        /// Opens the specified MS Access database.
        /// </summary>
        /// <param name="importSource">The import source.</param>
        /// <param name="fileName">Name of the file (if using AccessDb source</param>
        public static void OpenConnection( ImportSource importSource, string fileName )
        {
            ImportSource = importSource;
            FileName = fileName;
            string connStr = ConnectionString;

            if ( importSource == ImportSource.AccessDb )
            {
                _dbConnection = new OleDbConnection { ConnectionString = connStr };
            }
            else
            {
                _sqlConnection = new SqlConnection { ConnectionString = connStr };
            }

            AcsApi.IsConnected = true;
        }

        /// <summary>
        /// Exports the individuals.
        /// </summary>
        public static void ExportIndividuals( DateTime modifiedSince, string emailType, string campusKey, string photoDirectory )
        {
            try
            {
                _modifiedSince = modifiedSince;
                _emailType = emailType;

                //load attributes
                LoadPersonAttributes();

                // Add additional attributes 
                PersonAttributes.Add( "JoinedHow", "String" );
                PersonAttributes.Add( "OtherMaritalStatus", "String" );

                // write out the person attributes
                WritePersonAttributes();

                // export people
                using ( var dtPeople = GetTableData( SQL_PEOPLE ) )
                {
                    foreach ( DataRow row in dtPeople.Rows )
                    {
                        var importPerson = AcsPerson.Translate( row, campusKey, photoDirectory );

                        if ( importPerson != null )
                        {
                            ImportPackage.WriteToPackage( importPerson );
                        }
                    }
                }

                // export person notes
                using ( var dtPeopleNotes = GetTableData( SQL_PEOPLE_NOTES ) )
                {
                    foreach ( DataRow row in dtPeopleNotes.Rows )
                    {
                        var importNote = AcsPersonNote.Translate( row );

                        if ( importNote != null )
                        {
                            ImportPackage.WriteToPackage( importNote );
                        }
                    }
                }

                // export family notes
                using ( var dtFamilyNotes = GetTableData( SQL_FAMILY_NOTES ) )
                {
                    foreach ( DataRow row in dtFamilyNotes.Rows )
                    {
                        var importNote = AcsFamilyNote.Translate( row );

                        if ( importNote != null )
                        {
                            ImportPackage.WriteToPackage( importNote );
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
        /// Exports the phone numbers
        /// </summary>
        public static void ExportPhoneNumbers( DateTime modifiedSince )
        {
            try
            {
                _modifiedSince = modifiedSince;

                using ( var dtPhones = GetTableData( SQL_PHONES ) )
                {
                    foreach ( DataRow row in dtPhones.Rows )
                    {
                        var importPhone = AcsPhone.Translate( row );

                        if ( importPhone != null )
                        {
                            ImportPackage.WriteToPackage( importPhone );
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
        /// Exports the funds.
        /// </summary>
        public static void ExportFunds()
        {
            try
            {
                using ( var dtFunds = GetTableData( SQL_FINANCIAL_ACCOUNTS ) )
                {
                    foreach ( DataRow row in dtFunds.Rows )
                    {
                        var importAccount = AcsFinancialAccount.Translate( row );

                        if ( importAccount != null && importAccount.Id > 0 )
                        {
                            ImportPackage.WriteToPackage( importAccount );
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
        /// Exports any contributions.  Currently, the ACS export file doesn't include
        ///  batches.
        /// </summary>
        public static void ExportContributions( DateTime modifiedSince )
        {
            try
            {
                // since the ACS export doesn't include batches and Rock expects transactions
                //  to belong to a batch, create a batch for each day.
                using ( var dtBatches = GetTableData( SQL_FINANCIAL_BATCHES ) )
                {
                    foreach ( DataRow row in dtBatches.Rows )
                    {
                        var importFinancialBatch = AcsFinancialBatch.Translate( row );

                        if ( importFinancialBatch != null )
                        {
                            ImportPackage.WriteToPackage( importFinancialBatch );
                        }
                    }
                }

                using ( var dtContributions = GetTableData( SQL_FINANCIAL_TRANSACTIONS ) )
                {
                    foreach ( DataRow row in dtContributions.Rows )
                    {
                        var importFinancialTransaction = AcsFinancialTransaction.Translate( row );

                        if ( importFinancialTransaction != null )
                        {
                            ImportPackage.WriteToPackage( importFinancialTransaction );
                        }
                    }
                }

                using ( var dtContributionDetails = GetTableData( SQL_FINANCIAL_TRANSACTIONDETAILS ) )
                {
                    foreach ( DataRow row in dtContributionDetails.Rows )
                    {
                        var importFinancialTransactionDetail = AcsFinancialTransactionDetail.Translate( row );

                        if ( importFinancialTransactionDetail != null )
                        {
                            ImportPackage.WriteToPackage( importFinancialTransactionDetail );
                        }
                    }
                }

                using ( var dtPledges = GetTableData( SQL_FINANCIAL_PLEDGES ) )
                {
                    foreach ( DataRow row in dtPledges.Rows )
                    {
                        var importFinancialPledge = AcsFinancialPledge.Translate( row );

                        if ( importFinancialPledge != null )
                        {
                            ImportPackage.WriteToPackage( importFinancialPledge );
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
        /// Exports any groups found.  Currently, this export doesn't support
        ///  group heirarchies and all groups will be imported to the
        ///  root of the group viewer.
        /// </summary>
        public static void ExportGroups()
        {
            try
            {
                WriteGroupTypes();

                using ( var dtGroups = GetTableData( SQL_GROUPS ) )
                {
                    foreach ( DataRow row in dtGroups.Rows )
                    {
                        var importGroup = AcsGroup.Translate( row );

                        if ( importGroup != null )
                        {
                            ImportPackage.WriteToPackage( importGroup );
                        }
                    }
                }

                using ( var dtGroupMembers = GetTableData( SQL_GROUPMEMBERS ) )
                {
                    foreach ( DataRow row in dtGroupMembers.Rows )
                    {
                        var importGroupMember = AcsGroupMember.Translate( row );

                        if ( importGroupMember != null )
                        {
                            ImportPackage.WriteToPackage( importGroupMember );
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
        /// Gets the table data.
        /// </summary>
        /// <param name="command">The SQL command to run.</param>
        /// <returns></returns>
        public static DataTable GetTableData( string command )
        {
            DataSet dataSet = new DataSet();
            DataTable dataTable = new DataTable();

            if ( ImportSource == ImportSource.AccessDb )
            {
                OleDbCommand dbCommand = new OleDbCommand( command, _dbConnection );
                OleDbDataAdapter adapter = new OleDbDataAdapter();

                adapter.SelectCommand = dbCommand;
                adapter.Fill( dataSet );
            }
            else
            {
                SqlCommand sqlCommand = new SqlCommand( command, _sqlConnection );
                SqlDataAdapter sqlAdapter = new SqlDataAdapter();
                sqlAdapter.SelectCommand = sqlCommand;
                sqlAdapter.Fill( dataSet );
            }

            dataTable = dataSet.Tables["Table"];

            return dataTable;
        }

        /// <summary>
        /// Loads the available person attributes.
        /// </summary>
        public static void LoadPersonAttributes()
        {
            PersonAttributes = new Dictionary<string, string>();

            var dataTable = GetTableData( SQL_PEOPLE );

            foreach ( DataColumn column in dataTable.Columns )
            {
                string columnName = column.ColumnName;

                // Person attributes always start with "Ind"
                if ( columnName.StartsWith( "Ind" ) && !columnName.StartsWith( "Individual" ) )
                {
                    PersonAttributes.Add( column.ColumnName, column.DataType.Name );
                }
            }
        }

        /// <summary>
        /// Loads the available person fields from the ACS export.
        /// </summary>
        /// <returns></returns>
        public static List<string> LoadPersonFields()
        {
            var personFields = new List<string>();

            try
            {
                var dataTable = GetTableData( SQL_PEOPLE );

                foreach ( DataColumn column in dataTable.Columns )
                {
                    string columnName = column.ColumnName;

                    // Person attributes always start with "Ind"
                    if ( ( columnName.StartsWith( "Ind" ) || columnName.StartsWith( "Fam" ) ) && 
                           !columnName.StartsWith( "Individual" ) && !columnName.StartsWith( "Family" ) )
                    {
                        personFields.Add( column.ColumnName );
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }

            return personFields;
        }

        /// <summary>
        /// Writes the person attributes.
        /// </summary>
        public static void WritePersonAttributes()
        {
            foreach ( var attrib in PersonAttributes )
            {
                var attribute = new PersonAttribute();

                // Remove 'Ind' and Add spaces between words
                string name = attrib.Key.StartsWith( "Ind" ) ? attrib.Key.Substring( 3 ) : attrib.Key;
                attribute.Name = ExtensionMethods.SplitCase( name );
                attribute.Key = attrib.Key;
                attribute.Category = "Imported Attributes";

                switch ( attrib.Value )
                {
                    case "String":
                        attribute.FieldType = "Rock.Field.Types.TextFieldType";
                        break;
                    case "DateTime":
                        attribute.FieldType = "Rock.Field.Types.DateTimeFieldType";
                        break;
                    default:
                        attribute.FieldType = "Rock.Field.Types.TextFieldType";
                        break;
                }

                ImportPackage.WriteToPackage( attribute );
            }
        }

        /// <summary>
        /// Writes the group types.
        /// </summary>
        public static void WriteGroupTypes()
        {
            // hardcode a generic group type
            ImportPackage.WriteToPackage( new GroupType()
            {
                Id = 9999,
                Name = "Imported Group"
            } );
        }

        /// <summary>
        /// Gets the email types.
        /// </summary>
        /// <returns>A list of email types.</returns>
        public static List<string> GetEmailTypes()
        {
            List<string> emailTypes = new List<string>();

            try
            {
                using ( var dtEmailTypes = GetTableData( SQL_EMAIL_TYPES ) )
                {
                    foreach ( DataRow row in dtEmailTypes.Rows )
                    {
                        emailTypes.Add( row.Field<string>( "Description" ) );
                    }
                }
            }
            catch ( Exception ex )
            {
                ErrorMessage = ex.Message;
            }

            return emailTypes;
        }
    }

    public enum ImportSource
    {
        AccessDb = 0,
        CSVFiles = 1
    }
}
