using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Dynamic;
using System.IO;
using System.Linq;
using System.Text;

using CsvHelper;

using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Management.Smo;

namespace Slingshot.ACS.Utilities.CsvToSql
{
    public static class CreateTables
    {
        private static Dictionary<string, List<CsvFieldInfo>> fieldOverrides = new Dictionary<string, List<CsvFieldInfo>>();

        private const string DB_NAME = "ACSData";

        public static void FromFolder( string csvfolder, string sqlConnectionString )
        {
            fieldOverrides.Add( "CBGifts", new List<CsvFieldInfo>() );
            fieldOverrides["CBGifts"].Add ( new CsvFieldInfo { Name = "FundCode", IsBool = false, IsDateTime = false, IsDecimal = false, IsInteger = false, MaxLength = 50 } );

            if ( Directory.Exists( csvfolder ) )
            {
                SqlConnection sqlConnection = new SqlConnection( sqlConnectionString );
                ServerConnection conn = new ServerConnection( sqlConnection );
                Server server = new Server( conn );
                var db = server.Databases[AcsApi.LocalDatabaseFile];

                foreach ( string csvFile in Directory.GetFiles( csvfolder, "*.csv" ) )
                {
                    ProcessFile( csvfolder, sqlConnectionString, db, Path.GetFileNameWithoutExtension( csvFile ) );
                }
            }
        }

        private static void ProcessFile( string csvfolder, string sqlConnectionString, Database database, string tableName )
        {
            var records = GetRecords( csvfolder, tableName );

            var fields = GetFields( records );

            CreateTable( database, tableName, fields );

            InsertRecords( sqlConnectionString, tableName, fields, records );
        }

        private static List<CsvFieldInfo> GetFields( List<dynamic> records )
        {
            var fields = new List<CsvFieldInfo>();

            if ( records != null && records.Any() )
            {

                ExpandoObject obj = records[0];
                foreach ( var p in obj )
                {
                    fields.Add( new CsvFieldInfo { Name = p.Key } );
                }

                foreach ( dynamic record in records )
                {
                    foreach ( var p in record )
                    {
                        var fieldInfo = fields.FirstOrDefault( f => f.Name == p.Key );
                        if ( fieldInfo != null )
                        {
                            string value = p.Value.ToString().Trim();
                            if ( value != null && value != string.Empty )
                            {
                                int len = value.Length;
                                fieldInfo.MaxLength = len > fieldInfo.MaxLength ? len : fieldInfo.MaxLength;

                                if ( fieldInfo.IsDateTime )
                                {
                                    DateTime datetime;
                                    if ( !DateTime.TryParse( value, out datetime ) )
                                    {
                                        fieldInfo.IsDateTime = false;
                                    }
                                }

                                if ( fieldInfo.IsBool )
                                {
                                    bool boolValue;
                                    if ( !bool.TryParse( value, out boolValue ) )
                                    {
                                        fieldInfo.IsBool = false;
                                    }
                                }

                                if ( fieldInfo.IsInteger )
                                {
                                    int number;
                                    if ( !int.TryParse( value, out number ) )
                                    {
                                        fieldInfo.IsInteger = false;
                                    }
                                }

                                if ( !fieldInfo.IsInteger && fieldInfo.IsDecimal )
                                {
                                    decimal number;
                                    if ( !decimal.TryParse( value, out number ) )
                                    {
                                        fieldInfo.IsDecimal = false;
                                    }
                                }
                            }
                        }
                    }
                }

                foreach ( var field in fields )
                {
                    var l = field.MaxLength;
                    field.MaxLength = 50 - ( l % 50 ) + l;
                }
            }

            return fields;
        }

        private static List<dynamic> GetRecords( string folder, string tableName )
        {
            string fileName = Path.Combine( folder, $"{tableName}.csv" );
            if ( File.Exists( fileName ) )
            {
                using ( var s = File.OpenText( fileName ) )
                {
                    var reader = new CsvReader( s );
                    reader.Configuration.HasHeaderRecord = true;
                    var records = reader.GetRecords<dynamic>().ToList();
                    return records;
                }
            }

            return null;
        }

        private static void CreateTable( Database database, string tableName, List<CsvFieldInfo> fields )
        {
            if ( database.Tables.Contains( tableName ) ) 
            {
                var oldTbl = database.Tables[tableName];
                oldTbl.Drop();
            }

            var tbl = new Table( database, tableName );
            foreach ( var field in fields )
            {
                var fieldInfo = field;
                if ( fieldOverrides.ContainsKey( tableName ) )
                {
                    fieldInfo = fieldOverrides[tableName].FirstOrDefault( f => f.Name == fieldInfo.Name ) ?? field;
                }

                var dataType = fieldInfo.MaxLength >= 2000 ? DataType.VarCharMax : DataType.VarChar( fieldInfo.MaxLength );
                if ( fieldInfo.IsDecimal )
                {
                    dataType = DataType.Decimal( 2, 18 );
                }
                if ( fieldInfo.IsInteger )
                {
                    dataType = DataType.Int;
                }
                if ( fieldInfo.IsDateTime )
                {
                    dataType = DataType.DateTime;
                }
                if ( fieldInfo.IsBool )
                {
                    dataType = DataType.Bit;
                }

                var col = new Column( tbl, fieldInfo.Name, dataType );
                tbl.Columns.Add( col );
            }

            if ( tbl.Columns.Count > 0 )
            {
                tbl.Create();
            }
        }

        private static void InsertRecords( string sqlConnectionString, string tableName, List<CsvFieldInfo> fields, List<dynamic> records )
        {
            var insertPrefix = new StringBuilder();
            foreach ( var field in fields )
            {
                if ( insertPrefix.Length == 0 )
                {
                    insertPrefix.AppendFormat( "INSERT INTO [{0}] ( [{1}]", tableName, field.Name );
                }
                else
                {
                    insertPrefix.AppendFormat( ", [{0}]", field.Name );
                }
            }

            insertPrefix.AppendFormat( " ){0}    VALUES ", Environment.NewLine );

            using ( SqlConnection connection = new SqlConnection( sqlConnectionString ) )
            {
                connection.Open();

                foreach ( var record in records )
                {
                    var insertValues = new StringBuilder();

                    foreach ( var p in record )
                    {
                        var fieldInfo = fields.FirstOrDefault( f => f.Name == p.Key );
                        if ( fieldInfo != null )
                        {
                            string value = p.Value.ToString().Trim();

                            insertValues.Append( insertValues.Length == 0 ? "( " : ", " );

                            if ( fieldInfo.IsBool )
                            {
                                if ( value == null || value == "" )
                                {
                                    value = "NULL";
                                }
                                else
                                {
                                    if ( bool.TryParse( value, out bool selected ) )
                                    {
                                        value = selected ? "1" : "0";
                                    }
                                    else
                                    {
                                        value = "NULL";
                                    }
                                }
                            }
                            else if ( fieldInfo.IsDateTime )
                            {
                                if ( value == null || value == "" )
                                {
                                    value = "NULL";
                                }
                                else
                                {
                                    if ( DateTime.TryParse( value, out DateTime selected ) )
                                    {
                                        value = "'" + selected.ToString() + "'";
                                    }
                                    else
                                    {
                                        value = "NULL";
                                    }
                                }
                            }
                            else if ( fieldInfo.IsInteger )
                            {
                                if ( value == null || value == "" )
                                {
                                    value = "NULL";
                                }
                                else
                                {
                                    if ( !int.TryParse( value, out int selected ) )
                                    {
                                        value = "NULL";
                                    }
                                }
                            }
                            else if ( fieldInfo.IsDecimal )
                            {
                                if ( value == null || value == "" )
                                {
                                    value = "NULL";
                                }
                                else
                                {
                                    if ( !decimal.TryParse( value, out decimal selected ) )
                                    {
                                        value = "NULL";
                                    }
                                }
                            }
                            else
                            {
                                if ( value == null )
                                {
                                    value = "";
                                }
                                value = "'" + value.Replace( "'", "''" ) + "'";
                            }

                            insertValues.Append( value );
                        }
                    }

                    string insertStatement = insertPrefix.ToString() + insertValues.ToString() + " )" + Environment.NewLine;

                    using ( SqlCommand querySaveStaff = new SqlCommand( insertStatement, connection ) )
                    {
                        querySaveStaff.ExecuteNonQuery();
                    }
                }

                connection.Close();
            }
        }

        private class CsvFieldInfo
        {
            public string Name { get; set; }
            public int MaxLength { get; set; } = 0;
            public bool IsDateTime { get; set; } = true;
            public bool IsInteger { get; set; } = true;
            public bool IsDecimal { get; set; } = true;
            public bool IsBool { get; set; } = true;
        }

    }
}
