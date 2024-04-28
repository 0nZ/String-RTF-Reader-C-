using Microsoft.Data.Sqlite;
using Windows.Storage;
using System.IO;
using System.Windows.Controls;
using System.Data;
using DocumentFormat.OpenXml.Office.Word;
using System.Data.SqlClient;


namespace WpfApp1
{
    public static class DataAccess
    {

        public async static void InitializeDatabase()
        {
            await ApplicationData.Current.LocalFolder
                    .CreateFileAsync("database.db", CreationCollisionOption.OpenIfExists);
            string dbpath = Path.Combine(ApplicationData.Current.LocalFolder.Path,
                                         "database.db");
            using (var db = new SqliteConnection($"Filename={dbpath}"))
            {
                db.Open();

                string tableCommand = "CREATE TABLE IF NOT " +
                    "EXISTS FILES (nr_sequencia INTEGER PRIMARY KEY ON CONFLICT ROLLBACK AUTOINCREMENT UNIQUE ON CONFLICT ROLLBACK NOT NULL ON CONFLICT ROLLBACK, " +
                    "titulo TEXT NOT NULL ON CONFLICT ROLLBACK, " +
                    "caminho_arquivo TEXT NOT NULL ON CONFLICT ROLLBACK);";

                var createTable = new SqliteCommand(tableCommand, db);

                createTable.ExecuteReader();

                db.Close();
            }
        }

        public static void AddData(string titulo, string rtf)
        {
            
            string dbpath = Path.Combine(ApplicationData.Current.LocalFolder.Path,
                                         "database.db");
            using (var db = new SqliteConnection($"Filename={dbpath}"))
            {
                db.Open();

                var insertCommand = new SqliteCommand();
                insertCommand.Connection = db;

                // Use parameterized query to prevent SQL injection attacks
                insertCommand.CommandText = "INSERT INTO FILES VALUES (NULL, @Entry, @Entry2);";
                insertCommand.Parameters.AddWithValue("@Entry", titulo);
                insertCommand.Parameters.AddWithValue("@Entry2", rtf);

                insertCommand.ExecuteReader();

                db.Close();
            }

        }

        public static List<string> GetData()
        {

            var entries = new List<string>();
            
            string dbpath = Path.Combine(ApplicationData.Current.LocalFolder.Path,
                                         "database.db");
            using (var db = new SqliteConnection($"Filename={dbpath}"))
            {
                
                db.Open();

                var selectCommand = new SqliteCommand
                    ("SELECT nr_sequencia, titulo, caminho_arquivo from FILES", db);

                SqliteDataReader query = selectCommand.ExecuteReader();

                

                while (query.Read())
                {
                    entries.Add(query.GetString(2));
                }

                db.Close();

            }

            return entries;
        }

        public static void DeleteData()
        {
            string dbpath = Path.Combine(ApplicationData.Current.LocalFolder.Path,
                                         "database.db");
            using (var db = new SqliteConnection($"Filename={dbpath}"))
            {
                db.Open();

                var insertCommand = new SqliteCommand();
                insertCommand.Connection = db;

                // Use parameterized query to prevent SQL injection attacks
                insertCommand.CommandText = "delete from FILES;";

                insertCommand.ExecuteReader();

                db.Close();
            }
        }

    }
}
