using MySqlConnector;

namespace CSharpFlexGrid
{
    public class DatabaseHelper
    {
        private readonly string _connectionString;

        public DatabaseHelper()
        {
            DotNetEnv.Env.Load();

            var host = Environment.GetEnvironmentVariable("DB_HOST") ?? "localhost";
            var user = Environment.GetEnvironmentVariable("DB_USER") ?? "root";
            var password = Environment.GetEnvironmentVariable("DB_PASSWORD") ?? "";
            var database = Environment.GetEnvironmentVariable("DB_NAME") ?? "";
            var port = Environment.GetEnvironmentVariable("DB_PORT") ?? "3306";

            _connectionString = $"Server={host};Port={port};Database={database};User={user};Password={password};";
        }

        /// <summary>
        /// Executes a SQL query and returns the first column as a list of strings.
        /// Useful for populating ComboBox items.
        /// </summary>
        /// <param name="sql">The SQL query to execute</param>
        /// <returns>List of strings from the first column of the result</returns>
        public List<string> GetComboBoxList(string sql)
        {
            var results = new List<string>();

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                connection.Open();

                using var command = new MySqlCommand(sql, connection);
                using var reader = command.ExecuteReader();

                while (reader.Read())
                {
                    if (!reader.IsDBNull(0))
                    {
                        var value = reader.GetValue(0).ToString();
                        if (!string.IsNullOrWhiteSpace(value))
                        {
                            results.Add(value);
                        }
                    }
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show($"Database error: {ex.Message}", "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return results;
        }

        /// <summary>
        /// Tests the database connection.
        /// </summary>
        /// <returns>True if connection is successful, false otherwise</returns>
        public bool TestConnection()
        {
            try
            {
                using var connection = new MySqlConnection(_connectionString);
                connection.Open();
                return true;
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Connection failed: {ex.Message}", "Connection Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }
        }

        /// <summary>
        /// Executes a non-query SQL command (INSERT, UPDATE, DELETE).
        /// </summary>
        /// <param name="sql">The SQL command to execute</param>
        /// <returns>Number of rows affected</returns>
        public int ExecuteNonQuery(string sql)
        {
            int rowsAffected = 0;

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                connection.Open();

                using var command = new MySqlCommand(sql, connection);
                rowsAffected = command.ExecuteNonQuery();
            }
            catch (MySqlException ex)
            {
                MessageBox.Show($"Database error: {ex.Message}", "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return rowsAffected;
        }

        /// <summary>
        /// Executes an INSERT and returns the last inserted ID.
        /// Uses single connection to ensure LAST_INSERT_ID() works correctly.
        /// </summary>
        /// <param name="sql">The INSERT SQL command</param>
        /// <returns>The last inserted ID, or 0 if failed</returns>
        public int ExecuteInsertAndGetId(string sql)
        {
            int insertedId = 0;

            try
            {
                using var connection = new MySqlConnection(_connectionString);
                connection.Open();

                using var insertCommand = new MySqlCommand(sql, connection);
                insertCommand.ExecuteNonQuery();

                using var idCommand = new MySqlCommand("SELECT LAST_INSERT_ID()", connection);
                var result = idCommand.ExecuteScalar();
                if (result != null && int.TryParse(result.ToString(), out int id))
                {
                    insertedId = id;
                }
            }
            catch (MySqlException ex)
            {
                MessageBox.Show($"Database error: {ex.Message}", "Database Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            return insertedId;
        }
    }
}
