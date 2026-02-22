using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace BlazorDemos.Services
{
    /// <summary>
    /// Service for executing SQL Server queries.
    /// Register in Startup.cs with services.AddScoped&lt;SqlQueryService&gt;().
    /// Configure the connection string named "SqlServer" in appsettings.json.
    /// </summary>
    public class SqlQueryService
    {
        private readonly string _connectionString;

        public SqlQueryService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlServer")
                ?? throw new InvalidOperationException(
                    "Connection string 'SqlServer' not found in configuration. " +
                    "Add a 'SqlServer' entry under 'ConnectionStrings' in appsettings.json.");
        }

        /// <summary>
        /// Executes a SELECT query and returns results as a DataTable.
        /// Use parameterized queries (SqlParameter) to prevent SQL injection.
        /// </summary>
        /// <param name="sql">The SQL query to execute (use @param placeholders).</param>
        /// <param name="parameters">Optional parameters matching placeholders in the query.</param>
        public async Task<DataTable> ExecuteQueryAsync(string sql, SqlParameter[] parameters = null)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var command = new SqlCommand(sql, connection);
            if (parameters != null)
                command.Parameters.AddRange(parameters);

            var table = new DataTable();
            using var reader = await command.ExecuteReaderAsync();
            table.Load(reader);
            return table;
        }

        /// <summary>
        /// Executes a SELECT query and returns each row as a dictionary of column name to value.
        /// Use parameterized queries (SqlParameter) to prevent SQL injection.
        /// </summary>
        /// <param name="sql">The SQL query to execute (use @param placeholders).</param>
        /// <param name="parameters">Optional parameters matching placeholders in the query.</param>
        public async Task<List<Dictionary<string, object>>> ExecuteQueryAsDictionaryAsync(
            string sql, SqlParameter[] parameters = null)
        {
            var results = new List<Dictionary<string, object>>();

            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var command = new SqlCommand(sql, connection);
            if (parameters != null)
                command.Parameters.AddRange(parameters);

            using var reader = await command.ExecuteReaderAsync();
            while (await reader.ReadAsync())
            {
                var row = new Dictionary<string, object>();
                for (int i = 0; i < reader.FieldCount; i++)
                    row[reader.GetName(i)] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                results.Add(row);
            }

            return results;
        }

        /// <summary>
        /// Executes an INSERT, UPDATE, or DELETE statement.
        /// Use parameterized queries (SqlParameter) to prevent SQL injection.
        /// </summary>
        /// <param name="sql">The SQL statement to execute (use @param placeholders).</param>
        /// <param name="parameters">Optional parameters matching placeholders in the statement.</param>
        /// <returns>The number of rows affected.</returns>
        public async Task<int> ExecuteNonQueryAsync(string sql, SqlParameter[] parameters = null)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var command = new SqlCommand(sql, connection);
            if (parameters != null)
                command.Parameters.AddRange(parameters);

            return await command.ExecuteNonQueryAsync();
        }

        /// <summary>
        /// Executes a query and returns a single scalar value.
        /// Use parameterized queries (SqlParameter) to prevent SQL injection.
        /// </summary>
        /// <param name="sql">The SQL query to execute (use @param placeholders).</param>
        /// <param name="parameters">Optional parameters matching placeholders in the query.</param>
        public async Task<object> ExecuteScalarAsync(string sql, SqlParameter[] parameters = null)
        {
            using var connection = new SqlConnection(_connectionString);
            await connection.OpenAsync();
            using var command = new SqlCommand(sql, connection);
            if (parameters != null)
                command.Parameters.AddRange(parameters);

            return await command.ExecuteScalarAsync();
        }
    }
}
