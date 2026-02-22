using System;
using System.Collections.Generic;
using System.Data;
using System.Threading.Tasks;
using Microsoft.Data.SqlClient;
using Microsoft.Extensions.Configuration;

namespace BlazorDemos.Shared
{
    public class QueryResult
    {
        public List<string> Columns { get; set; } = new List<string>();
        public List<Dictionary<string, object>> Rows { get; set; } = new List<Dictionary<string, object>>();
        public string ErrorMessage { get; set; }
        public bool HasError => !string.IsNullOrEmpty(ErrorMessage);
        public int RowsAffected { get; set; }
    }

    public class SqlQueryService
    {
        private readonly string _connectionString;

        public SqlQueryService(IConfiguration configuration)
        {
            _connectionString = configuration.GetConnectionString("SqlServer");
        }

        public async Task<QueryResult> ExecuteQueryAsync(string sql)
        {
            var result = new QueryResult();

            if (string.IsNullOrWhiteSpace(sql))
            {
                result.ErrorMessage = "Query cannot be empty.";
                return result;
            }

            try
            {
                using var connection = new SqlConnection(_connectionString);
                await connection.OpenAsync();

                using var command = new SqlCommand(sql, connection);
                command.CommandTimeout = 30;

                // Determine whether this is a SELECT or a non-query statement
                var trimmed = sql.TrimStart();
                bool isSelect = trimmed.StartsWith("SELECT", StringComparison.OrdinalIgnoreCase)
                             || trimmed.StartsWith("WITH", StringComparison.OrdinalIgnoreCase)
                             || trimmed.StartsWith("EXEC", StringComparison.OrdinalIgnoreCase)
                             || trimmed.StartsWith("EXECUTE", StringComparison.OrdinalIgnoreCase);

                if (isSelect)
                {
                    using var reader = await command.ExecuteReaderAsync();

                    for (int i = 0; i < reader.FieldCount; i++)
                    {
                        result.Columns.Add(reader.GetName(i));
                    }

                    while (await reader.ReadAsync())
                    {
                        var row = new Dictionary<string, object>();
                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            row[result.Columns[i]] = reader.IsDBNull(i) ? null : reader.GetValue(i);
                        }
                        result.Rows.Add(row);
                    }
                }
                else
                {
                    result.RowsAffected = await command.ExecuteNonQueryAsync();
                }
            }
            catch (SqlException ex)
            {
                result.ErrorMessage = $"SQL Error {ex.Number}: {ex.Message}";
            }
            catch (Exception ex)
            {
                result.ErrorMessage = $"Error: {ex.Message}";
            }

            return result;
        }

        public async Task<bool> TestConnectionAsync()
        {
            try
            {
                using var connection = new SqlConnection(_connectionString);
                await connection.OpenAsync();
                return true;
            }
            catch
            {
                return false;
            }
        }
    }
}
