using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;

namespace Flank.Ssrs
{
    /// <summary>One result set from a procedure run.</summary>
    public sealed class ResultSet
    {
        public ResultSet(IReadOnlyList<string> columns, IReadOnlyList<object?[]> rows, bool truncated)
        {
            Columns = columns;
            Rows = rows;
            Truncated = truncated;
        }

        public IReadOnlyList<string> Columns { get; }

        /// <summary>Values in column order; database NULLs are null.</summary>
        public IReadOnlyList<object?[]> Rows { get; }

        /// <summary>True if there were more rows than the run kept.</summary>
        public bool Truncated { get; }
    }

    /// <summary>Runs a stored procedure and reads every result set it returns.</summary>
    public static class StoredProcedureRunner
    {
        /// <param name="parameters">
        /// Values by name (with '@'). A null value leaves the parameter out, so
        /// the procedure's default applies.
        /// </param>
        /// <param name="maxRows">Rows kept per result set.</param>
        /// <param name="outputParameters">
        /// The procedure's OUTPUT parameters; passed so the call doesn't fail
        /// for lack of them. Their values aren't returned.
        /// </param>
        public static IReadOnlyList<ResultSet> Run(
            IDbConnection connection,
            string procedureName,
            IReadOnlyList<KeyValuePair<string, object?>> parameters,
            int maxRows,
            IReadOnlyList<SprocParameter>? outputParameters = null,
            int timeoutSeconds = 120)
        {
            var results = new List<ResultSet>();

            using (var command = connection.CreateCommand())
            {
                command.CommandText = procedureName;
                command.CommandType = CommandType.StoredProcedure;
                command.CommandTimeout = timeoutSeconds;

                foreach (var parameter in parameters)
                {
                    if (parameter.Value == null)
                        continue;

                    var p = command.CreateParameter();
                    p.ParameterName = parameter.Key;
                    p.Value = parameter.Value;
                    command.Parameters.Add(p);
                }

                foreach (var output in outputParameters ?? Array.Empty<SprocParameter>())
                {
                    var p = command.CreateParameter();
                    p.ParameterName = output.Name;
                    p.Direction = ParameterDirection.InputOutput;
                    p.Value = DBNull.Value;
                    SetSqlType(p, output);
                    command.Parameters.Add(p);
                }

                using (var reader = command.ExecuteReader())
                {
                    do
                    {
                        // Statements that don't return rows (e.g. SET, UPDATE) still
                        // produce a "result" with no columns; skip those.
                        if (reader.FieldCount == 0)
                            continue;

                        var columns = new List<string>();
                        for (int i = 0; i < reader.FieldCount; i++)
                            columns.Add(reader.GetName(i));

                        var rows = new List<object?[]>();
                        bool truncated = false;

                        while (reader.Read())
                        {
                            if (rows.Count >= maxRows)
                            {
                                truncated = true;
                                break;
                            }

                            var values = new object?[reader.FieldCount];
                            reader.GetValues(values!);

                            for (int i = 0; i < values.Length; i++)
                            {
                                if (values[i] == DBNull.Value)
                                    values[i] = null;
                            }

                            rows.Add(values);
                        }

                        results.Add(new ResultSet(columns, rows, truncated));
                    }
                    while (reader.NextResult());
                }
            }

            return results;
        }

        // Output parameters need their real type and size, or SQL Server
        // truncates or rejects the value it writes back. Uses the generic
        // DbType: SSMS hands out Microsoft.Data.SqlClient connections, not
        // System.Data.SqlClient ones.
        private static void SetSqlType(IDbDataParameter parameter, SprocParameter source)
        {
            switch ((source.SqlType ?? "").ToLowerInvariant())
            {
                case "bit": parameter.DbType = DbType.Boolean; break;
                case "tinyint": parameter.DbType = DbType.Byte; break;
                case "smallint": parameter.DbType = DbType.Int16; break;
                case "int": parameter.DbType = DbType.Int32; break;
                case "bigint": parameter.DbType = DbType.Int64; break;
                case "real": parameter.DbType = DbType.Single; break;
                case "float": parameter.DbType = DbType.Double; break;
                case "date": parameter.DbType = DbType.Date; break;
                case "time": parameter.DbType = DbType.Time; break;
                case "datetime":
                case "smalldatetime": parameter.DbType = DbType.DateTime; break;
                case "datetime2": parameter.DbType = DbType.DateTime2; break;
                case "datetimeoffset": parameter.DbType = DbType.DateTimeOffset; break;
                case "uniqueidentifier": parameter.DbType = DbType.Guid; break;

                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                    parameter.DbType = DbType.Decimal;
                    parameter.Precision = source.Precision;
                    parameter.Scale = source.Scale;
                    break;

                case "char":
                case "varchar":
                    parameter.DbType = DbType.AnsiString;
                    parameter.Size = source.MaxLength < 0 ? -1 : source.MaxLength;
                    break;

                case "binary":
                case "varbinary":
                    parameter.DbType = DbType.Binary;
                    parameter.Size = source.MaxLength < 0 ? -1 : source.MaxLength;
                    break;

                default:
                    // nchar, nvarchar and anything else: Unicode text.
                    parameter.DbType = DbType.String;
                    parameter.Size = source.MaxLength < 0 ? -1 : Math.Max(source.MaxLength / 2, 1);
                    break;
            }
        }

        /// <summary>
        /// Converts typed-in text to the parameter's SQL type where it's
        /// unambiguous (numbers, bit, dates), otherwise keeps the text and lets
        /// SQL Server convert it. Blank text means "use the default" (null).
        /// </summary>
        public static object? ParseValue(string sqlType, string text)
        {
            if (string.IsNullOrWhiteSpace(text))
                return null;

            text = text.Trim();
            var culture = CultureInfo.CurrentCulture;

            switch ((sqlType ?? "").ToLowerInvariant())
            {
                case "tinyint":
                case "smallint":
                case "int":
                case "bigint":
                    if (long.TryParse(text, NumberStyles.Integer, culture, out long whole))
                        return whole;
                    break;

                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                    if (decimal.TryParse(text, NumberStyles.Number, culture, out decimal number))
                        return number;
                    break;

                case "float":
                case "real":
                    if (double.TryParse(text, NumberStyles.Float, culture, out double real))
                        return real;
                    break;

                case "bit":
                    if (text == "1" || text.Equals("true", StringComparison.OrdinalIgnoreCase))
                        return true;
                    if (text == "0" || text.Equals("false", StringComparison.OrdinalIgnoreCase))
                        return false;
                    break;

                case "date":
                case "datetime":
                case "datetime2":
                case "smalldatetime":
                    if (DateTime.TryParse(text, culture, DateTimeStyles.None, out DateTime date))
                        return date;
                    break;
            }

            return text;
        }
    }
}
