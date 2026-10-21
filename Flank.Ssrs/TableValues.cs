using System;
using System.Collections.Generic;
using System.Globalization;
using System.Text;

namespace Flank.Ssrs
{
    /// <summary>
    /// The rows of a table-valued parameter as the JSON the generated call
    /// reads with OPENJSON (see <see cref="ProcedureCall.BuildExec(string, IReadOnlyList{SprocParameter})"/>):
    /// an array of rows, each an array of column values in the table type's order.
    /// </summary>
    public static class TableValues
    {
        /// <summary>[[1,"a"],[2,null]]. Null for no rows (an empty table).</summary>
        public static string? ToJson(IReadOnlyList<IReadOnlyList<object?>>? rows)
        {
            if (rows == null || rows.Count == 0)
                return null;

            var json = new StringBuilder("[");

            for (int r = 0; r < rows.Count; r++)
            {
                if (r > 0)
                    json.Append(',');

                json.Append('[');

                for (int c = 0; c < rows[r].Count; c++)
                {
                    if (c > 0)
                        json.Append(',');

                    AppendValue(json, rows[r][c]);
                }

                json.Append(']');
            }

            return json.Append(']').ToString();
        }

        private static void AppendValue(StringBuilder json, object? value)
        {
            switch (value)
            {
                case null:
                case DBNull _:
                    json.Append("null");
                    break;

                case bool flag:
                    json.Append(flag ? "true" : "false");
                    break;

                case double real:
                    json.Append(real.ToString("R", CultureInfo.InvariantCulture));
                    break;

                case float single:
                    json.Append(single.ToString("R", CultureInfo.InvariantCulture));
                    break;

                case byte _:
                case short _:
                case int _:
                case long _:
                case decimal _:
                    json.Append(Convert.ToString(value, CultureInfo.InvariantCulture));
                    break;

                // ISO 8601, which OPENJSON converts to any date or time type.
                case DateTime date:
                    AppendString(json, date.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFF", CultureInfo.InvariantCulture).TrimEnd('.'));
                    break;

                case DateTimeOffset offset:
                    AppendString(json, offset.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFFFzzz", CultureInfo.InvariantCulture));
                    break;

                case TimeSpan time:
                    AppendString(json, time.ToString("c", CultureInfo.InvariantCulture));
                    break;

                default:
                    AppendString(json, Convert.ToString(value, CultureInfo.InvariantCulture) ?? "");
                    break;
            }
        }

        private static void AppendString(StringBuilder json, string text)
        {
            json.Append('"');

            foreach (char c in text)
            {
                switch (c)
                {
                    case '"': json.Append("\\\""); break;
                    case '\\': json.Append("\\\\"); break;
                    case '\n': json.Append("\\n"); break;
                    case '\r': json.Append("\\r"); break;
                    case '\t': json.Append("\\t"); break;
                    default:
                        if (c < ' ')
                            json.Append("\\u").Append(((int)c).ToString("x4", CultureInfo.InvariantCulture));
                        else
                            json.Append(c);
                        break;
                }
            }

            json.Append('"');
        }
    }
}
