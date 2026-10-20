using Flank.Excel;
using Flank.Ssrs;
using System.Collections.Generic;
using System.Linq;

namespace Flank.SsmsExtension
{
    /// <summary>Reviewed procedure parameters as workbook parameter cells, with checks for their SQL types.</summary>
    internal static class WorkbookParameters
    {
        public static IReadOnlyList<WorkbookParameter> From(IReadOnlyList<ParameterSettings> settings) =>
            settings.Select(From).ToList();

        private static WorkbookParameter From(ParameterSettings setting)
        {
            var parameter = setting.Parameter;

            var cell = new WorkbookParameter(setting.SqlName.TrimStart('@'), setting.DisplayName)
            {
                Note = setting.Note,
                Required = setting.IfBlank == BlankValue.Required,
            };

            if (setting.HasList)
            {
                cell.List = new WorkbookList(setting.ListSql, setting.ListValueColumn!, setting.ListLabelColumn!)
                {
                    Rows = setting.ListRows.Select(row => (row.Key, row.Value)).ToList(),
                };
            }

            switch ((parameter.SqlType ?? "").ToLowerInvariant())
            {
                case "tinyint":
                    WholeNumber(cell, 0, byte.MaxValue);
                    break;
                case "smallint":
                    WholeNumber(cell, short.MinValue, short.MaxValue);
                    break;
                case "int":
                    WholeNumber(cell, int.MinValue, int.MaxValue);
                    break;
                case "bigint":
                    // Excel numbers are doubles: exact only up to 2^53.
                    WholeNumber(cell, -9007199254740991L, 9007199254740991L);
                    break;

                case "decimal":
                case "numeric":
                case "money":
                case "smallmoney":
                case "float":
                case "real":
                    cell.Kind = WorkbookParameterKind.Decimal;
                    break;

                case "date":
                    cell.Kind = WorkbookParameterKind.Date;
                    break;

                case "datetime":
                case "datetime2":
                case "smalldatetime":
                case "datetimeoffset":
                    cell.Kind = WorkbookParameterKind.DateTime;
                    break;

                case "bit":
                    cell.Kind = WorkbookParameterKind.TrueFalse;
                    break;

                case "char":
                case "varchar":
                    cell.Kind = WorkbookParameterKind.Text;
                    cell.MaxLength = parameter.MaxLength > 0 ? parameter.MaxLength : (int?)null;
                    break;

                case "nchar":
                case "nvarchar":
                    // max_length is in bytes: two per character; -1 for max.
                    cell.Kind = WorkbookParameterKind.Text;
                    cell.MaxLength = parameter.MaxLength > 0 ? parameter.MaxLength / 2 : (int?)null;
                    break;

                case "text":
                case "ntext":
                    cell.Kind = WorkbookParameterKind.Text;
                    break;
            }

            return cell;
        }

        private static void WholeNumber(WorkbookParameter cell, long minimum, long maximum)
        {
            cell.Kind = WorkbookParameterKind.WholeNumber;
            cell.Minimum = minimum;
            cell.Maximum = maximum;
        }
    }
}
