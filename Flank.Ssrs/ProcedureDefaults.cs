using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>
    /// Reads parameter defaults from a procedure's definition. SQL Server only
    /// records defaults for CLR procedures (sys.parameters.has_default_value),
    /// so for T-SQL ones the CREATE PROCEDURE text is the only source.
    /// </summary>
    public static class ProcedureDefaults
    {
        public sealed class Default
        {
            public Default(string? value, string text)
            {
                Value = value;
                Text = text;
            }

            /// <summary>The literal as a value (5, abc, 2024-01-01); null for NULL or an expression.</summary>
            public string? Value { get; }

            /// <summary>As written in the definition: 5, N'abc', NULL, GETDATE().</summary>
            public string Text { get; }
        }

        /// <summary>
        /// Parameter name (with '@', case-insensitive) to its default, for the
        /// parameters that have one. Empty if the definition can't be parsed
        /// (e.g. an encrypted procedure, whose definition is null).
        /// </summary>
        public static IReadOnlyDictionary<string, Default> Parse(string? definition)
        {
            var defaults = new Dictionary<string, Default>(StringComparer.OrdinalIgnoreCase);

            if (string.IsNullOrWhiteSpace(definition))
                return defaults;

            var parser = new TSql160Parser(initialQuotedIdentifiers: true);
            TSqlFragment fragment;

            using (var reader = new StringReader(definition))
            {
                fragment = parser.Parse(reader, out IList<ParseError> errors);

                if (errors.Count > 0)
                    return defaults;
            }

            var procedure = (fragment as TSqlScript)?.Batches
                .SelectMany(b => b.Statements)
                .OfType<ProcedureStatementBodyBase>()
                .FirstOrDefault();

            if (procedure == null)
                return defaults;

            foreach (var parameter in procedure.Parameters)
            {
                if (parameter.Value == null)
                    continue;

                string text = string.Concat(
                    Enumerable.Range(parameter.Value.FirstTokenIndex, parameter.Value.LastTokenIndex - parameter.Value.FirstTokenIndex + 1)
                        .Select(i => parameter.Value.ScriptTokenStream[i].Text));

                defaults[parameter.VariableName.Value] = new Default(LiteralValue(parameter.Value), text.Trim());
            }

            return defaults;
        }

        private static string? LiteralValue(ScalarExpression expression)
        {
            switch (expression)
            {
                case NullLiteral _:
                    return null;

                case Literal literal:
                    return literal.Value;

                // -1
                case UnaryExpression unary when unary.UnaryExpressionType == UnaryExpressionType.Negative &&
                                                unary.Expression is Literal negated:
                    return "-" + negated.Value;

                default:
                    return null;
            }
        }
    }
}
