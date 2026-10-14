using Microsoft.SqlServer.TransactSql.ScriptDom;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace Flank.Ssrs
{
    /// <summary>A table, view or table-valued function a query reads.</summary>
    public sealed class QueryObject
    {
        public QueryObject(string? database, string? schema, string name)
        {
            Database = database;
            Schema = schema;
            Name = name;
        }

        /// <summary>Only set for three-part names (another database).</summary>
        public string? Database { get; }

        /// <summary>Null when the query doesn't say (the user's default schema applies).</summary>
        public string? Schema { get; }

        public string Name { get; }

        public override string ToString() =>
            string.Join(".", new[] { Database, Schema, Name }.Where(p => p != null));
    }

    /// <summary>
    /// Finds the objects a query reads by parsing it with SQL Server's own T-SQL
    /// parser: every table, view and table-valued function named in FROM,
    /// JOIN, subqueries, etc. CTE names, temp tables (#t) and table variables
    /// (@t) are left out. Names built at run time (dynamic SQL) can't be seen.
    /// </summary>
    public static class QueryObjects
    {
        public static IReadOnlyList<QueryObject> Find(string sql)
        {
            var parser = new TSql160Parser(initialQuotedIdentifiers: true);
            TSqlFragment fragment;

            using (var reader = new StringReader(sql))
            {
                fragment = parser.Parse(reader, out IList<ParseError> errors);

                if (errors.Count > 0)
                {
                    var first = errors[0];
                    throw new InvalidOperationException(
                        $"Couldn't parse the query (line {first.Line}): {first.Message}");
                }
            }

            var visitor = new ObjectVisitor();
            fragment.Accept(visitor);

            return visitor.Objects
                .Where(o => o.Schema != null || !visitor.CteNames.Contains(o.Name))
                .GroupBy(o => o.ToString(), StringComparer.OrdinalIgnoreCase)
                .Select(g => g.First())
                .ToList();
        }

        /// <summary>
        /// Resolves found objects in the connection's database, as [schema].[name]:
        /// fills in the default schema, and leaves out objects in other
        /// databases or that don't exist (e.g. a CTE the parser couldn't tell apart).
        /// </summary>
        public static IReadOnlyList<string> Resolve(IDbConnection connection, IReadOnlyList<QueryObject> objects)
        {
            var resolved = new List<string>();

            foreach (var obj in objects)
            {
                if (obj.Database != null &&
                    !string.Equals(obj.Database, connection.Database, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                string name = obj.Schema == null
                    ? AccessScript.QuoteName(obj.Name)
                    : AccessScript.QuoteName(obj.Schema) + "." + AccessScript.QuoteName(obj.Name);

                using (var command = connection.CreateCommand())
                {
                    // U = table, V = view, IF/TF = table-valued functions.
                    command.CommandText = @"
SELECT QUOTENAME(OBJECT_SCHEMA_NAME(o.object_id)) + '.' + QUOTENAME(o.name)
FROM sys.objects o
WHERE o.object_id = OBJECT_ID(@Name)
  AND o.type IN ('U', 'V', 'IF', 'TF');";

                    var parameter = command.CreateParameter();
                    parameter.ParameterName = "@Name";
                    parameter.Value = name;
                    command.Parameters.Add(parameter);

                    if (command.ExecuteScalar() is string full && !resolved.Contains(full))
                        resolved.Add(full);
                }
            }

            return resolved;
        }

        private sealed class ObjectVisitor : TSqlFragmentVisitor
        {
            public List<QueryObject> Objects { get; } = new List<QueryObject>();

            public HashSet<string> CteNames { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public override void Visit(CommonTableExpression node)
            {
                CteNames.Add(node.ExpressionName.Value);
                base.Visit(node);
            }

            public override void Visit(NamedTableReference node)
            {
                Add(node.SchemaObject);
                base.Visit(node);
            }

            public override void Visit(SchemaObjectFunctionTableReference node)
            {
                Add(node.SchemaObject);
                base.Visit(node);
            }

            private void Add(SchemaObjectName name)
            {
                string? baseName = name.BaseIdentifier?.Value;

                // #temp tables live in tempdb; linked servers are out of scope.
                if (baseName == null || baseName.StartsWith("#") || name.ServerIdentifier != null)
                    return;

                Objects.Add(new QueryObject(
                    name.DatabaseIdentifier?.Value,
                    name.SchemaIdentifier?.Value,
                    baseName));
            }
        }
    }
}
