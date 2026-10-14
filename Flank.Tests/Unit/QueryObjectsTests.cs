using Flank.Ssrs;

namespace Flank.Tests.Unit;

public class QueryObjectsTests
{
    private static List<string> Find(string sql) =>
        QueryObjects.Find(sql).Select(o => o.ToString()).ToList();

    [Fact]
    public void Finds_tables_in_from_and_joins()
    {
        Assert.Equal(
            new[] { "dbo.orders", "customers" },
            Find("SELECT o.id, c.name FROM dbo.orders o JOIN customers c ON c.id = o.customer_id"));
    }

    [Fact]
    public void Finds_tables_used_only_in_where_subqueries()
    {
        Assert.Equal(
            new[] { "leagues", "sales.banned" },
            Find("SELECT name FROM leagues WHERE id NOT IN (SELECT league_id FROM sales.banned)"));
    }

    [Fact]
    public void Leaves_out_ctes_temp_tables_and_table_variables()
    {
        string sql = @"
DECLARE @t TABLE (id int);
WITH recent AS (SELECT id FROM dbo.orders)
SELECT r.id FROM recent r
JOIN #scratch s ON s.id = r.id
JOIN @t t ON t.id = r.id";

        Assert.Equal(new[] { "dbo.orders" }, Find(sql));
    }

    [Fact]
    public void Finds_table_valued_functions_and_other_databases()
    {
        Assert.Equal(
            new[] { "dbo.recent_orders", "archive.dbo.orders" },
            Find("SELECT * FROM dbo.recent_orders(7) UNION ALL SELECT * FROM archive.dbo.orders"));
    }

    [Fact]
    public void Each_object_once()
    {
        Assert.Equal(new[] { "dbo.t" }, Find("SELECT * FROM dbo.t a JOIN dbo.t b ON a.id = b.parent_id"));
    }

    [Fact]
    public void Quoted_names_keep_their_text()
    {
        Assert.Equal(new[] { "dbo.order details" }, Find("SELECT * FROM [dbo].[order details]"));
    }

    [Fact]
    public void Reports_queries_it_cannot_parse()
    {
        Assert.Throws<InvalidOperationException>(() => QueryObjects.Find("SELECT FROM WHERE"));
    }
}
