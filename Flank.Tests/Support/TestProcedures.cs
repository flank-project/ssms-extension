namespace Flank.Tests.Support;

/// <summary>Procedures the tests create in the test database (idempotently).</summary>
internal static class TestProcedures
{
    /// <summary>
    /// One required parameter and defaults of several types. Returns its teams
    /// and echoes the values it received, so tests can see which applied.
    /// </summary>
    public const string Parameters = "dbo.flank_test_params";

    private const string ParametersDefinition = @"
CREATE OR ALTER PROCEDURE dbo.flank_test_params
    @League_Id int,
    @Name_Like nvarchar(50) = N'%',
    @Max_Id int = NULL,
    @Active bit = 1,
    @Since date = '2000-01-01'
AS
BEGIN
    SET NOCOUNT ON;
    SELECT t.team_name, t.team_id, t.league_id,
           @Name_Like AS name_like, @Max_Id AS max_id, @Active AS active, @Since AS since
    FROM dbo.sgm_teams AS t
    WHERE t.league_id = @League_Id
      AND t.team_name LIKE @Name_Like
      AND (@Max_Id IS NULL OR t.team_id <= @Max_Id);
END";

    /// <summary>
    /// Takes a table of rows (several columns) and a league; returns the rows
    /// it received, so tests can see what the JSON call passed.
    /// </summary>
    public const string TableRows = "dbo.flank_test_table_rows";

    /// <summary>Takes a one-column table of team ids; returns those teams.</summary>
    public const string TableIds = "dbo.flank_test_table_ids";

    // Table types can't be altered, only created once.
    private static readonly string[] TableDefinitions =
    {
        @"
IF TYPE_ID(N'dbo.flank_test_rows') IS NULL
    CREATE TYPE dbo.flank_test_rows AS TABLE (
        seq int IDENTITY,
        team_id int NOT NULL,
        label nvarchar(50) NULL,
        active bit NULL,
        since date NULL,
        amount decimal(9, 2) NULL);",
        @"
IF TYPE_ID(N'dbo.flank_test_ids') IS NULL
    CREATE TYPE dbo.flank_test_ids AS TABLE (team_id int NOT NULL);",
        @"
CREATE OR ALTER PROCEDURE dbo.flank_test_table_rows
    @League_Id int,
    @Rows dbo.flank_test_rows READONLY
AS
BEGIN
    SET NOCOUNT ON;
    SELECT @League_Id AS league_id, r.seq, r.team_id, r.label, r.active, r.since, r.amount
    FROM @Rows AS r
    ORDER BY r.seq;
END",
        @"
CREATE OR ALTER PROCEDURE dbo.flank_test_table_ids
    @Ids dbo.flank_test_ids READONLY
AS
BEGIN
    SET NOCOUNT ON;
    SELECT t.team_id, t.team_name
    FROM dbo.sgm_teams AS t
    WHERE t.team_id IN (SELECT team_id FROM @Ids);
END",
    };

    private static bool created;

    public static void EnsureCreated()
    {
        if (created)
            return;

        using var connection = TestSettings.OpenSqlConnection();

        foreach (string definition in TableDefinitions.Prepend(ParametersDefinition))
        {
            using var command = connection.CreateCommand();
            command.CommandText = definition;
            command.ExecuteNonQuery();
        }

        created = true;
    }
}
