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

    private static bool created;

    public static void EnsureCreated()
    {
        if (created)
            return;

        using var connection = TestSettings.OpenSqlConnection();
        using var command = connection.CreateCommand();
        command.CommandText = ParametersDefinition;
        command.ExecuteNonQuery();
        created = true;
    }
}
