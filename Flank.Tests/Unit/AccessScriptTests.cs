using Flank.Ssrs;

namespace Flank.Tests.Unit;

public class AccessScriptTests
{
    private static AccessRequest Request(
        AccessSignIn signIn = AccessSignIn.SqlLogin,
        bool azure = true,
        AccessScope scope = AccessScope.Procedure) => new AccessRequest
    {
        SignIn = signIn,
        Name = "report_user",
        Password = "Secret'1!",
        Database = "sales",
        IsAzureSqlDatabase = azure,
        Scope = scope,
        Procedure = "[dbo].[get_orders]",
        Objects = new[] { "[dbo].[orders]", "[dbo].[customers]" },
    };

    [Fact]
    public void Azure_sql_login_is_a_contained_user()
    {
        var script = AccessScript.Build(Request());

        Assert.Equal(
            "IF DATABASE_PRINCIPAL_ID(N'report_user') IS NULL\n    CREATE USER [report_user] WITH PASSWORD = N'Secret''1!';",
            script[0]);
        Assert.DoesNotContain(script, s => s.Contains("CREATE LOGIN"));
    }

    [Fact]
    public void Sql_server_login_creates_a_login_then_a_user()
    {
        var script = AccessScript.Build(Request(azure: false));

        Assert.Contains("CREATE LOGIN [report_user] WITH PASSWORD = N'Secret''1!', DEFAULT_DATABASE = [sales];", script[0]);
        Assert.Contains("CREATE USER [report_user] FOR LOGIN [report_user];", script[1]);
    }

    [Fact]
    public void Entra_user_comes_from_the_external_provider()
    {
        var request = Request(AccessSignIn.Entra);
        request.Name = "jane@contoso.com";

        Assert.Equal(
            "IF DATABASE_PRINCIPAL_ID(N'jane@contoso.com') IS NULL\n    CREATE USER [jane@contoso.com] FROM EXTERNAL PROVIDER;",
            AccessScript.Build(request)[0]);
    }

    [Fact]
    public void Windows_user_gets_a_windows_login()
    {
        var request = Request(AccessSignIn.Windows, azure: false);
        request.Name = @"CONTOSO\jsmith";

        var script = AccessScript.Build(request);

        Assert.Contains(@"CREATE LOGIN [CONTOSO\jsmith] FROM WINDOWS", script[0]);
        Assert.Contains(@"CREATE USER [CONTOSO\jsmith] FOR LOGIN [CONTOSO\jsmith];", script[1]);
    }

    [Fact]
    public void Procedure_scope_grants_execute_only()
    {
        Assert.Equal("GRANT EXECUTE ON OBJECT::[dbo].[get_orders] TO [report_user];", AccessScript.Build(Request()).Last());
    }

    [Fact]
    public void Objects_scope_grants_select_on_each()
    {
        var script = AccessScript.Build(Request(scope: AccessScope.Objects));

        Assert.Equal(
            new[] { "GRANT SELECT ON OBJECT::[dbo].[orders] TO [report_user];", "GRANT SELECT ON OBJECT::[dbo].[customers] TO [report_user];" },
            script.Skip(1));
    }

    [Fact]
    public void Existing_user_only_gets_the_grant()
    {
        var request = Request();
        request.ExistingUser = true;
        request.Password = "";

        Assert.Equal(
            new[] { "GRANT EXECUTE ON OBJECT::[dbo].[get_orders] TO [report_user];" },
            AccessScript.Build(request));
    }

    [Fact]
    public void Query_can_become_a_procedure_after_the_user_is_created()
    {
        var request = Request();
        request.ProcedureBody = "SELECT id FROM dbo.orders;  ";

        var script = AccessScript.Build(request);

        // The user first (re-runnable), then the procedure, then the grant.
        Assert.StartsWith("IF DATABASE_PRINCIPAL_ID", script[0]);
        Assert.Equal(
            "CREATE PROCEDURE [dbo].[get_orders]\nAS\nBEGIN\n    SET NOCOUNT ON;\n\nSELECT id FROM dbo.orders;\nEND;",
            script[script.Count - 2]);
        Assert.Equal("GRANT EXECUTE ON OBJECT::[dbo].[get_orders] TO [report_user];", script.Last());
    }

    [Fact]
    public void Script_text_separates_batches_with_go()
    {
        var request = Request();
        request.ProcedureBody = "SELECT 1";

        string text = AccessScript.ToText(request);

        Assert.Contains("END;" + Environment.NewLine + "GO", text);
        // One GO per statement: create procedure, create user, grant.
        var lines = text.Split(new[] { Environment.NewLine }, StringSplitOptions.None);
        Assert.Equal(3, lines.Count(l => l == "GO"));
    }

    [Fact]
    public void Names_are_quoted_so_they_cannot_break_out()
    {
        var request = Request();
        request.Name = "x]; DROP TABLE t; --";

        var script = AccessScript.Build(request);

        Assert.Contains("[x]]; DROP TABLE t; --]", script[0]);
        Assert.Contains("N'x]; DROP TABLE t; --'", script[0]);
    }

    [Fact]
    public void Sql_server_script_text_starts_in_the_database()
    {
        Assert.StartsWith("USE [sales];", AccessScript.ToText(Request(azure: false)));
        Assert.StartsWith("IF DATABASE_PRINCIPAL_ID", AccessScript.ToText(Request()));
    }

    [Fact]
    public void Requires_a_name_and_a_password_for_sql_logins()
    {
        var noName = Request(); noName.Name = " ";
        var noPassword = Request(); noPassword.Password = "";

        Assert.Throws<ArgumentException>(() => AccessScript.Build(noName));
        Assert.Throws<ArgumentException>(() => AccessScript.Build(noPassword));
    }

    [Fact]
    public void Generated_passwords_meet_the_complexity_policy()
    {
        for (int i = 0; i < 50; i++)
        {
            string password = AccessScript.GeneratePassword();

            Assert.Equal(20, password.Length);
            Assert.Contains(password, char.IsUpper);
            Assert.Contains(password, char.IsLower);
            Assert.Contains(password, char.IsDigit);
            Assert.Contains(password, c => !char.IsLetterOrDigit(c));
            Assert.DoesNotContain('\'', password);
        }
    }
}
