using Flank.PowerBi;

namespace Flank.Tests.Unit;

public class PowerBiHelpTests
{
    [Fact]
    public void Admin_approval_errors_explain_how_an_admin_approves_flank()
    {
        var error = new Exception("AADSTS90094: Admin consent is required.\nTrace ID: 123\nCorrelation ID: 456");

        string text = PowerBiHelp.Explain(PowerBiStep.SignIn, error);

        Assert.StartsWith("Flank couldn't sign you in to Power BI.", text);
        Assert.Contains("AADSTS90094: Admin consent is required.", text);
        Assert.DoesNotContain("Trace ID", text);
        Assert.Contains(PowerBiHelp.AsOf + ": your organization requires an admin to approve Flank.", text);
        Assert.Contains("Admin consent requests", text);
        Assert.Contains("Connection.Read.All", text);
    }

    [Fact]
    public void Another_tenant_is_told_this_version_cant_sign_in_there()
    {
        string text = PowerBiHelp.Explain(PowerBiStep.SignIn,
            new Exception("AADSTS50020: User account from identity provider does not exist in tenant."));

        Assert.Contains("can't sign in to your organization's Microsoft Entra tenant", text);
    }

    [Fact]
    public void Connect_failures_say_to_become_a_user_of_the_connection()
    {
        string text = PowerBiHelp.Explain(PowerBiStep.Connect,
            new PowerBiException("Power BI returned 401 Unauthorized: DMTS_UserNotFoundInADGraphError"));

        Assert.StartsWith("Flank uploaded the report but couldn't connect it to the database.", text);
        Assert.Contains("Manage users", text);
        Assert.Contains("Allow Azure services", text);
    }

    [Fact]
    public void Upload_failures_mention_the_workspace_role_and_licence()
    {
        string text = PowerBiHelp.Explain(PowerBiStep.Upload, new PowerBiException("Power BI returned 403 Forbidden."));

        Assert.StartsWith("Power BI didn't accept the report.", text);
        Assert.Contains("Contributor", text);
        Assert.Contains("Premium Per User", text);
    }

    [Fact]
    public void Unrecognised_workspace_errors_get_no_made_up_fix()
    {
        string text = PowerBiHelp.Explain(PowerBiStep.Workspaces, new PowerBiException("Power BI returned 500 Internal Server Error."));

        Assert.DoesNotContain(PowerBiHelp.AsOf, text);
    }
}
