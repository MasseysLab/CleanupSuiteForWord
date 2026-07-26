using System.Security.AccessControl;
using System.Security.Principal;

namespace MasseysLab.CleanupSuite.Engine;

public static class JobAccessPolicy
{
    public static void ValidateOwnerOnly(string jobDirectory)
    {
        SecurityIdentifier currentUser = CurrentUserSid();
        DirectorySecurity security = new DirectoryInfo(jobDirectory)
            .GetAccessControl(AccessControlSections.Owner | AccessControlSections.Access);
        SecurityIdentifier? owner = security.GetOwner(
            typeof(SecurityIdentifier)) as SecurityIdentifier;
        if (owner is null || !owner.Equals(currentUser))
        {
            throw SecurityError(
                "The job directory is not owned by the current user.");
        }

        if (!security.AreAccessRulesProtected)
        {
            throw SecurityError(
                "The job directory inherits permissions and is not owner-only.");
        }

        bool currentUserHasFullControl = false;
        AuthorizationRuleCollection rules = security.GetAccessRules(
            includeExplicit: true,
            includeInherited: true,
            targetType: typeof(SecurityIdentifier));
        foreach (FileSystemAccessRule rule in rules)
        {
            if (rule.IsInherited)
            {
                throw SecurityError(
                    "The job directory contains inherited access rules.");
            }

            if (rule.IdentityReference is not SecurityIdentifier identity
                || !identity.Equals(currentUser))
            {
                throw SecurityError(
                    "The job directory grants access to another identity.");
            }

            if (rule.AccessControlType == AccessControlType.Allow
                && (rule.FileSystemRights & FileSystemRights.FullControl)
                    == FileSystemRights.FullControl)
            {
                currentUserHasFullControl = true;
            }
        }

        if (!currentUserHasFullControl)
        {
            throw SecurityError(
                "The current user does not have full control of the job directory.");
        }
    }

    public static void HardenForCurrentUser(string jobDirectory)
    {
        SecurityIdentifier currentUser = CurrentUserSid();
        DirectorySecurity security = new();
        security.SetOwner(currentUser);
        security.SetAccessRuleProtection(
            isProtected: true,
            preserveInheritance: false);
        security.AddAccessRule(
            new FileSystemAccessRule(
                currentUser,
                FileSystemRights.FullControl,
                InheritanceFlags.ContainerInherit
                    | InheritanceFlags.ObjectInherit,
                PropagationFlags.None,
                AccessControlType.Allow));
        new DirectoryInfo(jobDirectory).SetAccessControl(security);
    }

    private static SecurityIdentifier CurrentUserSid() =>
        WindowsIdentity.GetCurrent().User
        ?? throw SecurityError(
            "The current Windows user identity is unavailable.");

    private static EngineContractException SecurityError(string message) =>
        new(
            "security-error",
            "security-policy-violation",
            message,
            ContractConstants.ExitSecurityError);
}
