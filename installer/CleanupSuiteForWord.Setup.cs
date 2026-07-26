using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using System.Web.Script.Serialization;
using Microsoft.Win32;

[assembly: AssemblyTitle("CleanupSuite For Word Setup")]
[assembly: AssemblyCompany("MasseysLab")]
[assembly: AssemblyProduct("CleanupSuite For Word")]
[assembly: AssemblyCopyright("Copyright (c) 2026 MasseysLab")]
[assembly: AssemblyVersion("0.9.5.0")]
[assembly: AssemblyFileVersion("0.9.5.0")]

namespace MasseysLab.CleanupSuiteForWord.Setup
{
    internal static class Program
    {
        internal const string ProductName = "CleanupSuite For Word";
        internal const string DisplayVersion = "0.9.5-beta";
        internal const string ProtocolVersion = "1.0";
        internal const string EngineVersion = "0.2.0";
        internal const int BackupLimit = 3;

        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Options options = Options.Parse(args);

            try
            {
                Paths paths = Paths.Create(options);
                if (options.InspectOnly)
                {
                    InstallationState inspected = Installer.Detect(paths);
                    options.WriteStateReport(inspected, 0);
                    return 0;
                }

                if (options.Silent || options.Action != MaintenanceAction.None)
                {
                    MaintenanceAction action = options.Action;
                    if (action == MaintenanceAction.None)
                        action = Installer.RecommendedAction(Installer.Detect(paths));
                    int code = Installer.Execute(action, options, null);
                    options.WriteStateReport(Installer.Detect(paths), code);
                    return code;
                }

                Application.Run(new SetupForm(options, paths));
                return SetupForm.ProcessExitCode;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                    MessageBox.Show("Setup could not finish.\r\n\r\n" + ex.Message,
                        ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }
    }

    internal enum MaintenanceAction
    {
        None,
        Install,
        Update,
        Repair,
        Uninstall
    }

    internal enum InstallationKind
    {
        Absent,
        Older,
        CurrentHealthy,
        CurrentDamaged,
        Newer
    }

    internal sealed class Options
    {
        internal bool Silent;
        internal bool InspectOnly;
        internal MaintenanceAction Action;
        internal string TestRoot;
        internal string StateReportPath;
        internal string TestFailurePoint;

        internal static Options Parse(string[] args)
        {
            Options result = new Options();
            foreach (string raw in args)
            {
                string arg = raw.Trim();
                if (arg.Equals("/silent", StringComparison.OrdinalIgnoreCase)) result.Silent = true;
                if (arg.Equals("/inspect", StringComparison.OrdinalIgnoreCase)) result.InspectOnly = true;
                if (arg.Equals("/install", StringComparison.OrdinalIgnoreCase)) result.Action = MaintenanceAction.Install;
                if (arg.Equals("/update", StringComparison.OrdinalIgnoreCase)) result.Action = MaintenanceAction.Update;
                if (arg.Equals("/repair", StringComparison.OrdinalIgnoreCase)) result.Action = MaintenanceAction.Repair;
                if (arg.Equals("/uninstall", StringComparison.OrdinalIgnoreCase)) result.Action = MaintenanceAction.Uninstall;
                if (arg.StartsWith("/test-root=", StringComparison.OrdinalIgnoreCase))
                    result.TestRoot = arg.Substring("/test-root=".Length).Trim('"');
                if (arg.StartsWith("/state-report=", StringComparison.OrdinalIgnoreCase))
                    result.StateReportPath = arg.Substring("/state-report=".Length).Trim('"');
                if (arg.StartsWith("/test-fail-after=", StringComparison.OrdinalIgnoreCase))
                    result.TestFailurePoint = arg.Substring("/test-fail-after=".Length).Trim('"');
            }
            return result;
        }

        internal void WriteStateReport(InstallationState state, int exitCode)
        {
            if (string.IsNullOrWhiteSpace(StateReportPath)) return;
            string parent = Path.GetDirectoryName(Path.GetFullPath(StateReportPath));
            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
            string[] lines =
            {
                "exit_code=" + exitCode,
                "state=" + state.Kind,
                "installed_version=" + (state.InstalledVersion ?? ""),
                "package_valid=" + state.PackageValid.ToString().ToLowerInvariant(),
                "detail=" + (state.Detail ?? "").Replace("\r", " ").Replace("\n", " ")
            };
            File.WriteAllLines(StateReportPath, lines);
        }
    }

    internal sealed class SetupForm : Form
    {
        internal static int ProcessExitCode = 1;
        private readonly Options options;
        private readonly Paths paths;
        private InstallationState state;
        private readonly Label explanation;
        private readonly Label status;
        private readonly ProgressBar progress;
        private readonly Button primaryButton;
        private readonly Button repairButton;
        private readonly Button uninstallButton;
        private readonly Button closeButton;

        internal SetupForm(Options setupOptions, Paths setupPaths)
        {
            options = setupOptions;
            paths = setupPaths;
            state = Installer.Detect(paths);

            Text = Program.ProductName + " Setup";
            ClientSize = new Size(590, 286);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            Label title = new Label();
            title.Text = "CleanupSuite For Word";
            title.Font = new Font("Segoe UI Semibold", 18F, FontStyle.Bold);
            title.Location = new Point(24, 18);
            title.Size = new Size(535, 38);
            Controls.Add(title);

            explanation = new Label();
            explanation.Location = new Point(27, 65);
            explanation.Size = new Size(536, 58);
            Controls.Add(explanation);

            status = new Label();
            status.Location = new Point(27, 132);
            status.Size = new Size(536, 38);
            Controls.Add(status);

            progress = new ProgressBar();
            progress.Location = new Point(28, 177);
            progress.Size = new Size(534, 8);
            progress.Style = ProgressBarStyle.Continuous;
            progress.Minimum = 0;
            progress.Maximum = 100;
            Controls.Add(progress);

            primaryButton = NewButton(28, 214, 120);
            repairButton = NewButton(158, 214, 120);
            uninstallButton = NewButton(288, 214, 120);
            closeButton = NewButton(442, 214, 120);
            closeButton.Text = "Cancel";
            closeButton.Click += delegate { Close(); };
            CancelButton = closeButton;

            ConfigureForState();
        }

        private Button NewButton(int x, int y, int width)
        {
            Button button = new Button();
            button.Location = new Point(x, y);
            button.Size = new Size(width, 34);
            Controls.Add(button);
            return button;
        }

        private void ConfigureForState()
        {
            primaryButton.Visible = false;
            repairButton.Visible = false;
            uninstallButton.Visible = false;
            primaryButton.Click -= PrimaryClicked;
            repairButton.Click -= RepairClicked;
            uninstallButton.Click -= UninstallClicked;

            switch (state.Kind)
            {
                case InstallationKind.Absent:
                    explanation.Text = "Installs the matched CleanupSuite template and self-contained analysis engine for this Windows account. No separate .NET installation is required.";
                    status.Text = "CleanupSuite is not installed. Ready to install " + Program.DisplayVersion + ".";
                    primaryButton.Text = "Install";
                    primaryButton.Visible = true;
                    primaryButton.Click += PrimaryClicked;
                    AcceptButton = primaryButton;
                    break;
                case InstallationKind.Older:
                    explanation.Text = "An older CleanupSuite installation was found. Update installs this matched package; Repair/Reinstall verifies and replaces its files. Existing user settings are preserved.";
                    status.Text = "Installed: " + state.InstalledVersion + "   Available: " + Program.DisplayVersion + ".";
                    primaryButton.Text = "Update";
                    primaryButton.Visible = true;
                    primaryButton.Click += PrimaryClicked;
                    repairButton.Text = "Repair/Reinstall";
                    repairButton.Visible = true;
                    repairButton.Click += RepairClicked;
                    uninstallButton.Text = "Uninstall";
                    uninstallButton.Visible = true;
                    uninstallButton.Click += UninstallClicked;
                    AcceptButton = primaryButton;
                    break;
                case InstallationKind.CurrentHealthy:
                    explanation.Text = "The current matched CleanupSuite package is installed. Repair/Reinstall verifies and safely replaces the template, engine, protocol files, and installation manifest.";
                    status.Text = Program.DisplayVersion + " is installed and its component hashes match.";
                    repairButton.Text = "Repair/Reinstall";
                    repairButton.Visible = true;
                    repairButton.Click += RepairClicked;
                    uninstallButton.Text = "Uninstall";
                    uninstallButton.Visible = true;
                    uninstallButton.Click += UninstallClicked;
                    AcceptButton = repairButton;
                    break;
                case InstallationKind.CurrentDamaged:
                    explanation.Text = "CleanupSuite is installed, but one or more matched components are missing or changed. Apply should remain unavailable until Repair/Reinstall completes.";
                    status.Text = state.Detail;
                    repairButton.Text = "Repair/Reinstall";
                    repairButton.Visible = true;
                    repairButton.Click += RepairClicked;
                    uninstallButton.Text = "Uninstall";
                    uninstallButton.Visible = true;
                    uninstallButton.Click += UninstallClicked;
                    AcceptButton = repairButton;
                    break;
                default:
                    explanation.Text = "A newer CleanupSuite version is installed. This older Setup will not replace it. Use the newer Setup to repair, or uninstall this installation.";
                    status.Text = "Installed: " + state.InstalledVersion + "   This Setup: " + Program.DisplayVersion + ".";
                    uninstallButton.Text = "Uninstall";
                    uninstallButton.Visible = true;
                    uninstallButton.Click += UninstallClicked;
                    AcceptButton = closeButton;
                    break;
            }
        }

        private void PrimaryClicked(object sender, EventArgs e)
        {
            MaintenanceAction action = state.Kind == InstallationKind.Absent
                ? MaintenanceAction.Install : MaintenanceAction.Update;
            RunAction(action);
        }

        private void RepairClicked(object sender, EventArgs e)
        {
            RunAction(MaintenanceAction.Repair);
        }

        private void UninstallClicked(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                    "Remove CleanupSuite For Word from this Windows account?\r\n\r\n" +
                    "The three newest installer backups and your CleanupSuite user settings will be kept.",
                    Program.ProductName, MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return;
            RunAction(MaintenanceAction.Uninstall);
        }

        private void RunAction(MaintenanceAction action)
        {
            SetButtonsEnabled(false);
            progress.Value = 5;
            status.Text = "Preparing " + action.ToString().ToLowerInvariant() + "...";
            Application.DoEvents();

            Thread worker = new Thread(delegate()
            {
                int code = Installer.Execute(action, options, delegate(string text, int percent)
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        status.Text = text;
                        progress.Value = Math.Max(0, Math.Min(100, percent));
                    });
                });

                BeginInvoke((MethodInvoker)delegate
                {
                    ProcessExitCode = code;
                    state = Installer.Detect(paths);
                    if (code == 0)
                    {
                        progress.Value = 100;
                        status.Text = action == MaintenanceAction.Uninstall
                            ? "CleanupSuite was removed. Backups and user settings were kept."
                            : "CleanupSuite " + Program.DisplayVersion + " is ready. Reopen Word to use it.";
                        primaryButton.Visible = false;
                        repairButton.Visible = false;
                        uninstallButton.Visible = false;
                        closeButton.Text = "Close";
                        AcceptButton = closeButton;
                    }
                    else
                    {
                        progress.Value = 0;
                        ConfigureForState();
                        SetButtonsEnabled(true);
                    }
                });
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void SetButtonsEnabled(bool enabled)
        {
            primaryButton.Enabled = enabled;
            repairButton.Enabled = enabled;
            uninstallButton.Enabled = enabled;
        }
    }

    internal static class Installer
    {
        internal const string TemplateName = "CleanupSuite.dotm";
        internal const string ProductFolderName = "CleanupSuiteForWord";
        internal const string ManifestName = "installation-manifest.json";
        private const string UninstallKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\CleanupSuiteForWord";

        internal static MaintenanceAction RecommendedAction(InstallationState state)
        {
            if (state.Kind == InstallationKind.Absent) return MaintenanceAction.Install;
            if (state.Kind == InstallationKind.Older) return MaintenanceAction.Update;
            if (state.Kind == InstallationKind.CurrentHealthy || state.Kind == InstallationKind.CurrentDamaged)
                return MaintenanceAction.Repair;
            return MaintenanceAction.None;
        }

        internal static int Execute(MaintenanceAction action, Options options, Action<string, int> report)
        {
            if (action == MaintenanceAction.Uninstall)
                return Uninstall(options, report);
            return InstallPackage(action, options, report);
        }

        internal static InstallationState Detect(Paths paths)
        {
            bool templatePresent = File.Exists(paths.TemplatePath);
            bool manifestPresent = File.Exists(paths.ManifestPath);
            bool enginePresent = File.Exists(paths.EnginePath);
            if (!templatePresent && !manifestPresent && !enginePresent)
                return new InstallationState(InstallationKind.Absent, "", true, "CleanupSuite is not installed.");

            if (!manifestPresent)
                return new InstallationState(InstallationKind.Older, "legacy or unknown", false,
                    "A pre-hybrid CleanupSuite installation was found.");

            try
            {
                byte[] manifestBytes = File.ReadAllBytes(paths.ManifestPath);
                PackageManifest manifest = PackageManifest.Parse(manifestBytes);
                string problem = ValidateInstalledPackage(paths, manifest);
                int comparison = CompareVersions(manifest.SuiteVersion, Program.DisplayVersion);
                if (comparison < 0)
                    return new InstallationState(InstallationKind.Older, manifest.SuiteVersion,
                        string.IsNullOrEmpty(problem), string.IsNullOrEmpty(problem)
                            ? "An older matched installation was found."
                            : "The older installation also needs repair: " + problem);
                if (comparison > 0)
                    return new InstallationState(InstallationKind.Newer, manifest.SuiteVersion,
                        string.IsNullOrEmpty(problem), string.IsNullOrEmpty(problem)
                            ? "A newer installation was found."
                            : "The newer installation has a component problem: " + problem);
                if (string.IsNullOrEmpty(problem))
                    return new InstallationState(InstallationKind.CurrentHealthy, manifest.SuiteVersion, true,
                        "The current matched package is healthy.");
                return new InstallationState(InstallationKind.CurrentDamaged, manifest.SuiteVersion, false,
                    "Repair is recommended: " + problem);
            }
            catch (Exception ex)
            {
                return new InstallationState(InstallationKind.CurrentDamaged, "unknown", false,
                    "Repair is recommended because the installation manifest is invalid: " + ex.Message);
            }
        }

        private static int InstallPackage(MaintenanceAction action, Options options, Action<string, int> report)
        {
            try
            {
                Paths paths = Paths.Create(options);
                InstallationState before = Detect(paths);
                ValidateRequestedAction(action, before);
                EnsureWordClosed(paths);

                Directory.CreateDirectory(paths.StartupFolder);
                Directory.CreateDirectory(paths.ProductFolder);
                Directory.CreateDirectory(paths.BackupFolder);

                Report(report, "Reading the matched package...", 12);
                EmbeddedPackage package = EmbeddedPackage.Load();
                if (!paths.TestMode) RequireSignedSetup();

                string staging = Path.Combine(paths.ProductFolder, ".staging-" + Guid.NewGuid().ToString("N"));
                Directory.CreateDirectory(staging);
                List<RollbackEntry> rollback = new List<RollbackEntry>();
                try
                {
                    Report(report, "Verifying the template, engine, and protocol files...", 30);
                    package.StageAndVerify(staging);
                    if (!paths.TestMode)
                    {
                        PackageComponent engineComponent =
                            package.Manifest.RequireComponent("analysis-engine");
                        string stagedEngine = EmbeddedPackage.SafeCombine(
                            staging, engineComponent.RelativePath);
                        if (!AuthenticodeTrust.IsTrusted(stagedEngine))
                            throw new InvalidDataException(
                                "The analysis engine does not have a valid Authenticode signature.");
                    }

                    Report(report, "Backing up the previous template...", 48);
                    BackupExisting(paths);
                    RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);

                    foreach (PackageComponent component in package.Manifest.Components)
                    {
                        string target = paths.ResolveComponentPath(component);
                        rollback.Add(RollbackEntry.Capture(target));
                    }
                    rollback.Add(RollbackEntry.Capture(paths.ManifestPath));
                    rollback.Add(RollbackEntry.Capture(paths.InstallerCopyPath));

                    Report(report, "Installing the self-contained analysis engine...", 64);
                    InstallComponent(package, "analysis-engine", paths, staging);
                    if (paths.TestMode && options.TestFailurePoint == "engine")
                        throw new IOException("Intentional installer matrix failure after engine replacement.");
                    InstallComponent(package, "tool-definitions", paths, staging);
                    InstallComponent(package, "rules", paths, staging);

                    Report(report, "Installing the Word template...", 78);
                    InstallComponent(package, "word-template", paths, staging);

                    Report(report, "Binding the matched installation...", 88);
                    AtomicWrite(paths.ManifestPath, package.ManifestBytes);
                    CopyInstaller(paths);

                    string installedProblem = ValidateInstalledPackage(paths, package.Manifest);
                    if (!string.IsNullOrEmpty(installedProblem))
                        throw new InvalidDataException("Post-install verification failed: " + installedProblem);

                    RegisterUninstall(paths, package.Manifest.SuiteVersion);
                    Report(report, "Installation complete.", 100);
                    return 0;
                }
                catch
                {
                    RestoreRollback(rollback);
                    throw;
                }
                finally
                {
                    DeleteDirectoryBestEffort(staging);
                }
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                    MessageBox.Show("Setup did not make changes.\r\n\r\n" + ex.Message,
                        Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }
        }

        private static int Uninstall(Options options, Action<string, int> report)
        {
            try
            {
                Paths paths = Paths.Create(options);
                EnsureWordClosed(paths);
                Directory.CreateDirectory(paths.BackupFolder);
                Report(report, "Backing up the installed template...", 24);
                BackupExisting(paths);
                RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);

                Report(report, "Removing CleanupSuite components...", 62);
                DeleteFileBestEffort(paths.TemplatePath);
                DeleteFileBestEffort(paths.EnginePath);
                DeleteFileBestEffort(paths.ProtocolPath);
                DeleteFileBestEffort(paths.OperationsPath);
                DeleteFileBestEffort(paths.ManifestPath);
                DeleteFileBestEffort(paths.InstallerCopyPath);
                DeleteEmptyDirectoryBestEffort(paths.EngineFolder);
                DeleteEmptyDirectoryBestEffort(paths.ContractFolder);
                if (!paths.TestMode) Registry.CurrentUser.DeleteSubKeyTree(UninstallKeyPath, false);
                Report(report, "CleanupSuite was removed; backups and user settings were kept.", 100);
                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                    MessageBox.Show("CleanupSuite could not be removed.\r\n\r\n" + ex.Message,
                        Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 1;
            }
        }

        private static void ValidateRequestedAction(MaintenanceAction action, InstallationState state)
        {
            if (action == MaintenanceAction.Install && state.Kind != InstallationKind.Absent)
                throw new InvalidOperationException("CleanupSuite is already installed. Choose Update, Repair/Reinstall, or Uninstall.");
            if (action == MaintenanceAction.Update && state.Kind != InstallationKind.Older)
                throw new InvalidOperationException("Update is available only when an older CleanupSuite version is installed.");
            if (action == MaintenanceAction.Repair &&
                state.Kind != InstallationKind.Older &&
                state.Kind != InstallationKind.CurrentHealthy &&
                state.Kind != InstallationKind.CurrentDamaged)
                throw new InvalidOperationException("Repair/Reinstall is not available for the detected installation state.");
            if (state.Kind == InstallationKind.Newer)
                throw new InvalidOperationException("This Setup will not replace a newer CleanupSuite installation.");
        }

        private static void EnsureWordClosed(Paths paths)
        {
            if (!paths.TestMode && Process.GetProcessesByName("WINWORD").Length > 0)
                throw new InvalidOperationException("Microsoft Word is still open. Close every Word window, then run Setup again.");
        }

        private static void RequireSignedSetup()
        {
            if (!AuthenticodeTrust.IsTrusted(Application.ExecutablePath))
                throw new InvalidDataException(
                    "This official hybrid Setup must be Authenticode-signed before it can install executable components. " +
                    "Download a signed CleanupSuite Setup or use Repair from a signed copy.");
        }

        private static void InstallComponent(
            EmbeddedPackage package,
            string id,
            Paths paths,
            string staging)
        {
            PackageComponent component = package.Manifest.RequireComponent(id);
            string stagedPath = Path.Combine(staging, component.RelativePath);
            string targetPath = paths.ResolveComponentPath(component);
            byte[] bytes = File.ReadAllBytes(stagedPath);
            AtomicWrite(targetPath, bytes);
        }

        private static string ValidateInstalledPackage(Paths paths, PackageManifest manifest)
        {
            if (!manifest.Publisher.Equals("MasseysLab", StringComparison.Ordinal))
                return "the publisher identity is not recognized";
            if (!manifest.ProtocolVersion.Equals(Program.ProtocolVersion, StringComparison.Ordinal))
                return "the protocol version does not match";
            if (!manifest.AuthenticodeRequired)
                return "the official package does not require Authenticode";

            foreach (PackageComponent component in manifest.Components)
            {
                if (component.Id == "word-template" &&
                    !component.Version.Equals(manifest.SuiteVersion, StringComparison.Ordinal))
                    return "the Word template version does not match the package";
                if (component.Id == "analysis-engine" &&
                    !component.Version.Equals(Program.EngineVersion, StringComparison.Ordinal))
                    return "the analysis engine version does not match";
                if ((component.Id == "tool-definitions" || component.Id == "rules") &&
                    !component.Version.Equals("1.0.0", StringComparison.Ordinal))
                    return component.Id + " has an unsupported version";
                string target;
                try { target = paths.ResolveComponentPath(component); }
                catch (Exception ex) { return ex.Message; }
                if (!File.Exists(target)) return component.Id + " is missing";
                FileInfo info = new FileInfo(target);
                if (info.Length != component.ByteLength) return component.Id + " has the wrong length";
                if (!Sha256File(target).Equals(component.Sha256, StringComparison.OrdinalIgnoreCase))
                    return component.Id + " has a hash mismatch";
            }
            return "";
        }

        private static void CopyInstaller(Paths paths)
        {
            string source = Path.GetFullPath(Application.ExecutablePath);
            string target = Path.GetFullPath(paths.InstallerCopyPath);
            if (source.Equals(target, StringComparison.OrdinalIgnoreCase)) return;
            Directory.CreateDirectory(Path.GetDirectoryName(target));
            File.Copy(source, target, true);
        }

        private static void BackupExisting(Paths paths)
        {
            if (!File.Exists(paths.TemplatePath)) return;
            string backupName = "CleanupSuite-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".dotm";
            File.Copy(paths.TemplatePath, Path.Combine(paths.BackupFolder, backupName), false);
        }

        private static void RetainNewestBackups(string backupFolder, int limit)
        {
            DirectoryInfo folder = new DirectoryInfo(backupFolder);
            FileInfo[] backups = folder.GetFiles("CleanupSuite-*.dotm");
            Array.Sort(backups, delegate(FileInfo left, FileInfo right)
            {
                return right.LastWriteTimeUtc.CompareTo(left.LastWriteTimeUtc);
            });
            for (int index = limit; index < backups.Length; index++) backups[index].Delete();
        }

        private static void AtomicWrite(string targetPath, byte[] bytes)
        {
            string parent = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
            string temporary = targetPath + ".new-" + Guid.NewGuid().ToString("N");
            File.WriteAllBytes(temporary, bytes);
            try
            {
                if (File.Exists(targetPath)) File.Replace(temporary, targetPath, null);
                else File.Move(temporary, targetPath);
            }
            finally
            {
                DeleteFileBestEffort(temporary);
            }
        }

        private static void RestoreRollback(List<RollbackEntry> entries)
        {
            for (int index = entries.Count - 1; index >= 0; index--)
            {
                try { entries[index].Restore(); }
                catch { }
            }
        }

        private static void RegisterUninstall(Paths paths, string version)
        {
            if (paths.TestMode) return;
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKeyPath))
            {
                key.SetValue("DisplayName", Program.ProductName);
                key.SetValue("DisplayVersion", version);
                key.SetValue("Publisher", "MasseysLab");
                key.SetValue("URLInfoAbout", "https://github.com/MasseysLab/CleanupSuiteForWord");
                key.SetValue("UninstallString", "\"" + paths.InstallerCopyPath + "\" /uninstall");
                key.SetValue("QuietUninstallString", "\"" + paths.InstallerCopyPath + "\" /uninstall /silent");
                key.SetValue("ModifyPath", "\"" + paths.InstallerCopyPath + "\"");
                key.SetValue("InstallLocation", paths.ProductFolder);
                key.SetValue("NoModify", 0, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 0, RegistryValueKind.DWord);
            }
        }

        internal static string Sha256(byte[] bytes)
        {
            using (SHA256 hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
        }

        internal static string Sha256File(string path)
        {
            using (FileStream stream = File.OpenRead(path))
            using (SHA256 hash = SHA256.Create())
                return BitConverter.ToString(hash.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
        }

        private static int CompareVersions(string left, string right)
        {
            Version leftVersion = ParseNumericVersion(left);
            Version rightVersion = ParseNumericVersion(right);
            int result = leftVersion.CompareTo(rightVersion);
            if (result != 0) return result;
            return left.Equals(right, StringComparison.OrdinalIgnoreCase) ? 0 : -1;
        }

        private static Version ParseNumericVersion(string value)
        {
            string numeric = (value ?? "").Split('-')[0];
            Version parsed;
            if (!Version.TryParse(numeric, out parsed)) return new Version(0, 0);
            if (parsed.Build < 0) return new Version(parsed.Major, parsed.Minor, 0);
            return parsed;
        }

        private static void Report(Action<string, int> report, string message, int percent)
        {
            if (report != null) report(message, percent);
        }

        private static void DeleteFileBestEffort(string path)
        {
            try { if (File.Exists(path)) File.Delete(path); }
            catch { }
        }

        private static void DeleteDirectoryBestEffort(string path)
        {
            try { if (Directory.Exists(path)) Directory.Delete(path, true); }
            catch { }
        }

        private static void DeleteEmptyDirectoryBestEffort(string path)
        {
            try
            {
                if (Directory.Exists(path) && Directory.GetFileSystemEntries(path).Length == 0)
                    Directory.Delete(path, false);
            }
            catch { }
        }
    }

    internal sealed class InstallationState
    {
        internal readonly InstallationKind Kind;
        internal readonly string InstalledVersion;
        internal readonly bool PackageValid;
        internal readonly string Detail;

        internal InstallationState(
            InstallationKind kind,
            string installedVersion,
            bool packageValid,
            string detail)
        {
            Kind = kind;
            InstalledVersion = installedVersion;
            PackageValid = packageValid;
            Detail = detail;
        }
    }

    internal sealed class EmbeddedPackage
    {
        internal readonly PackageManifest Manifest;
        internal readonly byte[] ManifestBytes;
        private readonly Dictionary<string, byte[]> resources;

        private EmbeddedPackage(
            PackageManifest manifest,
            byte[] manifestBytes,
            Dictionary<string, byte[]> packageResources)
        {
            Manifest = manifest;
            ManifestBytes = manifestBytes;
            resources = packageResources;
        }

        internal static EmbeddedPackage Load()
        {
            byte[] manifestBytes = ReadResource("CleanupSuite.Payload.Manifest.json");
            PackageManifest manifest = PackageManifest.Parse(manifestBytes);
            if (!manifest.SuiteVersion.Equals(Program.DisplayVersion, StringComparison.Ordinal))
                throw new InvalidDataException("The embedded package version does not match this Setup.");
            if (!manifest.ProtocolVersion.Equals(Program.ProtocolVersion, StringComparison.Ordinal))
                throw new InvalidDataException("The embedded protocol version does not match this Setup.");

            Dictionary<string, byte[]> resources = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            resources.Add("word-template", ReadResource("CleanupSuite.Payload.Template.dotm"));
            resources.Add("analysis-engine", ReadResource("CleanupSuite.Payload.Engine.exe"));
            resources.Add("tool-definitions", ReadResource("CleanupSuite.Payload.Protocol.json"));
            resources.Add("rules", ReadResource("CleanupSuite.Payload.Operations.json"));
            foreach (PackageComponent component in manifest.Components)
            {
                byte[] bytes;
                if (!resources.TryGetValue(component.Id, out bytes))
                    throw new InvalidDataException("The embedded resource for " + component.Id + " is missing.");
                if (bytes.LongLength != component.ByteLength ||
                    !Installer.Sha256(bytes).Equals(component.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The embedded " + component.Id + " resource failed hash verification.");
            }
            return new EmbeddedPackage(manifest, manifestBytes, resources);
        }

        internal void StageAndVerify(string staging)
        {
            foreach (PackageComponent component in Manifest.Components)
            {
                string destination = SafeCombine(staging, component.RelativePath);
                Directory.CreateDirectory(Path.GetDirectoryName(destination));
                File.WriteAllBytes(destination, resources[component.Id]);
                if (new FileInfo(destination).Length != component.ByteLength ||
                    !Installer.Sha256File(destination).Equals(component.Sha256, StringComparison.OrdinalIgnoreCase))
                    throw new InvalidDataException("The staged " + component.Id + " resource failed verification.");
            }
        }

        private static byte[] ReadResource(string name)
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(name))
            {
                if (stream == null) throw new FileNotFoundException("Embedded resource not found: " + name);
                using (MemoryStream output = new MemoryStream())
                {
                    stream.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        internal static string SafeCombine(string root, string relativePath)
        {
            if (string.IsNullOrWhiteSpace(relativePath) ||
                Path.IsPathRooted(relativePath) ||
                relativePath.Contains(".."))
                throw new InvalidDataException("A package component path is unsafe.");
            string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar) + Path.DirectorySeparatorChar;
            string combined = Path.GetFullPath(Path.Combine(fullRoot, relativePath));
            if (!combined.StartsWith(fullRoot, StringComparison.OrdinalIgnoreCase))
                throw new InvalidDataException("A package component escaped its installation root.");
            return combined;
        }
    }

    internal sealed class PackageManifest
    {
        internal string SuiteVersion;
        internal string ProtocolVersion;
        internal string Publisher;
        internal bool AuthenticodeRequired;
        internal readonly List<PackageComponent> Components = new List<PackageComponent>();

        internal static PackageManifest Parse(byte[] bytes)
        {
            string json = new System.Text.UTF8Encoding(false, true).GetString(bytes);
            JavaScriptSerializer serializer = new JavaScriptSerializer();
            object parsed = serializer.DeserializeObject(json);
            Dictionary<string, object> root = parsed as Dictionary<string, object>;
            if (root == null) throw new InvalidDataException("The installation manifest root is invalid.");
            RequireExactKeys(root, new[]
            {
                "manifestVersion", "suiteVersion", "protocolVersion",
                "publisher", "authenticodeRequired", "components"
            });
            if (RequireString(root, "manifestVersion") != "1.0")
                throw new InvalidDataException("The installation manifest version is unsupported.");

            PackageManifest manifest = new PackageManifest();
            manifest.SuiteVersion = RequireString(root, "suiteVersion");
            manifest.ProtocolVersion = RequireString(root, "protocolVersion");
            manifest.Publisher = RequireString(root, "publisher");
            manifest.AuthenticodeRequired = RequireBoolean(root, "authenticodeRequired");

            object[] components = root["components"] as object[];
            if (components == null || components.Length < 4)
                throw new InvalidDataException("The installation manifest component list is incomplete.");
            foreach (object raw in components)
            {
                Dictionary<string, object> item = raw as Dictionary<string, object>;
                if (item == null) throw new InvalidDataException("A component entry is invalid.");
                RequireExactKeys(item, new[] { "id", "version", "relativePath", "sha256", "byteLength" });
                PackageComponent component = new PackageComponent();
                component.Id = RequireString(item, "id");
                component.Version = RequireString(item, "version");
                component.RelativePath = RequireString(item, "relativePath");
                component.Sha256 = RequireString(item, "sha256").ToLowerInvariant();
                component.ByteLength = Convert.ToInt64(item["byteLength"]);
                if (component.ByteLength < 1 || component.Sha256.Length != 64)
                    throw new InvalidDataException("A component identity is invalid.");
                if (manifest.Components.Exists(delegate(PackageComponent existing)
                    { return existing.Id.Equals(component.Id, StringComparison.Ordinal); }))
                    throw new InvalidDataException("A component id is duplicated.");
                manifest.Components.Add(component);
            }
            manifest.RequireComponent("word-template");
            manifest.RequireComponent("analysis-engine");
            manifest.RequireComponent("tool-definitions");
            manifest.RequireComponent("rules");
            if (!manifest.Publisher.Equals("MasseysLab", StringComparison.Ordinal))
                throw new InvalidDataException("The installation manifest publisher is not recognized.");
            if (!manifest.AuthenticodeRequired)
                throw new InvalidDataException("The installation manifest does not require Authenticode.");
            return manifest;
        }

        internal PackageComponent RequireComponent(string id)
        {
            PackageComponent found = Components.Find(delegate(PackageComponent component)
                { return component.Id.Equals(id, StringComparison.Ordinal); });
            if (found == null) throw new InvalidDataException("Required component missing: " + id);
            return found;
        }

        private static void RequireExactKeys(Dictionary<string, object> values, string[] expected)
        {
            if (values.Count != expected.Length)
                throw new InvalidDataException("The installation manifest contains unexpected fields.");
            foreach (string key in expected)
                if (!values.ContainsKey(key))
                    throw new InvalidDataException("The installation manifest is missing " + key + ".");
        }

        private static string RequireString(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || !(value is string) || string.IsNullOrWhiteSpace((string)value))
                throw new InvalidDataException("The installation manifest field " + key + " is invalid.");
            return (string)value;
        }

        private static bool RequireBoolean(Dictionary<string, object> values, string key)
        {
            object value;
            if (!values.TryGetValue(key, out value) || !(value is bool))
                throw new InvalidDataException("The installation manifest field " + key + " is invalid.");
            return (bool)value;
        }
    }

    internal sealed class PackageComponent
    {
        internal string Id;
        internal string Version;
        internal string RelativePath;
        internal string Sha256;
        internal long ByteLength;
    }

    internal sealed class Paths
    {
        internal string StartupFolder;
        internal string ProductFolder;
        internal string BackupFolder;
        internal string EngineFolder;
        internal string ContractFolder;
        internal string TemplatePath;
        internal string EnginePath;
        internal string ProtocolPath;
        internal string OperationsPath;
        internal string ManifestPath;
        internal string InstallerCopyPath;
        internal bool TestMode;

        internal static Paths Create(Options options)
        {
            string appData;
            string localAppData;
            if (!string.IsNullOrEmpty(options.TestRoot))
            {
                string root = Path.GetFullPath(options.TestRoot);
                appData = Path.Combine(root, "Roaming");
                localAppData = Path.Combine(root, "Local");
            }
            else
            {
                appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }

            Paths result = new Paths();
            result.TestMode = !string.IsNullOrEmpty(options.TestRoot);
            result.StartupFolder = Path.Combine(appData, "Microsoft", "Word", "STARTUP");
            result.ProductFolder = Path.Combine(localAppData, "MasseysLab", Installer.ProductFolderName);
            result.BackupFolder = Path.Combine(result.ProductFolder, "Backups");
            result.EngineFolder = Path.Combine(result.ProductFolder, "Engine");
            result.ContractFolder = Path.Combine(result.ProductFolder, "Contracts", "Hybrid", "v1");
            result.TemplatePath = Path.Combine(result.StartupFolder, Installer.TemplateName);
            result.EnginePath = Path.Combine(result.EngineFolder, "CleanupSuite.Engine.exe");
            result.ProtocolPath = Path.Combine(result.ContractFolder, "protocol.json");
            result.OperationsPath = Path.Combine(result.ContractFolder, "operation-vocabulary.json");
            result.ManifestPath = Path.Combine(result.ProductFolder, Installer.ManifestName);
            result.InstallerCopyPath = Path.Combine(result.ProductFolder, "CleanupSuiteForWord-Setup.exe");
            return result;
        }

        internal string ResolveComponentPath(PackageComponent component)
        {
            string expected;
            if (component.Id == "word-template")
            {
                if (component.RelativePath != "CleanupSuite.dotm")
                    throw new InvalidDataException("The template path in the manifest is invalid.");
                return TemplatePath;
            }
            if (component.Id == "analysis-engine") expected = @"Engine\CleanupSuite.Engine.exe";
            else if (component.Id == "tool-definitions") expected = @"Contracts\Hybrid\v1\protocol.json";
            else if (component.Id == "rules") expected = @"Contracts\Hybrid\v1\operation-vocabulary.json";
            else throw new InvalidDataException("The manifest contains an unknown component.");
            if (!component.RelativePath.Equals(expected, StringComparison.Ordinal))
                throw new InvalidDataException("The " + component.Id + " path in the manifest is invalid.");
            return EmbeddedPackage.SafeCombine(ProductFolder, component.RelativePath);
        }
    }

    internal sealed class RollbackEntry
    {
        private readonly string path;
        private readonly bool existed;
        private readonly byte[] bytes;

        private RollbackEntry(string target, bool wasPresent, byte[] originalBytes)
        {
            path = target;
            existed = wasPresent;
            bytes = originalBytes;
        }

        internal static RollbackEntry Capture(string path)
        {
            return File.Exists(path)
                ? new RollbackEntry(path, true, File.ReadAllBytes(path))
                : new RollbackEntry(path, false, null);
        }

        internal void Restore()
        {
            if (existed) InstallerAtomicWrite(path, bytes);
            else if (File.Exists(path)) File.Delete(path);
        }

        private static void InstallerAtomicWrite(string targetPath, byte[] original)
        {
            string parent = Path.GetDirectoryName(targetPath);
            if (!string.IsNullOrEmpty(parent)) Directory.CreateDirectory(parent);
            string temporary = targetPath + ".rollback-" + Guid.NewGuid().ToString("N");
            File.WriteAllBytes(temporary, original);
            if (File.Exists(targetPath)) File.Replace(temporary, targetPath, null);
            else File.Move(temporary, targetPath);
        }
    }

    internal static class AuthenticodeTrust
    {
        private static readonly Guid GenericVerifyV2 =
            new Guid("00AAC56B-CD44-11d0-8CC2-00C04FC295EE");
        private const uint WtdUiNone = 2;
        private const uint WtdRevokeNone = 0;
        private const uint WtdChoiceFile = 1;
        private const uint WtdStateActionIgnore = 0;
        private const uint WtdCacheOnlyUrlRetrieval = 0x00001000;

        [DllImport("wintrust.dll", ExactSpelling = true, SetLastError = true)]
        private static extern uint WinVerifyTrust(
            IntPtr windowHandle,
            [MarshalAs(UnmanagedType.LPStruct)] Guid actionId,
            IntPtr trustData);

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WinTrustFileInfo
        {
            internal uint StructureSize;
            [MarshalAs(UnmanagedType.LPWStr)]
            internal string FilePath;
            internal IntPtr FileHandle;
            internal IntPtr KnownSubject;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        private struct WinTrustData
        {
            internal uint StructureSize;
            internal IntPtr PolicyCallbackData;
            internal IntPtr SipClientData;
            internal uint UiChoice;
            internal uint RevocationChecks;
            internal uint UnionChoice;
            internal IntPtr FileInfo;
            internal uint StateAction;
            internal IntPtr StateData;
            internal IntPtr UrlReference;
            internal uint ProviderFlags;
            internal uint UiContext;
        }

        internal static bool IsTrusted(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path)) return false;
            IntPtr fileInfoPointer = IntPtr.Zero;
            IntPtr trustDataPointer = IntPtr.Zero;
            try
            {
                WinTrustFileInfo fileInfo = new WinTrustFileInfo();
                fileInfo.StructureSize = (uint)Marshal.SizeOf(typeof(WinTrustFileInfo));
                fileInfo.FilePath = Path.GetFullPath(path);
                fileInfo.FileHandle = IntPtr.Zero;
                fileInfo.KnownSubject = IntPtr.Zero;
                fileInfoPointer = Marshal.AllocCoTaskMem(
                    Marshal.SizeOf(typeof(WinTrustFileInfo)));
                Marshal.StructureToPtr(fileInfo, fileInfoPointer, false);

                WinTrustData trustData = new WinTrustData();
                trustData.StructureSize = (uint)Marshal.SizeOf(typeof(WinTrustData));
                trustData.PolicyCallbackData = IntPtr.Zero;
                trustData.SipClientData = IntPtr.Zero;
                trustData.UiChoice = WtdUiNone;
                trustData.RevocationChecks = WtdRevokeNone;
                trustData.UnionChoice = WtdChoiceFile;
                trustData.FileInfo = fileInfoPointer;
                trustData.StateAction = WtdStateActionIgnore;
                trustData.StateData = IntPtr.Zero;
                trustData.UrlReference = IntPtr.Zero;
                trustData.ProviderFlags = WtdCacheOnlyUrlRetrieval;
                trustData.UiContext = 0;
                trustDataPointer = Marshal.AllocCoTaskMem(
                    Marshal.SizeOf(typeof(WinTrustData)));
                Marshal.StructureToPtr(trustData, trustDataPointer, false);

                return WinVerifyTrust(
                    new IntPtr(-1), GenericVerifyV2, trustDataPointer) == 0;
            }
            catch
            {
                return false;
            }
            finally
            {
                if (trustDataPointer != IntPtr.Zero)
                {
                    Marshal.DestroyStructure(
                        trustDataPointer, typeof(WinTrustData));
                    Marshal.FreeCoTaskMem(trustDataPointer);
                }
                if (fileInfoPointer != IntPtr.Zero)
                {
                    Marshal.DestroyStructure(
                        fileInfoPointer, typeof(WinTrustFileInfo));
                    Marshal.FreeCoTaskMem(fileInfoPointer);
                }
            }
        }
    }
}
