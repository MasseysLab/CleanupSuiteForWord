using System;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("CleanupSuite For Word VBA-Only Setup")]
[assembly: AssemblyCompany("MasseysLab")]
[assembly: AssemblyProduct("CleanupSuite For Word")]
[assembly: AssemblyCopyright("Copyright (c) 2026 MasseysLab")]
[assembly: AssemblyVersion("0.9.3.0")]
[assembly: AssemblyFileVersion("0.9.3.0")]

namespace MasseysLab.CleanupSuiteForWord.VbaOnlySetup
{
    internal static class Program
    {
        internal const string ProductName = "CleanupSuite For Word";
        internal const string DisplayVersion = "0.9.3-beta-vba-final";
        internal const string DisplayLabel = "0.9.3 Stable - Final VBA-Only Release";
        internal const string ExpectedTemplateSha256 =
            "4f815ee0df4bacbda5150a82db2d66739592fa2098184c8d7ec7b535d73ab10c";
        internal const int BackupLimit = 3;

        [STAThread]
        private static int Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            Options options = Options.Parse(args);
            try
            {
                if (options.Uninstall)
                    return Installer.Uninstall(options);

                if (options.Silent)
                    return Installer.Install(options, null);

                Application.Run(new SetupForm(options));
                return SetupForm.ProcessExitCode;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "Setup could not finish.\r\n\r\n" + ex.Message,
                        ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                return 1;
            }
        }
    }

    internal sealed class Options
    {
        internal bool Silent;
        internal bool Uninstall;
        internal string TestRoot;

        internal static Options Parse(string[] args)
        {
            Options result = new Options();
            foreach (string raw in args)
            {
                string arg = raw.Trim();
                if (arg.Equals("/silent", StringComparison.OrdinalIgnoreCase))
                    result.Silent = true;
                if (arg.Equals("/uninstall", StringComparison.OrdinalIgnoreCase))
                    result.Uninstall = true;
                if (arg.StartsWith("/test-root=", StringComparison.OrdinalIgnoreCase))
                    result.TestRoot = arg.Substring("/test-root=".Length).Trim('"');
            }
            return result;
        }
    }

    internal sealed class SetupForm : Form
    {
        internal static int ProcessExitCode = 1;

        private readonly Options options;
        private readonly Paths paths;
        private readonly Label explanation;
        private readonly Label status;
        private readonly ProgressBar progress;
        private readonly Button primaryButton;
        private readonly Button repairButton;
        private readonly Button uninstallButton;
        private readonly Button closeButton;

        internal SetupForm(Options setupOptions)
        {
            options = setupOptions;
            paths = Paths.Create(options);
            Text = Program.ProductName + " Setup";
            ClientSize = new Size(590, 330);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            Label title = new Label();
            title.Text = "CleanupSuite For Word";
            title.Font = new Font("Segoe UI Semibold", 18F, FontStyle.Bold);
            title.Location = new Point(24, 18);
            title.Size = new Size(520, 38);
            Controls.Add(title);

            Label release = new Label();
            release.Text = Program.DisplayLabel;
            release.Font = new Font("Segoe UI Semibold", 10F, FontStyle.Bold);
            release.ForeColor = Color.FromArgb(31, 78, 121);
            release.Location = new Point(27, 59);
            release.Size = new Size(535, 24);
            Controls.Add(release);

            explanation = new Label();
            explanation.Location = new Point(27, 91);
            explanation.Size = new Size(535, 48);
            Controls.Add(explanation);

            Label warning = new Label();
            warning.Text =
                "Unsigned release: continue only after verifying the official SHA-256 " +
                "and scanning Setup with Microsoft Defender.";
            warning.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
            warning.ForeColor = Color.FromArgb(176, 58, 46);
            warning.Location = new Point(27, 145);
            warning.Size = new Size(535, 42);
            Controls.Add(warning);

            status = new Label();
            status.Location = new Point(27, 194);
            status.Size = new Size(535, 38);
            Controls.Add(status);

            progress = new ProgressBar();
            progress.Location = new Point(28, 239);
            progress.Size = new Size(534, 8);
            progress.Style = ProgressBarStyle.Continuous;
            progress.Minimum = 0;
            progress.Maximum = 100;
            Controls.Add(progress);

            primaryButton = NewButton(28, 270, 120);
            repairButton = NewButton(158, 270, 120);
            uninstallButton = NewButton(288, 270, 120);
            closeButton = NewButton(442, 270, 120);
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
            bool installed = File.Exists(paths.TemplatePath);
            bool current = installed && FileSha256(paths.TemplatePath).Equals(
                Program.ExpectedTemplateSha256,
                StringComparison.OrdinalIgnoreCase);

            primaryButton.Visible = false;
            repairButton.Visible = false;
            uninstallButton.Visible = false;
            primaryButton.Click -= InstallOrUpdateClicked;
            repairButton.Click -= RepairClicked;
            uninstallButton.Click -= UninstallClicked;

            if (!installed)
            {
                explanation.Text =
                    "Installs the stable standalone VBA-only CleanupSuite template for this " +
                    "Windows account. Word must be closed. Existing installations are backed up.";
                status.Text = "CleanupSuite is not installed. Ready to install 0.9.3 Stable.";
                primaryButton.Text = "Install";
                primaryButton.Visible = true;
                primaryButton.Click += InstallOrUpdateClicked;
                AcceptButton = primaryButton;
            }
            else if (current)
            {
                explanation.Text =
                    "CleanupSuite 0.9.3 Stable is installed. Repair/Reinstall verifies and " +
                    "safely replaces the template; Uninstall removes it while keeping backups.";
                status.Text = "The installed template matches the published stable release.";
                repairButton.Text = "Repair/Reinstall";
                repairButton.Visible = true;
                repairButton.Click += RepairClicked;
                uninstallButton.Text = "Uninstall";
                uninstallButton.Visible = true;
                uninstallButton.Click += UninstallClicked;
                AcceptButton = repairButton;
            }
            else
            {
                explanation.Text =
                    "Another CleanupSuite template is installed. Update replaces it with " +
                    "0.9.3 Stable; Repair/Reinstall performs the same verified replacement.";
                status.Text = "An existing CleanupSuite installation was found.";
                primaryButton.Text = "Update";
                primaryButton.Visible = true;
                primaryButton.Click += InstallOrUpdateClicked;
                repairButton.Text = "Repair/Reinstall";
                repairButton.Visible = true;
                repairButton.Click += RepairClicked;
                uninstallButton.Text = "Uninstall";
                uninstallButton.Visible = true;
                uninstallButton.Click += UninstallClicked;
                AcceptButton = primaryButton;
            }
        }

        private static string FileSha256(string path)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(path))
                return BitConverter.ToString(sha.ComputeHash(stream)).Replace("-", "").ToLowerInvariant();
        }

        private void InstallOrUpdateClicked(object sender, EventArgs e)
        {
            RunInstall();
        }

        private void RepairClicked(object sender, EventArgs e)
        {
            RunInstall();
        }

        private void UninstallClicked(object sender, EventArgs e)
        {
            if (MessageBox.Show(
                    "Remove CleanupSuite For Word from this Windows account?\r\n\r\n" +
                    "The three newest installer backups will be kept.",
                    Program.ProductName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
                return;

            SetButtonsEnabled(false);
            progress.Value = 12;
            status.Text = "Removing CleanupSuite...";
            Application.DoEvents();

            Thread worker = new Thread(delegate()
            {
                int code = Installer.Uninstall(options, false);
                BeginInvoke((MethodInvoker)delegate
                {
                    ProcessExitCode = code;
                    progress.Value = code == 0 ? 100 : 0;
                    status.Text = code == 0
                        ? "CleanupSuite was removed. Installer backups were kept."
                        : "CleanupSuite could not be removed. Close Word and try again.";
                    if (code == 0)
                    {
                        primaryButton.Visible = false;
                        repairButton.Visible = false;
                        uninstallButton.Visible = false;
                        closeButton.Text = "Close";
                        AcceptButton = closeButton;
                    }
                    else
                    {
                        ConfigureForState();
                        SetButtonsEnabled(true);
                    }
                });
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }

        private void RunInstall()
        {
            SetButtonsEnabled(false);
            progress.Value = 12;
            status.Text = "Verifying the embedded 0.9.3 template...";
            Application.DoEvents();

            Thread worker = new Thread(delegate()
            {
                int code = Installer.Install(options, delegate(string text, int percent)
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
                    if (code == 0)
                    {
                        status.Text =
                            "CleanupSuite 0.9.3 Stable is ready. Reopen Word to use it.";
                        progress.Value = 100;
                        primaryButton.Visible = false;
                        repairButton.Visible = false;
                        uninstallButton.Visible = false;
                        closeButton.Text = "Close";
                        AcceptButton = closeButton;
                    }
                    else
                    {
                        status.Text =
                            "Setup did not finish. Close Word and try again.";
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
        private const string UninstallKeyPath =
            @"Software\Microsoft\Windows\CurrentVersion\Uninstall\CleanupSuiteForWord";

        internal static int Install(Options options, Action<string, int> report)
        {
            try
            {
                if (string.IsNullOrEmpty(options.TestRoot) &&
                    Process.GetProcessesByName("WINWORD").Length > 0)
                {
                    throw new InvalidOperationException(
                        "Microsoft Word is still open. Close every Word window, then run Setup again.");
                }

                Paths paths = Paths.Create(options);
                Directory.CreateDirectory(paths.StartupFolder);
                Directory.CreateDirectory(paths.ProductFolder);
                Directory.CreateDirectory(paths.BackupFolder);

                Report(report, "Reading the embedded 0.9.3 template...", 20);
                byte[] template = ReadEmbeddedTemplate();
                Report(report, "Verifying the published SHA-256...", 42);
                string templateHash = Sha256(template);
                if (!templateHash.Equals(
                        Program.ExpectedTemplateSha256,
                        StringComparison.OrdinalIgnoreCase))
                {
                    throw new InvalidDataException(
                        "The embedded template did not match the published 0.9.3 SHA-256.");
                }

                Report(report, "Backing up the previous installation...", 60);
                BackupExisting(paths);
                RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);

                Report(report, "Installing CleanupSuite 0.9.3 for this account...", 78);
                AtomicWrite(paths.TemplatePath, template);
                File.Copy(Application.ExecutablePath, paths.InstallerCopyPath, true);
                RegisterUninstall(paths);

                Report(report, "Installation complete.", 100);
                if (!options.Silent && report == null)
                {
                    MessageBox.Show(
                        Program.DisplayLabel + " was installed for this Windows account.\r\n\r\n" +
                        "Reopen Word and select CleanupSuite on the ribbon.",
                        Program.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "Setup did not make changes.\r\n\r\n" + ex.Message,
                        Program.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return 1;
            }
        }

        internal static int Uninstall(Options options)
        {
            return Uninstall(options, true);
        }

        internal static int Uninstall(Options options, bool confirm)
        {
            if (string.IsNullOrEmpty(options.TestRoot) &&
                Process.GetProcessesByName("WINWORD").Length > 0)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "Close Microsoft Word before uninstalling CleanupSuite.",
                        Program.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);
                }
                return 1;
            }

            if (confirm && !options.Silent &&
                MessageBox.Show(
                    "Remove CleanupSuite For Word from this Windows account?\r\n\r\n" +
                    "Your three newest installer backups will be kept.",
                    Program.ProductName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Question) != DialogResult.Yes)
            {
                return 2;
            }

            try
            {
                Paths paths = Paths.Create(options);
                Directory.CreateDirectory(paths.BackupFolder);
                BackupExisting(paths);
                RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);
                if (File.Exists(paths.TemplatePath))
                    File.Delete(paths.TemplatePath);
                if (!paths.TestMode)
                    Registry.CurrentUser.DeleteSubKeyTree(UninstallKeyPath, false);
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "CleanupSuite was removed. Installer backups were kept at:\r\n" +
                        paths.BackupFolder,
                        Program.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Information);
                }
                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                {
                    MessageBox.Show(
                        "CleanupSuite could not be removed.\r\n\r\n" + ex.Message,
                        Program.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                }
                return 1;
            }
        }

        private static byte[] ReadEmbeddedTemplate()
        {
            using (Stream stream =
                Assembly.GetExecutingAssembly().GetManifestResourceStream("CleanupSuite.dotm"))
            {
                if (stream == null)
                    throw new FileNotFoundException(
                        "The embedded CleanupSuite template was not found.");
                using (MemoryStream output = new MemoryStream())
                {
                    stream.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        private static void BackupExisting(Paths paths)
        {
            if (!File.Exists(paths.TemplatePath))
                return;

            string backupName =
                "CleanupSuite-" + DateTime.Now.ToString("yyyyMMdd-HHmmss-fff") + ".dotm";
            File.Copy(
                paths.TemplatePath,
                Path.Combine(paths.BackupFolder, backupName),
                false);
        }

        private static void RetainNewestBackups(string backupFolder, int limit)
        {
            DirectoryInfo folder = new DirectoryInfo(backupFolder);
            FileInfo[] backups = folder.GetFiles("CleanupSuite-*.dotm");
            Array.Sort(backups, delegate(FileInfo left, FileInfo right)
            {
                return right.LastWriteTimeUtc.CompareTo(left.LastWriteTimeUtc);
            });
            for (int index = limit; index < backups.Length; index++)
                backups[index].Delete();
        }

        private static void AtomicWrite(string targetPath, byte[] bytes)
        {
            string temporary = targetPath + ".new";
            File.WriteAllBytes(temporary, bytes);
            try
            {
                if (File.Exists(targetPath))
                    File.Replace(temporary, targetPath, null);
                else
                    File.Move(temporary, targetPath);
            }
            finally
            {
                if (File.Exists(temporary))
                    File.Delete(temporary);
            }
        }

        private static void RegisterUninstall(Paths paths)
        {
            if (paths.TestMode)
                return;

            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(UninstallKeyPath))
            {
                key.SetValue("DisplayName", Program.ProductName);
                key.SetValue("DisplayVersion", Program.DisplayVersion);
                key.SetValue("Publisher", "MasseysLab");
                key.SetValue(
                    "URLInfoAbout",
                    "https://github.com/MasseysLab/CleanupSuiteForWord");
                key.SetValue(
                    "UninstallString",
                    "\"" + paths.InstallerCopyPath + "\" /uninstall");
                key.SetValue("InstallLocation", paths.ProductFolder);
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
        }

        private static string Sha256(byte[] bytes)
        {
            using (SHA256 hash = SHA256.Create())
            {
                return BitConverter.ToString(hash.ComputeHash(bytes))
                    .Replace("-", "")
                    .ToLowerInvariant();
            }
        }

        private static void Report(
            Action<string, int> report,
            string message,
            int percent)
        {
            if (report != null)
                report(message, percent);
        }
    }

    internal sealed class Paths
    {
        internal string StartupFolder;
        internal string ProductFolder;
        internal string BackupFolder;
        internal string TemplatePath;
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
                appData = Environment.GetFolderPath(
                    Environment.SpecialFolder.ApplicationData);
                localAppData = Environment.GetFolderPath(
                    Environment.SpecialFolder.LocalApplicationData);
            }

            Paths result = new Paths();
            result.TestMode = !string.IsNullOrEmpty(options.TestRoot);
            result.StartupFolder =
                Path.Combine(appData, "Microsoft", "Word", "STARTUP");
            result.ProductFolder =
                Path.Combine(localAppData, "MasseysLab", Installer.ProductFolderName);
            result.BackupFolder = Path.Combine(result.ProductFolder, "Backups");
            result.TemplatePath =
                Path.Combine(result.StartupFolder, Installer.TemplateName);
            result.InstallerCopyPath =
                Path.Combine(result.ProductFolder, "CleanupSuiteForWord-Setup.exe");
            return result;
        }
    }
}
