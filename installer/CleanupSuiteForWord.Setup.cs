using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Threading;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("CleanupSuite For Word Setup")]
[assembly: AssemblyCompany("MasseysLab")]
[assembly: AssemblyProduct("CleanupSuite For Word")]
[assembly: AssemblyCopyright("Copyright (c) 2026 MasseysLab")]
[assembly: AssemblyVersion("0.9.2.0")]
[assembly: AssemblyFileVersion("0.9.2.0")]

namespace MasseysLab.CleanupSuiteForWord.Setup
{
    internal static class Program
    {
        internal const string ProductName = "CleanupSuite For Word";
        internal const string DisplayVersion = "0.9.2-alpha";
        internal const string ManifestUrl = "https://raw.githubusercontent.com/MasseysLab/CleanupSuiteForWord/main/release/latest-release.ini";
        internal const int BackupLimit = 3;

        [STAThread]
        private static int Main(string[] args)
        {
            ServicePointManager.SecurityProtocol = (SecurityProtocolType)3072;
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
                    MessageBox.Show("Setup could not finish.\r\n\r\n" + ex.Message,
                        ProductName, MessageBoxButtons.OK, MessageBoxIcon.Error);
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
                if (arg.Equals("/silent", StringComparison.OrdinalIgnoreCase)) result.Silent = true;
                if (arg.Equals("/uninstall", StringComparison.OrdinalIgnoreCase)) result.Uninstall = true;
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
        private readonly Label status;
        private readonly ProgressBar progress;
        private readonly Button installButton;

        internal SetupForm(Options setupOptions)
        {
            options = setupOptions;
            Text = Program.ProductName + " Setup";
            ClientSize = new Size(510, 238);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            StartPosition = FormStartPosition.CenterScreen;
            Font = new Font("Segoe UI", 9F);

            Label title = new Label();
            title.Text = "CleanupSuite For Word";
            title.Font = new Font("Segoe UI Semibold", 18F, FontStyle.Bold);
            title.Location = new Point(24, 20);
            title.Size = new Size(460, 38);
            Controls.Add(title);

            Label explanation = new Label();
            explanation.Text = "Installs the current CleanupSuite template for this Windows account. " +
                "Word must be closed. Existing installations are backed up automatically.";
            explanation.Location = new Point(27, 67);
            explanation.Size = new Size(455, 48);
            Controls.Add(explanation);

            status = new Label();
            status.Text = "Ready to install version " + Program.DisplayVersion + ".";
            status.Location = new Point(27, 126);
            status.Size = new Size(455, 22);
            Controls.Add(status);

            progress = new ProgressBar();
            progress.Location = new Point(28, 151);
            progress.Size = new Size(454, 8);
            progress.Style = ProgressBarStyle.Continuous;
            progress.Minimum = 0;
            progress.Maximum = 100;
            Controls.Add(progress);

            installButton = new Button();
            installButton.Text = "Install";
            installButton.Font = new Font("Segoe UI Semibold", 9F, FontStyle.Bold);
            installButton.Location = new Point(286, 183);
            installButton.Size = new Size(94, 32);
            installButton.Click += InstallClicked;
            Controls.Add(installButton);

            Button cancelButton = new Button();
            cancelButton.Text = "Cancel";
            cancelButton.Location = new Point(388, 183);
            cancelButton.Size = new Size(94, 32);
            cancelButton.Click += delegate { Close(); };
            Controls.Add(cancelButton);
            CancelButton = cancelButton;
            AcceptButton = installButton;
        }

        private void InstallClicked(object sender, EventArgs e)
        {
            installButton.Enabled = false;
            progress.Value = 12;
            status.Text = "Checking the current release...";
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
                        status.Text = "Installed. Reopen Word and select CleanupSuite on the ribbon.";
                        progress.Value = 100;
                        installButton.Text = "Close";
                        installButton.Enabled = true;
                        installButton.Click -= InstallClicked;
                        installButton.Click += delegate { Close(); };
                    }
                    else
                    {
                        status.Text = "Setup did not make changes. Close Word and try again.";
                        progress.Value = 0;
                        installButton.Enabled = true;
                    }
                });
            });
            worker.IsBackground = true;
            worker.SetApartmentState(ApartmentState.STA);
            worker.Start();
        }
    }

    internal static class Installer
    {
        internal const string TemplateName = "CleanupSuite.dotm";
        internal const string ProductFolderName = "CleanupSuiteForWord";
        private const string UninstallKeyPath = @"Software\Microsoft\Windows\CurrentVersion\Uninstall\CleanupSuiteForWord";

        internal static int Install(Options options, Action<string, int> report)
        {
            try
            {
                if (string.IsNullOrEmpty(options.TestRoot) && Process.GetProcessesByName("WINWORD").Length > 0)
                    throw new InvalidOperationException("Microsoft Word is still open. Close every Word window, then run Setup again.");

                Paths paths = Paths.Create(options);
                Directory.CreateDirectory(paths.StartupFolder);
                Directory.CreateDirectory(paths.ProductFolder);
                Directory.CreateDirectory(paths.BackupFolder);

                Report(report, "Checking the current release...", 18);
                TemplatePayload payload = GetTemplatePayload(paths);
                Report(report, "Verifying the template...", 42);
                if (!string.IsNullOrEmpty(payload.ExpectedSha256))
                {
                    string actualHash = Sha256(payload.Bytes);
                    if (!actualHash.Equals(payload.ExpectedSha256, StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("The downloaded template did not match its published SHA-256 hash.");
                }

                Report(report, "Backing up the previous installation...", 60);
                BackupExisting(paths);
                RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);

                Report(report, "Installing CleanupSuite for this account...", 76);
                AtomicWrite(paths.TemplatePath, payload.Bytes);
                File.Copy(Application.ExecutablePath, paths.InstallerCopyPath, true);
                RegisterUninstall(paths, payload.Version);

                Report(report, "Installation complete.", 100);
                if (!options.Silent && report == null)
                    MessageBox.Show("CleanupSuite was installed for this Windows account.\r\n\r\n" +
                        "Reopen Word and select CleanupSuite on the ribbon.", Program.ProductName,
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                return 0;
            }
            catch (Exception ex)
            {
                if (!options.Silent)
                    MessageBox.Show("Setup did not make changes.\r\n\r\n" + ex.Message,
                        Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }
        }

        internal static int Uninstall(Options options)
        {
            if (string.IsNullOrEmpty(options.TestRoot) && Process.GetProcessesByName("WINWORD").Length > 0)
            {
                if (!options.Silent)
                    MessageBox.Show("Close Microsoft Word before uninstalling CleanupSuite.",
                        Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return 1;
            }

            if (!options.Silent && MessageBox.Show(
                    "Remove CleanupSuite For Word from this Windows account?\r\n\r\n" +
                    "Your three newest installer backups will be kept.", Program.ProductName,
                    MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes)
                return 2;

            try
            {
                Paths paths = Paths.Create(options);
                Directory.CreateDirectory(paths.BackupFolder);
                BackupExisting(paths);
                RetainNewestBackups(paths.BackupFolder, Program.BackupLimit);
                if (File.Exists(paths.TemplatePath)) File.Delete(paths.TemplatePath);
                if (!paths.TestMode) Registry.CurrentUser.DeleteSubKeyTree(UninstallKeyPath, false);
                if (!options.Silent)
                    MessageBox.Show("CleanupSuite was removed. Installer backups were kept at:\r\n" +
                        paths.BackupFolder, Program.ProductName, MessageBoxButtons.OK, MessageBoxIcon.Information);
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

        private static TemplatePayload GetTemplatePayload(Paths paths)
        {
            try
            {
                using (WebClient client = NewWebClient())
                {
                    string manifest = client.DownloadString(Program.ManifestUrl);
                    Dictionary<string, string> values = ParseManifest(manifest);
                    string templateUrl = Require(values, "template_url");
                    if (!templateUrl.StartsWith("https://github.com/MasseysLab/CleanupSuiteForWord/", StringComparison.OrdinalIgnoreCase))
                        throw new InvalidDataException("The release manifest contained an unexpected download location.");
                    byte[] bytes = client.DownloadData(templateUrl);
                    return new TemplatePayload(bytes, Require(values, "template_sha256"), Require(values, "version"));
                }
            }
            catch
            {
                byte[] embedded = ReadEmbeddedTemplate();
                return new TemplatePayload(embedded, Sha256(embedded), Program.DisplayVersion);
            }
        }

        private static WebClient NewWebClient()
        {
            WebClient client = new WebClient();
            client.Headers[HttpRequestHeader.UserAgent] = "CleanupSuiteForWord-Setup/" + Program.DisplayVersion;
            return client;
        }

        private static byte[] ReadEmbeddedTemplate()
        {
            using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream("CleanupSuite.dotm"))
            {
                if (stream == null) throw new FileNotFoundException("The embedded CleanupSuite template was not found.");
                using (MemoryStream output = new MemoryStream())
                {
                    stream.CopyTo(output);
                    return output.ToArray();
                }
            }
        }

        private static Dictionary<string, string> ParseManifest(string text)
        {
            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string raw in text.Replace("\r", "").Split('\n'))
            {
                string line = raw.Trim();
                if (line.Length == 0 || line.StartsWith("#")) continue;
                int equals = line.IndexOf('=');
                if (equals > 0) values[line.Substring(0, equals).Trim()] = line.Substring(equals + 1).Trim();
            }
            return values;
        }

        private static string Require(Dictionary<string, string> values, string key)
        {
            string value;
            if (!values.TryGetValue(key, out value) || string.IsNullOrWhiteSpace(value))
                throw new InvalidDataException("The release manifest did not include " + key + ".");
            return value;
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
            string temporary = targetPath + ".new";
            File.WriteAllBytes(temporary, bytes);
            try
            {
                if (File.Exists(targetPath)) File.Replace(temporary, targetPath, null);
                else File.Move(temporary, targetPath);
            }
            finally
            {
                if (File.Exists(temporary)) File.Delete(temporary);
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
                key.SetValue("InstallLocation", paths.ProductFolder);
                key.SetValue("NoModify", 1, RegistryValueKind.DWord);
                key.SetValue("NoRepair", 1, RegistryValueKind.DWord);
            }
        }

        private static string Sha256(byte[] bytes)
        {
            using (SHA256 hash = SHA256.Create())
            {
                return BitConverter.ToString(hash.ComputeHash(bytes)).Replace("-", "").ToLowerInvariant();
            }
        }

        private static void Report(Action<string, int> report, string message, int percent)
        {
            if (report != null) report(message, percent);
        }
    }

    internal sealed class TemplatePayload
    {
        internal readonly byte[] Bytes;
        internal readonly string ExpectedSha256;
        internal readonly string Version;

        internal TemplatePayload(byte[] bytes, string expectedSha256, string version)
        {
            Bytes = bytes;
            ExpectedSha256 = expectedSha256;
            Version = version;
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
                appData = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                localAppData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            }

            Paths result = new Paths();
            result.TestMode = !string.IsNullOrEmpty(options.TestRoot);
            result.StartupFolder = Path.Combine(appData, "Microsoft", "Word", "STARTUP");
            result.ProductFolder = Path.Combine(localAppData, "MasseysLab", Installer.ProductFolderName);
            result.BackupFolder = Path.Combine(result.ProductFolder, "Backups");
            result.TemplatePath = Path.Combine(result.StartupFolder, Installer.TemplateName);
            result.InstallerCopyPath = Path.Combine(result.ProductFolder, "CleanupSuiteForWord-Setup.exe");
            return result;
        }
    }
}
