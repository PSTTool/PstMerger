using System;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace PstMerger
{
    public partial class MainForm : Form
    {
        private PstService _pstService;
        private System.Threading.CancellationTokenSource _cts;
        private string _logFile;

        public MainForm()
        {
            InitializeComponent();
            this.Text = string.Format("PST Merge Tool v{0}", Application.ProductVersion);
            _pstService = new PstService();
            _logFile = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, string.Format("PstMerge_{0:yyyyMMdd_HHmmss}.log", DateTime.Now));
            Log("Tool initialized. Enterprise Log started: " + _logFile);
        }

        private void btnBrowseSource_Click(object sender, EventArgs e)
        {
            using (var fbd = new FolderBrowserDialog())
            {
                if (fbd.ShowDialog() == DialogResult.OK)
                {
                    txtSourceFolder.Text = fbd.SelectedPath;
                }
            }
        }

        private void btnBrowseDest_Click(object sender, EventArgs e)
        {
            using (var sfd = new SaveFileDialog())
            {
                sfd.Filter = "Outlook Data File (*.pst)|*.pst";
                if (sfd.ShowDialog() == DialogResult.OK)
                {
                    txtDestPst.Text = sfd.FileName;
                }
            }
        }

        private void btnFixRegistry_Click(object sender, EventArgs e)
        {
            try
            {
                Log("Applying PST size limit fixes to registry...");
                
                // We target Outlook 15.0 and 16.0
                string[] versions = { "15.0", "16.0" };
                foreach (var v in versions)
                {
                    string keyPath = string.Format(@"Software\Microsoft\Office\{0}\Outlook\PST", v);
                    using (RegistryKey key = Registry.CurrentUser.CreateSubKey(keyPath))
                    {
                        if (key != null)
                        {
                            // Values in MB. 2000000 MB = ~2 TB (effectively unlimited)
                            key.SetValue("MaxLargeFileSize", 2000000, RegistryValueKind.DWord);
                            key.SetValue("WarnLargeFileSize", 1900000, RegistryValueKind.DWord);
                        }
                    }
                }

                Log("SUCCESS: PST size limits increased to 2TB (effectively unlimited). Please restart Outlook if it's open.");
                MessageBox.Show("Registry updated. Please restart Outlook if it is currently running.", "Registry Fix Applied", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                Log("ERROR applying registry fix: " + ex.Message);
                MessageBox.Show("Failed to update registry. You may need to run as Administrator.\n\n" + ex.Message, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private async void btnStartMerge_Click(object sender, EventArgs e)
        {
            string sourceDir = txtSourceFolder.Text;
            string destFile = txtDestPst.Text;

            if (string.IsNullOrEmpty(sourceDir) || !Directory.Exists(sourceDir))
            {
                MessageBox.Show("Please select a valid source folder.");
                return;
            }

            if (string.IsNullOrEmpty(destFile))
            {
                MessageBox.Show("Please select a destination PST file.");
                return;
            }

            // Check Disk Space
            try
            {
                string drive = Path.GetPathRoot(Path.GetFullPath(destFile));
                DriveInfo di = new DriveInfo(drive);
                long totalSourceSize = 0;
                var pstFilesCheck = Directory.GetFiles(sourceDir, "*.pst", SearchOption.TopDirectoryOnly);
                foreach (var f in pstFilesCheck) totalSourceSize += new FileInfo(f).Length;
                
                if (di.AvailableFreeSpace < (totalSourceSize * 1.1)) // 10% buffer
                {
                    var msg = string.Format("Warning: You might not have enough disk space on {0}.\nAvailable: {1} GB\nRequired (est): {2} GB\n\nContinue anyway?", 
                        drive, di.AvailableFreeSpace / 1024 / 1024 / 1024, (totalSourceSize * 1.1) / 1024 / 1024 / 1024);
                    if (MessageBox.Show(msg, "Disk Space Warning", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) == DialogResult.No)
                        return;
                }
            }
            catch { }

            var pstFiles = Directory.GetFiles(sourceDir, "*.pst", SearchOption.TopDirectoryOnly);
            if (pstFiles.Length == 0)
            {
                MessageBox.Show("No PST files found in source folder.");
                return;
            }

            btnStartMerge.Enabled = false;
            btnFixRegistry.Enabled = false;
            btnCancel.Visible = true;
            btnCancel.Enabled = true;
            progressBar.Value = 0;
            progressBar.Maximum = pstFiles.Length;

            Log(string.Format("Starting merge of {0} files...", pstFiles.Length));

            try
            {
                _cts = new System.Threading.CancellationTokenSource();
                await Task.Run(() => 
                {
                    _pstService.MergeFiles(pstFiles, destFile, _cts.Token, (progress, message) => 
                    {
                        this.Invoke(new Action(() => 
                        {
                            Log(message);
                            if (progress > 0) progressBar.Value = Math.Min(progress, progressBar.Maximum);
                        }));
                    });
                });

                if (_cts.Token.IsCancellationRequested)
                {
                    Log("STOPPED: Merge was cancelled by user.");
                    MessageBox.Show("Process was cancelled.", "Stopped", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                }
                else
                {
                    Log("COMPLETED: All PST files merged successfully.");
                    MessageBox.Show("Merge completed successfully!", "Success", MessageBoxButtons.OK, MessageBoxIcon.Information);
                }
            }
            catch (OperationCanceledException)
            {
                Log("STOPPED: Merge was cancelled.");
            }
            catch (Exception ex)
            {
                Log("FATAL ERROR: " + ex.Message);
                MessageBox.Show("An error occurred during the merge:\n\n" + ex.Message, "Merge Failed", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                btnStartMerge.Enabled = true;
                btnFixRegistry.Enabled = true;
                btnCancel.Visible = false;
                if (_cts != null) _cts.Dispose();
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            if (_cts != null)
            {
                Log("Cancellation requested... waiting for current item to finish...");
                _cts.Cancel();
                btnCancel.Enabled = false;
            }
        }

        private void Log(string message)
        {
            if (txtLog.InvokeRequired)
            {
                txtLog.Invoke(new Action(() => Log(message)));
                return;
            }
            string line = string.Format("[{0:HH:mm:ss}] {1}", DateTime.Now, message);
            txtLog.AppendText(line + Environment.NewLine);
            
            // Persistent File Logging
            try { File.AppendAllText(_logFile, line + Environment.NewLine); } catch { }
        }

        private void btnAbout_Click(object sender, EventArgs e)
        {
            Version v = new Version(Application.ProductVersion);
            string displayVersion = string.Format("{0}.{1}.{2}", v.Major, v.Minor, v.Build);

            string about = string.Format("PST Merge Tool v{0}\n\n", displayVersion) +
                           "Developed by: Mithun\n" +
                           "© DataGuardNXT 2026\n\n" +
                           "All Rights Reserved.\n\n" +
                           "Enterprise-grade PST merging solution\n" +
                           "for large mailbox recovery.";
            MessageBox.Show(about, "About PST Merge Tool", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
