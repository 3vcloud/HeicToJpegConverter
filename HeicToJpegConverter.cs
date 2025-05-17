using System;
using System.Diagnostics;
using System.IO;
using System.Windows.Forms;
using System.Drawing;
using System.Drawing.Imaging;
using Microsoft.Win32;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Reflection;
using ImageMagick;

namespace HeicToJpegConverter
{
	internal static class Program
	{
		private const string AppId = "HeicToJpegConverter";

		[STAThread]
		static void Main(string[] args)
		{
			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault(false);

			// Check for uninstall command
			if (args.Length > 0 && (args[0].Equals("/uninstall", StringComparison.OrdinalIgnoreCase) ||
								   args[0].Equals("-uninstall", StringComparison.OrdinalIgnoreCase)))
			{
				try
				{
					UnregisterContextMenu();
					MessageBox.Show("HEIC to JPEG converter has been successfully uninstalled!",
						"Uninstall Complete", MessageBoxButtons.OK, MessageBoxIcon.Information);
				}
				catch (Exception ex)
				{
					MessageBox.Show($"Failed to uninstall: {ex.Message}\n\nMake sure to run as administrator.",
						"Uninstall Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
				}
				return;
			}

			// If no conversion arguments provided, show installer UI
			if (args.Length == 0)
			{
				Application.Run(new InstallerForm());
				return;
			}

			// If arguments are provided, we're being called from the context menu to convert files
			try
			{
				ConvertFiles(args);
			}
			catch (Exception ex)
			{
				MessageBox.Show($"Error converting files: {ex.Message}",
					"Conversion Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		// Public methods for use by the InstallerForm
		public static void RegisterContextMenu()
		{
			// Get the path to the executable
			string executablePath = $"\"{Application.ExecutablePath}\"";

			// Register for .heic files
			using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.heic\shell\ConvertToJPEG"))
			{
				key.SetValue("", "Convert to JPEG");
				key.SetValue("Icon", executablePath);

				using (RegistryKey commandKey = key.CreateSubKey("command"))
				{
					commandKey.SetValue("", $"{executablePath} \"%1\"");
				}
			}

			// Register for multiple .heic file selection
			using (RegistryKey key = Registry.ClassesRoot.CreateSubKey(@"SystemFileAssociations\.heic\shell\ConvertMultipleToJPEG"))
			{
				key.SetValue("", "Convert Selected to JPEG");
				key.SetValue("Icon", executablePath);
				key.SetValue("MultiSelectModel", "Player");

				using (RegistryKey commandKey = key.CreateSubKey("command"))
				{
					commandKey.SetValue("", $"{executablePath} \"%1\"");
				}
			}

			// Register application in Add/Remove Programs
			RegisterInAddRemovePrograms();
		}

		public static void UnregisterContextMenu()
		{
			// Remove .heic file association
			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree(@"SystemFileAssociations\.heic\shell\ConvertToJPEG", false);
			}
			catch (Exception) { /* Ignore if key doesn't exist */ }

			try
			{
				Registry.ClassesRoot.DeleteSubKeyTree(@"SystemFileAssociations\.heic\shell\ConvertMultipleToJPEG", false);
			}
			catch (Exception) { /* Ignore if key doesn't exist */ }

			// Remove from Add/Remove Programs
			UnregisterFromAddRemovePrograms();
		}

		private static void RegisterInAddRemovePrograms()
		{
			string appName = "HEIC to JPEG Converter";
			string appVersion = Assembly.GetExecutingAssembly().GetName().Version.ToString();
			string appExe = Application.ExecutablePath;
			string uninstallString = $"\"{appExe}\" /uninstall";

			using (RegistryKey key = Registry.LocalMachine.CreateSubKey(
				$@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}"))
			{
				key.SetValue("DisplayName", appName);
				key.SetValue("DisplayVersion", appVersion);
				key.SetValue("DisplayIcon", appExe);
				key.SetValue("UninstallString", uninstallString);
				key.SetValue("Publisher", "HEIC Converter");
				key.SetValue("NoModify", 1);
				key.SetValue("NoRepair", 1);
				key.SetValue("EstimatedSize", 5000); // Size in KB (approx.)
				key.SetValue("InstallLocation", Path.GetDirectoryName(appExe));
			}
		}

		private static void UnregisterFromAddRemovePrograms()
		{
			try
			{
				Registry.LocalMachine.DeleteSubKeyTree($@"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\{AppId}", false);
			}
			catch (Exception) { /* Ignore if key doesn't exist */ }
		}

        private static void ConvertFiles(string[] filePaths)
		{
			// Check if files exist and are HEIC
			List<string> validFiles = new List<string>();
			foreach (string filePath in filePaths)
			{
				if (File.Exists(filePath) &&
					Path.GetExtension(filePath).Equals(".heic", StringComparison.OrdinalIgnoreCase))
				{
					validFiles.Add(filePath);
				}
			}

			if (validFiles.Count == 0)
			{
				MessageBox.Show("No valid HEIC files found.", "Information",
					MessageBoxButtons.OK, MessageBoxIcon.Information);
				return;
			}

			// Show progress form for multiple files
			using (var progressForm = new ProgressForm(validFiles.Count))
			{
				// Start conversion in background
				Task.Run(() =>
				{
					int converted = 0;
					int failed = 0;

					foreach (string file in validFiles)
					{
						try
						{
							string outputPath = Path.ChangeExtension(file, ".jpg");

							// Use ImageMagick for conversion
							using (var image = new MagickImage(file))
							{
								// Set JPEG quality
								image.Quality = 90;
								image.Write(outputPath);
							}

							converted++;
							progressForm.UpdateProgress(converted + failed);
						}
						catch
						{
							failed++;
							progressForm.UpdateProgress(converted + failed);
						}
					}

					// Close the form when done
					progressForm.InvokeIfRequired(() =>
					{
						progressForm.Close();
						MessageBox.Show($"Conversion complete!\nSuccessfully converted: {converted}\nFailed: {failed}",
							"Conversion Results", MessageBoxButtons.OK, MessageBoxIcon.Information);
					});
				});

				progressForm.ShowDialog();
			}
		}
	}

	// Form to show conversion progress
	public class ProgressForm : Form
{
	private ProgressBar progressBar;
	private Label statusLabel;
	private int totalFiles;

	public ProgressForm(int totalFiles)
	{
		this.totalFiles = totalFiles;
		InitializeComponents();
	}

	private void InitializeComponents()
	{
		this.Size = new Size(400, 150);
		this.FormBorderStyle = FormBorderStyle.FixedDialog;
		this.MaximizeBox = false;
		this.MinimizeBox = false;
		this.StartPosition = FormStartPosition.CenterScreen;
		this.Text = "Converting HEIC to JPEG";

		statusLabel = new Label
		{
			Text = $"Converting 0 of {totalFiles} files...",
			AutoSize = true,
			Location = new Point(20, 20)
		};

		progressBar = new ProgressBar
		{
			Minimum = 0,
			Maximum = totalFiles,
			Value = 0,
			Width = 360,
			Location = new Point(20, 60)
		};

		this.Controls.Add(statusLabel);
		this.Controls.Add(progressBar);
	}

	public void UpdateProgress(int completed)
	{
		InvokeIfRequired(() =>
		{
			progressBar.Value = Math.Min(completed, totalFiles);
			statusLabel.Text = $"Converting {completed} of {totalFiles} files...";
		});
	}

	public void InvokeIfRequired(Action action)
	{
		if (this.InvokeRequired)
		{
			this.Invoke(action);
		}
		else
		{
			action();
		}
	}
}
}