using System;
using System.Drawing;
using System.Windows.Forms;

namespace HeicToJpegConverter
{
	public class InstallerForm : Form
	{
		private Button installButton;
		private Button uninstallButton;
		private Label statusLabel;
		private PictureBox iconPictureBox;
		private Label titleLabel;
		private Label descriptionLabel;

		public InstallerForm()
		{
			InitializeComponents();
		}

		private void InitializeComponents()
		{
			// Form properties
			this.Text = "HEIC to JPEG Converter";
			this.Size = new Size(500, 400);
			this.FormBorderStyle = FormBorderStyle.FixedDialog;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.StartPosition = FormStartPosition.CenterScreen;

			// Title
			titleLabel = new Label
			{
				Text = "HEIC to JPEG Converter",
				Font = new Font("Segoe UI", 16, FontStyle.Bold),
				Location = new Point(150, 30),
				AutoSize = true
			};

			// Icon
			iconPictureBox = new PictureBox
			{
				Size = new Size(64, 64),
				Location = new Point(50, 20),
				BorderStyle = BorderStyle.None,
				Image = SystemIcons.Application.ToBitmap()
			};

			// Description
			descriptionLabel = new Label
			{
				Text = "This application adds a context menu option to Windows Explorer " +
					   "that allows you to convert HEIC images to JPEG format with a single click.\n\n" +
					   "The context menu entry will appear when you right-click on HEIC files.",
				Location = new Point(50, 100),
				Size = new Size(400, 100),
				Font = new Font("Segoe UI", 10),
			};

			// Install button
			installButton = new Button
			{
				Text = "Install Context Menu",
				Location = new Point(90, 240),
				Size = new Size(150, 40),
				BackColor = Color.FromArgb(0, 120, 212),
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 10, FontStyle.Bold),
				FlatStyle = FlatStyle.Flat
			};
			installButton.FlatAppearance.BorderSize = 0;
			installButton.Click += InstallButton_Click;

			// Uninstall button
			uninstallButton = new Button
			{
				Text = "Uninstall",
				Location = new Point(260, 240),
				Size = new Size(150, 40),
				BackColor = Color.FromArgb(232, 17, 35),
				ForeColor = Color.White,
				Font = new Font("Segoe UI", 10, FontStyle.Bold),
				FlatStyle = FlatStyle.Flat
			};
			uninstallButton.FlatAppearance.BorderSize = 0;
			uninstallButton.Click += UninstallButton_Click;

			// Status label
			statusLabel = new Label
			{
				Text = "",
				Location = new Point(50, 310),
				Size = new Size(400, 40),
				Font = new Font("Segoe UI", 9),
				TextAlign = ContentAlignment.MiddleCenter
			};

			// Add controls to form
			this.Controls.Add(iconPictureBox);
			this.Controls.Add(titleLabel);
			this.Controls.Add(descriptionLabel);
			this.Controls.Add(installButton);
			this.Controls.Add(uninstallButton);
			this.Controls.Add(statusLabel);
		}

		private void InstallButton_Click(object sender, EventArgs e)
		{
			try
			{
				Program.RegisterContextMenu();
				statusLabel.Text = "Context menu successfully installed!";
				statusLabel.ForeColor = Color.Green;
			}
			catch (Exception ex)
			{
				statusLabel.Text = $"Error: {ex.Message}";
				statusLabel.ForeColor = Color.Red;

				MessageBox.Show($"Failed to install context menu: {ex.Message}\n\nMake sure to run as administrator.",
					"Installation Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}

		private void UninstallButton_Click(object sender, EventArgs e)
		{
			try
			{
				Program.UnregisterContextMenu();
				statusLabel.Text = "Context menu successfully uninstalled!";
				statusLabel.ForeColor = Color.Green;
			}
			catch (Exception ex)
			{
				statusLabel.Text = $"Error: {ex.Message}";
				statusLabel.ForeColor = Color.Red;

				MessageBox.Show($"Failed to uninstall context menu: {ex.Message}\n\nMake sure to run as administrator.",
					"Uninstall Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
		}
	}
}