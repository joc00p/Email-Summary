using System;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace Reporter;

public class SettingsForm : Form
{
    private readonly AppSettings _settings;
    private TextBox _templatePathBox = null!;
    private TextBox _exportDirBox = null!;

    public SettingsForm(AppSettings settings)
    {
        _settings = settings;
        BuildUI();
    }

    private void BuildUI()
    {
        Text = "Settings";
        Size = new Size(600, 220);
        MinimumSize = new Size(500, 220);
        MaximumSize = new Size(int.MaxValue, 220);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;

        var label = new Label
        {
            Text = "PPTX Template:",
            Location = new Point(12, 20),
            AutoSize = true,
        };

        _templatePathBox = new TextBox
        {
            Location = new Point(130, 17),
            Width = 360,
            Text = _settings.TemplatePath,
            ReadOnly = true,
            BackColor = SystemColors.Window,
        };

        var browseBtn = new Button
        {
            Text = "Browse...",
            Location = new Point(498, 15),
            Size = new Size(80, 26),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
        };
        browseBtn.Click += Browse_Click;

        var exportLabel = new Label
        {
            Text = "Export Folder:",
            Location = new Point(12, 60),
            AutoSize = true,
        };

        _exportDirBox = new TextBox
        {
            Location = new Point(130, 57),
            Width = 360,
            Text = _settings.ExportDir,
            ReadOnly = true,
            BackColor = SystemColors.Window,
        };

        var exportBrowseBtn = new Button
        {
            Text = "Browse...",
            Location = new Point(498, 55),
            Size = new Size(80, 26),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(70, 130, 180),
            ForeColor = Color.White,
        };
        exportBrowseBtn.Click += BrowseExportDir_Click;

        var exportHint = new Label
        {
            Text = "Where reports/PPTX exports default to. Leave blank to use the last-used folder.",
            Location = new Point(130, 84),
            AutoSize = true,
            ForeColor = Color.Gray,
            Font = new Font("Segoe UI", 8.5f),
        };

        var saveBtn = new Button
        {
            Text = "Save",
            Location = new Point(410, 140),
            Size = new Size(80, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(16, 137, 62),
            ForeColor = Color.White,
            DialogResult = DialogResult.OK,
        };
        saveBtn.Click += Save_Click;

        var cancelBtn = new Button
        {
            Text = "Cancel",
            Location = new Point(498, 140),
            Size = new Size(80, 28),
            FlatStyle = FlatStyle.Flat,
            BackColor = Color.FromArgb(100, 100, 100),
            ForeColor = Color.White,
            DialogResult = DialogResult.Cancel,
        };

        AcceptButton = saveBtn;
        CancelButton = cancelBtn;

        Controls.AddRange(new Control[]
        {
            label, _templatePathBox, browseBtn,
            exportLabel, _exportDirBox, exportBrowseBtn, exportHint,
            saveBtn, cancelBtn,
        });
    }

    private void BrowseExportDir_Click(object? sender, EventArgs e)
    {
        using var dlg = new FolderBrowserDialog
        {
            Description = "Select the folder where reports and PPTX exports are written",
            UseDescriptionForTitle = true,
        };

        if (Directory.Exists(_exportDirBox.Text))
            dlg.SelectedPath = _exportDirBox.Text;

        if (dlg.ShowDialog() == DialogResult.OK)
            _exportDirBox.Text = dlg.SelectedPath;
    }

    private void Browse_Click(object? sender, EventArgs e)
    {
        using var dlg = new OpenFileDialog
        {
            Title = "Select PPTX Template",
            Filter = "PowerPoint Template|*.pptx|All Files|*.*",
            FileName = _templatePathBox.Text,
        };

        if (File.Exists(_templatePathBox.Text))
            dlg.InitialDirectory = Path.GetDirectoryName(_templatePathBox.Text);

        if (dlg.ShowDialog() == DialogResult.OK)
            _templatePathBox.Text = dlg.FileName;
    }

    private void Save_Click(object? sender, EventArgs e)
    {
        _settings.TemplatePath = _templatePathBox.Text;
        _settings.ExportDir = _exportDirBox.Text;
        _settings.Save();
    }
}
