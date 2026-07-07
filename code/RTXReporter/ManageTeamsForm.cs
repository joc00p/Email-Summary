using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace RTXReporter;

public class ManageTeamsForm : Form
{
    private readonly TeamConfig _config;
    private readonly List<string> _outlookUsers;

    private ComboBox _towerCombo = null!;
    private ListBox _membersBox = null!;
    private ListBox _usersBox = null!;
    private Button _addBtn = null!;
    private Button _removeBtn = null!;

    public ManageTeamsForm(TeamConfig config, IEnumerable<string> outlookUsers)
    {
        _config = config;
        _outlookUsers = outlookUsers.OrderBy(u => u).ToList();
        BuildUI();
        PopulateTower();
    }

    private void BuildUI()
    {
        Text = "Manage Team Towers";
        Size = new Size(700, 480);
        MinimumSize = new Size(600, 400);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);

        // Tower selector row
        var towerLabel = new Label
        {
            Text = "Tower:",
            AutoSize = true,
            Location = new Point(12, 16),
        };

        _towerCombo = new ComboBox
        {
            DropDownStyle = ComboBoxStyle.DropDownList,
            Location = new Point(65, 12),
            Width = 180,
        };
        foreach (var t in TeamConfig.TowerNames)
            _towerCombo.Items.Add(t);
        _towerCombo.SelectedIndex = 0;
        _towerCombo.SelectedIndexChanged += (_, _) => PopulateTower();

        // Members panel (left)
        var membersLabel = new Label
        {
            Text = "Tower Members",
            AutoSize = true,
            Location = new Point(12, 52),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };

        _membersBox = new ListBox
        {
            Location = new Point(12, 72),
            Size = new Size(260, 320),
            SelectionMode = SelectionMode.MultiExtended,
            Sorted = true,
        };
        _membersBox.SelectedIndexChanged += (_, _) => UpdateButtons();

        // Arrow buttons (center)
        _addBtn = new Button
        {
            Text = "◀ Add",
            Location = new Point(285, 160),
            Size = new Size(90, 34),
            BackColor = Color.FromArgb(16, 137, 62),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        _addBtn.Click += AddBtn_Click;

        _removeBtn = new Button
        {
            Text = "Remove ▶",
            Location = new Point(285, 204),
            Size = new Size(90, 34),
            BackColor = Color.FromArgb(180, 60, 30),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        _removeBtn.Click += RemoveBtn_Click;

        // Outlook users panel (right)
        var usersLabel = new Label
        {
            Text = "Outlook Users",
            AutoSize = true,
            Location = new Point(390, 52),
            Font = new Font("Segoe UI", 9f, FontStyle.Bold),
        };

        _usersBox = new ListBox
        {
            Location = new Point(390, 72),
            Size = new Size(280, 320),
            SelectionMode = SelectionMode.MultiExtended,
            Sorted = true,
        };
        _usersBox.SelectedIndexChanged += (_, _) => UpdateButtons();

        // Bottom buttons
        var saveBtn = new Button
        {
            Text = "Save",
            DialogResult = DialogResult.None,
            Location = new Point(490, 410),
            Size = new Size(85, 30),
            BackColor = Color.FromArgb(0, 84, 166),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };
        saveBtn.Click += (_, _) =>
        {
            _config.Save();
            MessageBox.Show("Team configuration saved.", "Saved",
                MessageBoxButtons.OK, MessageBoxIcon.Information);
        };

        var closeBtn = new Button
        {
            Text = "Close",
            DialogResult = DialogResult.OK,
            Location = new Point(585, 410),
            Size = new Size(85, 30),
            BackColor = Color.FromArgb(100, 100, 100),
            ForeColor = Color.White,
            FlatStyle = FlatStyle.Flat,
        };

        Controls.AddRange(new Control[]
        {
            towerLabel, _towerCombo,
            membersLabel, _membersBox,
            _addBtn, _removeBtn,
            usersLabel, _usersBox,
            saveBtn, closeBtn,
        });

        AcceptButton = saveBtn;
    }

    private void PopulateTower()
    {
        string tower = _towerCombo.SelectedItem as string ?? "";

        _membersBox.BeginUpdate();
        _membersBox.Items.Clear();
        if (_config.Teams.TryGetValue(tower, out var members))
            foreach (var m in members.OrderBy(x => x))
                _membersBox.Items.Add(m);
        _membersBox.EndUpdate();

        _usersBox.BeginUpdate();
        _usersBox.Items.Clear();
        var memberSet = new HashSet<string>(
            members ?? Enumerable.Empty<string>(),
            StringComparer.OrdinalIgnoreCase);

        foreach (var user in _outlookUsers)
        {
            if (memberSet.Contains(user)) continue;
            string otherTower = _config.GetTeam(user);
            string display = otherTower == "Other" || otherTower == tower
                ? user
                : $"{user}  ({otherTower})";
            _usersBox.Items.Add(display);
        }
        _usersBox.EndUpdate();

        UpdateButtons();
    }

    private void AddBtn_Click(object? sender, EventArgs e)
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        var selected = _usersBox.SelectedItems.Cast<string>().ToList();
        foreach (var item in selected)
        {
            // Strip the "(Tower)" suffix if present
            string name = item.Contains("  (") ? item[..item.IndexOf("  (")] : item;
            _config.AddMember(tower, name);
        }
        PopulateTower();
    }

    private void RemoveBtn_Click(object? sender, EventArgs e)
    {
        string tower = _towerCombo.SelectedItem as string ?? "";
        var selected = _membersBox.SelectedItems.Cast<string>().ToList();
        foreach (var name in selected)
            _config.RemoveMember(tower, name);
        PopulateTower();
    }

    private void UpdateButtons()
    {
        _addBtn.Enabled = _usersBox.SelectedItems.Count > 0;
        _removeBtn.Enabled = _membersBox.SelectedItems.Count > 0;
    }
}
