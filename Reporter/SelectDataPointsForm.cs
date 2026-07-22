using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;

namespace Reporter;

// Shown between extraction and final generation. Each tower's candidate data points are listed with
// checkboxes; the user can include/exclude, EDIT the text inline (double-click / F2), and DRAG to
// reorder (top = higher priority, which drives the order in the report). Top N are pre-checked.
public class SelectDataPointsForm : Form
{
    private readonly List<(string Tower, ListView List)> _lists = new();

    public Dictionary<string, List<string>> SelectedByTower { get; } = new();

    public SelectDataPointsForm(Dictionary<string, List<string>> candidates, int defaultChecked)
    {
        Text = "Select & Arrange Data Points";
        Size = new Size(720, 760);
        MinimumSize = new Size(520, 440);
        StartPosition = FormStartPosition.CenterParent;
        Font = new Font("Segoe UI", 10f);

        var flow = new FlowLayoutPanel
        {
            Dock = DockStyle.Fill,
            FlowDirection = FlowDirection.TopDown,
            WrapContents = false,
            AutoScroll = true,
            Padding = new Padding(14, 10, 14, 10),
        };

        flow.Controls.Add(new Label
        {
            Text = "Check the points to include. Double-click a point to edit its text. Drag points to reorder — the order here is the order in the report (top = most important).",
            AutoSize = true,
            MaximumSize = new Size(650, 0),
            Margin = new Padding(0, 0, 0, 8),
            ForeColor = Color.FromArgb(80, 80, 80),
        });

        foreach (var tower in candidates.Keys)
        {
            flow.Controls.Add(new Label
            {
                Text = tower,
                Font = new Font("Segoe UI", 11f, FontStyle.Bold),
                ForeColor = Color.FromArgb(0, 84, 166),
                AutoSize = true,
                Margin = new Padding(0, 12, 0, 4),
            });

            var lv = MakeList();
            int i = 0;
            foreach (var point in candidates[tower])
                lv.Items.Add(new ListViewItem(point) { Checked = i++ < defaultChecked });
            lv.Columns[0].Width = 640;
            lv.Height = Math.Max(1, lv.Items.Count) * 24 + 8;
            flow.Controls.Add(lv);
            _lists.Add((tower, lv));
        }

        var buttonBar = new Panel { Dock = DockStyle.Bottom, Height = 50 };
        var generate = new Button
        {
            Text = "Generate", DialogResult = DialogResult.OK, Size = new Size(110, 32),
            BackColor = Color.FromArgb(16, 137, 62), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
        };
        var cancel = new Button
        {
            Text = "Cancel", DialogResult = DialogResult.Cancel, Size = new Size(90, 32),
            BackColor = Color.FromArgb(100, 100, 100), ForeColor = Color.White, FlatStyle = FlatStyle.Flat,
        };
        generate.Click += (_, _) =>
        {
            foreach (var (tower, lv) in _lists)
            {
                var picked = lv.Items.Cast<ListViewItem>()
                    .Where(it => it.Checked)
                    .Select(it => it.Text.Trim())
                    .Where(t => t.Length > 0)
                    .ToList();
                if (picked.Count > 0) SelectedByTower[tower] = picked;
            }
        };
        buttonBar.Controls.Add(generate);
        buttonBar.Controls.Add(cancel);
        void Place() { generate.Location = new Point(buttonBar.Width - 225, 9); cancel.Location = new Point(buttonBar.Width - 105, 9); }
        buttonBar.Resize += (_, _) => Place();
        Place();

        AcceptButton = generate;
        CancelButton = cancel;
        Controls.Add(flow);
        Controls.Add(buttonBar);
    }

    private static ListView MakeList()
    {
        var lv = new ListView
        {
            View = View.Details,
            CheckBoxes = true,
            LabelEdit = true,
            FullRowSelect = true,
            MultiSelect = false,
            HeaderStyle = ColumnHeaderStyle.None,
            AllowDrop = true,
            Width = 656,
            BorderStyle = BorderStyle.FixedSingle,
        };
        lv.Columns.Add("point");

        // Double-click (or F2) edits the text
        lv.DoubleClick += (_, _) => { if (lv.SelectedItems.Count > 0) lv.SelectedItems[0].BeginEdit(); };

        // Drag to reorder within the tower's list
        lv.ItemDrag += (_, e) => lv.DoDragDrop(e.Item!, DragDropEffects.Move);
        lv.DragEnter += (_, e) => e.Effect = DragDropEffects.Move;
        lv.DragOver += (_, e) =>
        {
            e.Effect = DragDropEffects.Move;
            var p = lv.PointToClient(new Point(e.X, e.Y));
            var over = lv.GetItemAt(p.X, p.Y);
            if (over == null) { lv.InsertionMark.Index = -1; return; }
            var r = over.Bounds;
            lv.InsertionMark.AppearsAfterItem = p.Y > r.Top + r.Height / 2;
            lv.InsertionMark.Index = over.Index;
        };
        lv.DragDrop += (_, e) =>
        {
            if (e.Data?.GetData(typeof(ListViewItem)) is not ListViewItem dragged) return;
            var p = lv.PointToClient(new Point(e.X, e.Y));
            var over = lv.GetItemAt(p.X, p.Y);
            bool after = false;
            if (over != null) { var r = over.Bounds; after = p.Y > r.Top + r.Height / 2; }

            lv.Items.Remove(dragged);
            int insertAt = over == null ? lv.Items.Count : over.Index + (after ? 1 : 0);
            insertAt = Math.Max(0, Math.Min(insertAt, lv.Items.Count));
            lv.Items.Insert(insertAt, dragged);
            lv.InsertionMark.Index = -1;
        };
        return lv;
    }
}
