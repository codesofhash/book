using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

public class AutoComplete : IDisposable
{
    private ListBox listBox;
    private List<string> allSuggestions;
    private bool isUpdatingText = false;
    private Control parentControl;
    private int activationId = 0; // Tracks activation to prevent stale cancel events

    public TextBox TextBox { get; private set; }

    public delegate void ItemSelectedEventHandler(string selectedItem, Keys commitKey);
    public event ItemSelectedEventHandler ItemSelected;
    public event Action EditCancelled;
    public event Action<MouseEventArgs> TextBoxMouseDown;

    public AutoComplete(Control parent, Point location, int width, Font font, List<string> suggestions)
    {
        this.parentControl = parent;

        // Create and configure the TextBox
        this.TextBox = new TextBox
        {
            Location = location,
            Width = width,
            Font = font,
            Visible = false,
            BackColor = System.Drawing.Color.LightYellow,  // Debug: make it visible
            BorderStyle = BorderStyle.FixedSingle
        };
        parent.Controls.Add(this.TextBox);

        this.allSuggestions = suggestions;

        // Create and configure the ListBox
        listBox = new ListBox
        {
            Visible = false,
            Width = this.TextBox.Width,
            Left = this.TextBox.Left,
            Top = this.TextBox.Top + this.TextBox.Height,
            Font = font,
            IntegralHeight = false,
            Height = 150
        };

        // Add ListBox to the same parent
        parent.Controls.Add(listBox);
        listBox.BringToFront();
        this.TextBox.BringToFront();

        // Attach event handlers
        this.TextBox.TextChanged += TextBox_TextChanged;
        this.TextBox.KeyDown += TextBox_KeyDown;
        this.TextBox.PreviewKeyDown += TextBox_PreviewKeyDown;
        this.TextBox.LostFocus += TextBox_LostFocus;
        this.TextBox.MouseDown += TextBox_MouseDown;
        listBox.KeyDown += ListBox_KeyDown;
        listBox.DoubleClick += ListBox_DoubleClick;
        listBox.LostFocus += ListBox_LostFocus;
    }

    private void TextBox_MouseDown(object sender, MouseEventArgs e)
    {
        // Raise event so parent can handle clicks that might need to go elsewhere
        TextBoxMouseDown?.Invoke(e);
    }

    public string Text
    {
        get => TextBox.Text;
        set
        {
            isUpdatingText = true;
            TextBox.Text = value;
            isUpdatingText = false;
        }
    }

    public void SetPosition(Point location, int width)
    {
        TextBox.Location = location;
        TextBox.Width = width;
        listBox.Left = location.X;
        listBox.Top = location.Y + TextBox.Height;
        listBox.Width = width;
    }

    public void Clear()
    {
        isUpdatingText = true;
        TextBox.Text = "";
        isUpdatingText = false;
        listBox.Items.Clear();
        listBox.Visible = false;
    }

    public void Activate()
    {
        activationId++; // Increment to invalidate any pending cancel events
        TextBox.Visible = true;
        TextBox.BringToFront();
        listBox.BringToFront();
        TextBox.Focus();
        TextBox.SelectAll();

        // Show all suggestions immediately on activation
        ShowSuggestions(TextBox.Text);
    }

    public void Hide()
    {
        TextBox.Visible = false;
        listBox.Visible = false;
    }

    private void TextBox_LostFocus(object sender, EventArgs e)
    {
        // Use BeginInvoke to check focus after it has fully transferred
        if (parentControl != null && !parentControl.IsDisposed)
        {
            int currentActivationId = activationId; // Capture current activation
            parentControl.BeginInvoke(new Action(() =>
            {
                // Only fire EditCancelled if no new activation happened
                if (currentActivationId == activationId &&
                    !TextBox.Focused && !listBox.Focused && TextBox.Visible)
                {
                    listBox.Visible = false;
                    EditCancelled?.Invoke();
                }
            }));
        }
    }

    private void ListBox_LostFocus(object sender, EventArgs e)
    {
        if (parentControl != null && !parentControl.IsDisposed)
        {
            int currentActivationId = activationId; // Capture current activation
            parentControl.BeginInvoke(new Action(() =>
            {
                // Only fire EditCancelled if no new activation happened
                if (currentActivationId == activationId &&
                    !TextBox.Focused && !listBox.Focused && TextBox.Visible)
                {
                    listBox.Visible = false;
                    EditCancelled?.Invoke();
                }
            }));
        }
    }

    private void TextBox_PreviewKeyDown(object sender, PreviewKeyDownEventArgs e)
    {
        if (e.KeyCode == Keys.Tab)
        {
            e.IsInputKey = true;
        }
    }

    private void TextBox_TextChanged(object sender, EventArgs e)
    {
        if (isUpdatingText) return;
        ShowSuggestions(TextBox.Text);
    }

    private void ShowSuggestions(string filterText)
    {
        List<string> filteredSuggestions;

        if (string.IsNullOrEmpty(filterText))
        {
            // Show all suggestions when text is empty
            filteredSuggestions = allSuggestions;
        }
        else
        {
            // Filter suggestions based on input
            filteredSuggestions = allSuggestions
                .Where(s => s.StartsWith(filterText, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        if (filteredSuggestions.Any())
        {
            listBox.BeginUpdate();
            listBox.Items.Clear();
            listBox.Items.AddRange(filteredSuggestions.ToArray());
            listBox.EndUpdate();
            listBox.Visible = true;
            listBox.BringToFront();
        }
        else
        {
            listBox.Visible = false;
        }
    }

    private void TextBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Escape)
        {
            e.SuppressKeyPress = true;
            listBox.Visible = false;
            EditCancelled?.Invoke();
            return;
        }

        if (!listBox.Visible || listBox.Items.Count == 0) return;

        if (e.KeyCode == Keys.Down)
        {
            e.SuppressKeyPress = true; // Prevent cursor from moving
            listBox.Focus();
            listBox.SelectedIndex = 0;
        }
        else if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
        {
            e.SuppressKeyPress = true;
            // If no item is selected, default to the first item.
            if (listBox.SelectedItem == null)
            {
                listBox.SelectedIndex = 0;
            }
            CommitSelection(e.KeyCode);
        }
    }

    private void ListBox_KeyDown(object sender, KeyEventArgs e)
    {
        if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.Tab)
        {
            e.SuppressKeyPress = true;
            CommitSelection(e.KeyCode);
        }
        else if (e.KeyCode == Keys.Escape)
        {
            e.SuppressKeyPress = true;
            listBox.Visible = false;
            TextBox.Focus();
            TextBox.SelectionStart = TextBox.Text.Length;
        }
    }

    private void ListBox_DoubleClick(object sender, EventArgs e)
    {
        CommitSelection(Keys.Enter); // Treat double-click like an Enter press
    }

    private void CommitSelection(Keys commitKey)
    {
        if (listBox.SelectedItem == null) return;

        isUpdatingText = true; // Prevent TextChanged event from re-filtering
        string selectedValue = listBox.SelectedItem.ToString();
        TextBox.Text = selectedValue;
        isUpdatingText = false;

        listBox.Visible = false;
        TextBox.Focus();
        TextBox.SelectionStart = TextBox.Text.Length; // Move cursor to end

        // Raise the event to notify the caller
        ItemSelected?.Invoke(selectedValue, commitKey);
    }

    public bool Visible
    {
        get { return TextBox.Visible; }
        set
        {
            TextBox.Visible = value;
            // The ListBox's visibility is managed by other logic,
            // but we ensure it's hidden when the whole control is hidden.
            if (!value)
            {
                listBox.Visible = false;
            }
        }
    }

    public void UpdateSuggestions(List<string> newSuggestions)
    {
        this.allSuggestions = newSuggestions;
    }

    public void Dispose()
    {
        TextBox.TextChanged -= TextBox_TextChanged;
        TextBox.KeyDown -= TextBox_KeyDown;
        TextBox.PreviewKeyDown -= TextBox_PreviewKeyDown;
        TextBox.LostFocus -= TextBox_LostFocus;
        TextBox.MouseDown -= TextBox_MouseDown;
        listBox.KeyDown -= ListBox_KeyDown;
        listBox.DoubleClick -= ListBox_DoubleClick;
        listBox.LostFocus -= ListBox_LostFocus;

        TextBox.Parent?.Controls.Remove(TextBox);
        listBox.Parent?.Controls.Remove(listBox);
        TextBox.Dispose();
        listBox.Dispose();
    }
}
