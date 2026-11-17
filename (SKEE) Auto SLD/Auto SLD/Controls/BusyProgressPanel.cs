using DevExpress.XtraEditors;
using DevExpress.XtraEditors.Controls;
using DevExpress.XtraWaitForm;

namespace Auto_SLD.Controls;

public sealed class BusyProgressPanel : PanelControl
{
    private readonly ProgressPanel _progressPanel;
    private Control? _layoutParent;

    public BusyProgressPanel()
    {
        Appearance.BackColor = Color.White;
        Appearance.BorderColor = Color.Silver;
        Appearance.Options.UseBackColor = true;
        Appearance.Options.UseBorderColor = true;
        BorderStyle = BorderStyles.Simple;
        MinimumSize = new Size(280, 66);
        Padding = new Padding(12, 0, 12, 0);
        Visible = false;

        _progressPanel = new ProgressPanel
        {
            Dock = DockStyle.Fill,
            ContentAlignment = ContentAlignment.MiddleCenter
        };
        _progressPanel.LookAndFeel.UseDefaultLookAndFeel = false;
        _progressPanel.LookAndFeel.Style = DevExpress.LookAndFeel.LookAndFeelStyle.Flat;
        _progressPanel.Appearance.BackColor = Color.White;
        _progressPanel.Appearance.Options.UseBackColor = true;
        Controls.Add(_progressPanel);
    }

    public string Caption
    {
        get => _progressPanel.Caption;
        set => _progressPanel.Caption = value;
    }

    public string Description
    {
        get => _progressPanel.Description;
        set => _progressPanel.Description = value;
    }

    public ContentAlignment ContentAlignment
    {
        get => _progressPanel.ContentAlignment;
        set => _progressPanel.ContentAlignment = value;
    }

    public void ShowProgress(string caption, string description)
    {
        Caption = caption;
        Description = description;
        ResizeAndCenter();
        Visible = true;
        BringToFront();
    }

    public void UpdateProgress(string description, string? caption = null)
    {
        if (caption != null)
            Caption = caption;

        Description = description;
        ResizeAndCenter();
    }

    public void HideProgress()
    {
        Visible = false;
    }

    protected override void OnParentChanged(EventArgs e)
    {
        if (_layoutParent != null)
            _layoutParent.SizeChanged -= LayoutParent_SizeChanged;

        base.OnParentChanged(e);
        _layoutParent = Parent;

        if (_layoutParent != null)
            _layoutParent.SizeChanged += LayoutParent_SizeChanged;

        ResizeAndCenter();
    }

    protected override void Dispose(bool disposing)
    {
        if (disposing && _layoutParent != null)
            _layoutParent.SizeChanged -= LayoutParent_SizeChanged;

        base.Dispose(disposing);
    }

    private void LayoutParent_SizeChanged(object? sender, EventArgs e)
    {
        ResizeAndCenter();
    }

    private void ResizeAndCenter()
    {
        if (Parent == null)
            return;

        Size bestSize = _progressPanel.CalcBestSize();
        Size = new Size(Math.Max(MinimumSize.Width, bestSize.Width + Padding.Horizontal + 24), Math.Max(MinimumSize.Height, bestSize.Height + Padding.Vertical));
        Location = new Point((Parent.ClientSize.Width - Width) / 2, (Parent.ClientSize.Height - Height) / 2);
    }
}
