namespace Auto_SLD.Models;

internal sealed class ProgressInfo
{
    public ProgressInfo(string caption, string description)
    {
        Caption = caption;
        Description = description;
    }

    public string Caption { get; }
    public string Description { get; }
}
