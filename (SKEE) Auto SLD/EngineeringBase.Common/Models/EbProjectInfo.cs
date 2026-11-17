namespace EngineeringBase.Common.Models;

public sealed class EbProjectInfo
{
    public EbProjectInfo(string id, string name)
    {
        Id = id;
        Name = name;
    }

    public string Id { get; }
    public string Name { get; }
}
