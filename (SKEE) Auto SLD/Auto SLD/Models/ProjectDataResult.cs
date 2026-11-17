namespace Auto_SLD.Models;

internal sealed class ProjectDataResult
{
    public ProjectDataResult(string projectId, string projectName, IReadOnlyList<LoadDeviceDto> loadDevices)
    {
        ProjectId = projectId;
        ProjectName = projectName;
        LoadDevices = loadDevices;
    }

    public string ProjectId { get; }
    public string ProjectName { get; }
    public IReadOnlyList<LoadDeviceDto> LoadDevices { get; }
}
