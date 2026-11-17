using Aucotec.EngineeringBase.Interop;
using EngineeringBase.Common.Models;

namespace EngineeringBase.Common;

/// <summary>
/// 하나의 EB Application 연결을 유지하면서 프로젝트와 객체 계층을 반복 탐색합니다.
/// 반환된 COM 객체는 사용자가 ReleaseComObject로 해제해야 합니다.
/// </summary>
public interface IEngineeringBaseClient : IDisposable
{
    bool IsRunning();
    bool IsConnected { get; }
    IAucApplication Application { get; }

    void Connect();
    void Disconnect();

    IAucProject GetActiveProject();
    IAucProject FindProject(string nameOrId);
    EbProjectInfo GetProjectInfo(IAucProject project);

    IAucObject GetEquipmentFolder(IAucProject project);
    IAucObject GetDocumentsFolder(IAucProject project);
    IAucObject GetDrawingsFolder(IAucProject project);
    IAucObject GetTemplatesFolder(IAucProject project);
    IAucObject GetFunctionsFolder(IAucProject project);
    IAucObject GetProcessesFolder(IAucProject project);
    IAucObject GetAdvancedDataTrackingFolder(IAucProject project);
    IAucObject GetProjectsRoot();
    IAucObject GetGlobalTemplatesRoot();
    IAucObject GetStencilsRoot();

    IAucObject GetObjectById(string objectId);
    IAucObject? FindObject(IAucObject parent, string name, bool recursive = true);
    IAucObject? FindFolderPath(IAucObject root, string relativePath);
    IReadOnlyList<IAucObject> GetChildren(IAucObject parent);
    IAucObject CopyObject(IAucObject source, IAucObject targetParent, bool includeChildren = true);

    string GetAttributeValue(IAucObject item, int attributeId);
    void SetAttributeValue(IAucObject item, int attributeId, string value);
    void SaveObject(IAucObject item);

    void ReleaseComObject(object? instance);

    EbDrawingResult DrawOnActiveSheet(
        IAucProject project,
        IReadOnlyList<EbSymbolPlacement> placements,
        Action<int, int, string>? progress = null);
}
