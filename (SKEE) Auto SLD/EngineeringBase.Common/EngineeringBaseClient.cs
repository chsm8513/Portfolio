using System.Diagnostics;
using System.Globalization;
using Aucotec.EngineeringBase.Interop;
using EngineeringBase.Common.Models;

namespace EngineeringBase.Common;

/// <summary>
/// EB Application COM 연결만 장기간 유지합니다.
/// Project 및 Folder는 호출 시 전달받아 탐색하므로 특정 프로젝트에 고정되지 않습니다.
/// </summary>
public sealed class EngineeringBaseClient : IEngineeringBaseClient
{
    public const string DefaultProgId = "EngineeringBase.Application.31";
    private const string DrawingMacroName = "ExPlugin_csm.ExPlugin_csm";

    private static readonly string[] CircuitComponentFolderNames =
    {
        "Circuit Components",
        "Circuit Componets"
    };

    private readonly string _progId;
    private IAucApplication? _application;

    public EngineeringBaseClient(string progId = DefaultProgId)
    {
        if (string.IsNullOrWhiteSpace(progId))
            throw new ArgumentException("Engineering Base COM ProgID가 비어 있습니다.", nameof(progId));

        _progId = progId.Trim();
    }

    public bool IsConnected => _application is not null;

    public IAucApplication Application => _application
        ?? throw new InvalidOperationException("먼저 Engineering Base에 연결하세요.");

    public bool IsRunning()
    {
        Process[] processes = Process.GetProcessesByName("EngineeringBase");
        try
        {
            return processes.Length > 0;
        }
        finally
        {
            foreach (Process process in processes) process.Dispose();
        }
    }

    public void Connect()
    {
        if (_application is not null) return;
        if (!IsRunning())
            throw new InvalidOperationException("Engineering Base를 실행하세요.");

        Type? applicationType = Type.GetTypeFromProgID(_progId);
        if (applicationType is null)
            throw new InvalidOperationException(
                $"Engineering Base COM ProgID '{_progId}'가 등록되어 있지 않습니다.");

        object? instance = null;
        try
        {
            instance = Activator.CreateInstance(applicationType);
            _application = instance as IAucApplication
                ?? throw new InvalidOperationException("Engineering Base Application 객체를 만들지 못했습니다.");
        }
        catch
        {
            EbCom.Release(instance);
            throw;
        }
    }

    public void Disconnect()
    {
        EbCom.Release(_application);
        _application = null;
    }

    public IAucProject GetActiveProject()
    {
        EnsureConnected();

        IAucObjects? selection = null;
        IAucObject? selectedObject = null;
        try
        {
            selection = Application.Selection;
            if (selection is not null && selection.Count > 0)
            {
                selectedObject = selection.Item(1);
                if (selectedObject is IAucProject selectedProject)
                {
                    selectedObject = null;
                    return selectedProject;
                }

                IAucProject? containingProject = selectedObject?.Project;
                if (containingProject is not null) return containingProject;
            }

            return Application.ActiveProject
                ?? throw new InvalidOperationException(
                    "Engineering Base에 선택되었거나 활성화된 프로젝트가 없습니다.");
        }
        finally
        {
            EbCom.Release(selectedObject);
            EbCom.Release(selection);
        }
    }

    public IAucProject FindProject(string nameOrId)
    {
        EnsureConnected();
        if (string.IsNullOrWhiteSpace(nameOrId))
            throw new ArgumentException("찾을 프로젝트 이름 또는 ID가 비어 있습니다.", nameof(nameOrId));

        string value = nameOrId.Trim();
        if (Guid.TryParse(value, out _))
        {
            IAucObject byId = GetObjectById(value);
            if (byId is IAucProject projectById) return projectById;

            EbCom.Release(byId);
            throw new InvalidOperationException($"ID '{value}'는 Project 객체가 아닙니다.");
        }

        IAucVbaAppFolders? folders = null;
        IAucObject? projectsRoot = null;
        IAucObject? found = null;
        try
        {
            folders = Application.Folders;
            projectsRoot = folders.Projects;
            found = FindObject(projectsRoot, value, false);
            if (found is IAucProject project)
            {
                found = null;
                return project;
            }

            throw new InvalidOperationException($"Engineering Base에서 프로젝트 '{value}'을(를) 찾을 수 없습니다.");
        }
        finally
        {
            EbCom.Release(found);
            EbCom.Release(projectsRoot);
            EbCom.Release(folders);
        }
    }

    public EbProjectInfo GetProjectInfo(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return new EbProjectInfo(project.ID ?? string.Empty, project.Name ?? string.Empty);
    }

    public IAucObject GetEquipmentFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.EquipmentFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Equipment Folder가 없습니다.");
    }

    public IAucObject GetDrawingsFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.DrawingsFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Drawings Folder가 없습니다.");
    }

    /// <summary>
    /// EB API에서 Documents 루트는 DrawingsFolder라는 이름으로 제공됩니다.
    /// </summary>
    public IAucObject GetDocumentsFolder(IAucProject project) => GetDrawingsFolder(project);

    public IAucObject GetTemplatesFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.TemplatesFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Templates Folder가 없습니다.");
    }

    public IAucObject GetFunctionsFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.FunctionsFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Functions Folder가 없습니다.");
    }

    public IAucObject GetProcessesFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.ProcessesFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Processes Folder가 없습니다.");
    }

    public IAucObject GetAdvancedDataTrackingFolder(IAucProject project)
    {
        EnsureNotNull(project, nameof(project));
        return project.AdvancedDataTrackingFolder
            ?? throw new InvalidOperationException($"프로젝트 '{project.Name}'의 Data Tracking Folder가 없습니다.");
    }

    public IAucObject GetProjectsRoot() => GetApplicationRoot(folders => folders.Projects, "Projects");

    public IAucObject GetGlobalTemplatesRoot() => GetApplicationRoot(folders => folders.Templates, "Templates");

    public IAucObject GetStencilsRoot()
        => GetApplicationRoot(folders => folders.Stencils, "Stencils");

    private IAucObject GetApplicationRoot(
        Func<IAucVbaAppFolders, IAucObject> selector,
        string rootName)
    {
        EnsureConnected();
        IAucVbaAppFolders? folders = null;
        try
        {
            folders = Application.Folders;
            return selector(folders)
                ?? throw new InvalidOperationException($"Engineering Base {rootName} 루트가 없습니다.");
        }
        finally
        {
            EbCom.Release(folders);
        }
    }

    public IAucObject GetObjectById(string objectId)
    {
        EnsureConnected();
        if (string.IsNullOrWhiteSpace(objectId))
            throw new ArgumentException("EB Object ID가 비어 있습니다.", nameof(objectId));

        IAucVbaUtils? utils = null;
        try
        {
            utils = Application.Utils;
            return utils.GetSnglObjectByID(objectId.Trim())
                ?? throw new InvalidOperationException($"EB Object ID '{objectId}'를 찾을 수 없습니다.");
        }
        finally
        {
            EbCom.Release(utils);
        }
    }

    public IAucObject? FindObject(IAucObject parent, string name, bool recursive = true)
    {
        EnsureNotNull(parent, nameof(parent));
        if (string.IsNullOrWhiteSpace(name))
            throw new ArgumentException("찾을 객체 이름이 비어 있습니다.", nameof(name));

        return FindFirstByName(parent, new[] { name.Trim() }, recursive ? 32 : 0);
    }

    public IAucObject? FindFolderPath(IAucObject root, string relativePath)
    {
        EnsureNotNull(root, nameof(root));
        if (string.IsNullOrWhiteSpace(relativePath)) return root;

        string[] parts = relativePath.Split(
                new[] { '/', '\\' }, StringSplitOptions.RemoveEmptyEntries)
            .Select(part => part.Trim())
            .ToArray();

        IAucObject current = root;
        bool ownsCurrent = false;
        foreach (string part in parts)
        {
            IAucObject? next = FindObject(current, part, false);
            if (ownsCurrent) EbCom.Release(current);
            if (next is null) return null;

            current = next;
            ownsCurrent = true;
        }

        return current;
    }

    /// <summary>
    /// 반환된 각 COM 객체는 호출자가 ReleaseComObject로 해제해야 합니다.
    /// </summary>
    public IReadOnlyList<IAucObject> GetChildren(IAucObject parent)
    {
        EnsureNotNull(parent, nameof(parent));
        IAucObjects? children = null;
        var result = new List<IAucObject>();
        try
        {
            children = parent.Children;
            for (int index = 1; index <= children.Count; index++)
            {
                IAucObject? child = children.Item(index);
                if (child is not null) result.Add(child);
            }

            return result;
        }
        catch
        {
            foreach (IAucObject item in result) EbCom.Release(item);
            throw;
        }
        finally
        {
            EbCom.Release(children);
        }
    }

    public IAucObject CopyObject(
        IAucObject source,
        IAucObject targetParent,
        bool includeChildren = true)
    {
        EnsureNotNull(source, nameof(source));
        EnsureNotNull(targetParent, nameof(targetParent));
        return source.CopyTo(targetParent, includeChildren)
            ?? throw new InvalidOperationException(
                $"'{source.Name}' 객체를 '{targetParent.Name}' 아래로 복사하지 못했습니다.");
    }

    public string GetAttributeValue(IAucObject item, int attributeId)
    {
        IAucAttribute attribute = GetAttribute(item, attributeId);
        try
        {
            return Convert.ToString(attribute.Value, CultureInfo.CurrentCulture) ?? string.Empty;
        }
        finally
        {
            EbCom.Release(attribute);
        }
    }

    public void SetAttributeValue(IAucObject item, int attributeId, string value)
    {
        IAucAttribute attribute = GetAttribute(item, attributeId);
        try
        {
            attribute.Value = value ?? string.Empty;
        }
        finally
        {
            EbCom.Release(attribute);
        }
    }

    public void SaveObject(IAucObject item)
    {
        EnsureNotNull(item, nameof(item));
        item.Store();
    }

    public void ReleaseComObject(object? instance) => EbCom.Release(instance);

    public EbDrawingResult DrawOnActiveSheet(
        IAucProject project,
        IReadOnlyList<EbSymbolPlacement> placements,
        Action<int, int, string>? progress = null)
    {
        EnsureConnected();
        EnsureNotNull(project, nameof(project));
        EnsureNotNull(placements, nameof(placements));
        if (placements.Count == 0)
            throw new InvalidOperationException("배치할 EB 객체 데이터가 없습니다.");

        IAucSheet? sheet = null;
        IAucObject? stencilsRoot = null;
        IAucObject? circuitComponents = null;
        var errors = new List<string>();
        int placedCount = 0;

        try
        {
            sheet = Application.ActiveSheet;
            if (sheet is null)
                throw new InvalidOperationException("심볼을 배치할 EB 도면을 먼저 열어 활성화하세요.");

            stencilsRoot = GetStencilsRoot();
            circuitComponents = FindFirstByName(stencilsRoot, CircuitComponentFolderNames, 12);
            if (circuitComponents is null)
                throw new InvalidOperationException(
                    "EB Stencils에서 'Circuit Components' 폴더를 찾을 수 없습니다.");

            for (int index = 0; index < placements.Count; index++)
            {
                EbSymbolPlacement placement = placements[index];
                progress?.Invoke(index + 1, placements.Count, placement.SymbolName);
                IAucObject? symbol = null;
                try
                {
                    symbol = FindObject(circuitComponents, placement.SymbolName);
                    if (symbol is null)
                    {
                        errors.Add($"입력 {placement.SourceIndex}: Stencil 객체 '{placement.SymbolName}'을(를) 찾지 못했습니다.");
                        continue;
                    }

                    DropSymbol(Application, project, sheet, symbol, placement.X, placement.Y);
                    placedCount++;
                }
                catch (Exception exception)
                {
                    errors.Add($"입력 {placement.SourceIndex} '{placement.SymbolName}': {exception.Message}");
                }
                finally
                {
                    EbCom.Release(symbol);
                }
            }

            if (placedCount > 0) sheet.Store();
            return new EbDrawingResult(placements.Count, placedCount, errors);
        }
        finally
        {
            EbCom.Release(circuitComponents);
            EbCom.Release(stencilsRoot);
            EbCom.Release(sheet);
        }
    }

    public void Dispose() => Disconnect();

    private void EnsureConnected()
    {
        if (_application is null) Connect();
    }

    private static IAucAttribute GetAttribute(IAucObject item, int attributeId)
    {
        EnsureNotNull(item, nameof(item));
        if (attributeId <= 0) throw new ArgumentOutOfRangeException(nameof(attributeId));

        IAucAttributes? attributes = null;
        try
        {
            attributes = item.Attributes;
            return attributes.ItemByID((AucAttribute)attributeId)
                ?? throw new InvalidOperationException(
                    $"'{item.Name}' 객체에서 Attribute ID {attributeId}을(를) 찾을 수 없습니다.");
        }
        finally
        {
            EbCom.Release(attributes);
        }
    }

    private static IAucObject? FindFirstByName(
        IAucObject parent,
        IReadOnlyCollection<string> names,
        int remainingDepth)
    {
        IAucObjects? children = null;
        try
        {
            children = parent.Children;
            for (int index = 1; index <= children.Count; index++)
            {
                IAucObject? child = null;
                bool keepChild = false;
                try
                {
                    child = children.Item(index);
                    if (child is null) continue;

                    if (names.Any(name => string.Equals(
                        child.Name, name, StringComparison.CurrentCultureIgnoreCase)))
                    {
                        keepChild = true;
                        return child;
                    }

                    if (remainingDepth > 0)
                    {
                        IAucObject? descendant = FindFirstByName(child, names, remainingDepth - 1);
                        if (descendant is not null) return descendant;
                    }
                }
                catch
                {
                    // Children이 없는 EB 객체는 건너뛴다.
                }
                finally
                {
                    if (!keepChild) EbCom.Release(child);
                }
            }

            return null;
        }
        finally
        {
            EbCom.Release(children);
        }
    }

    private static void EnsureNotNull(object? value, string parameterName)
    {
        if (value is null) throw new ArgumentNullException(parameterName);
    }

    private static void DropSymbol(
        IAucApplication application,
        IAucProject project,
        IAucSheet sheet,
        IAucObject symbol,
        double x,
        double y)
    {
        application.Automation.Parameters.Clear();
        application.Automation.Parameters.Add("sbDropSymbol");
        application.Automation.Parameters.Add(sheet);
        application.Automation.Parameters.Add(symbol);
        application.Automation.Parameters.Add(x);
        application.Automation.Parameters.Add(y);
        application.Utils.RunMacro(DrawingMacroName, project);
    }
}
