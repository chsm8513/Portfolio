using Aucotec.EngineeringBase.Interop;
using EbApp = Aucotec.EngineeringBase.Interop.Application;
using Auto_SLD.Models;
using EngineeringBase;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using Visio = Microsoft.Office.Interop.Visio;

namespace Auto_SLD.Service;

// 묶음별 원본 Copy -> 목적지 Paste -> Special Paste -> association 교체 -> Unit 이동.
internal sealed class AutoDrawingService
{
    private const int MaxOpenTypicalSheets = 3;
    private readonly IEbAccessService _access;
    private readonly IEbRepository _repository;
    private readonly Queue<TypicalDrawing> _openTypicalDrawings = new Queue<TypicalDrawing>();
    private Visio.Application? _visio;
    private HashSet<string> _originalEquipmentIds = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
    // device 하나를 받는 association 메서드에 현재 붙여넣기 묶음만 제공한다.
    private IReadOnlyList<IAucObject> _pastedEquipment = Array.Empty<IAucObject>();
    private bool _pastingIncomer;

    public AutoDrawingService(IEbAccessService access, IEbRepository repository) { _access = access; _repository = repository; }

    public AutoDrawingResult New_draw(string projectId, IReadOnlyList<LoadDeviceDto> devices, IProgress<ProgressInfo> progress, DrawingLayoutOptions? options = null)
    {
        options ??= new DrawingLayoutOptions();
        var result = new AutoDrawingResult();
        var drawings = new Dictionary<string, TypicalDrawing>(StringComparer.OrdinalIgnoreCase);
        IAucProject project = (IAucProject)_repository.GetObjectById(projectId);
        IAucObject? equipment = null, unitParent = null, documents = null, destination = null, templates = null, favorites = null, typicalDocuments = null;
        IAucSheet? sheet = null;
        IAucObject? unit = null;
        string stage = "프로젝트 경로 조회", sheetId = "생성 전";
        int currentDeviceIndex = -1;
        bool completed = false;
        try
        {
            IAucObject parent = _access.GetApplication().Folders.ProjectTemplates;          
            try
            {
                IReadOnlyList<IAucObject> matches = _repository.FindObjects(parent, new EbObjectFilter(null, null, null), false);
                try
                {
                    IAucProject typical = matches.OfType<IAucProject>().FirstOrDefault(item => string.Equals(item.Name.ToLower(), "sld typical project", StringComparison.OrdinalIgnoreCase)) ?? throw new InvalidOperationException("현재 프로젝트 Template 아래에서 'SLD Typical Project'를 찾지 못했습니다.");
                    typicalDocuments = _repository.GetDocumentsRoot(typical);
                }
                finally { foreach (IAucObject item in matches) Release(item); }
            }
            finally { Release(parent); }
            equipment = _repository.GetEquipmentRoot(project);
            _originalEquipmentIds = ReadEquipmentIds(equipment);
            unitParent = _repository.FindFolderPath(equipment, @"Electrical & Instrumentation\SLD") ?? throw new InvalidOperationException("Equipment\\Electrical & Instrumentation\\SLD 경로가 없습니다.");
            documents = _repository.GetDocumentsRoot(project);
            destination = _repository.FindFolderPath(documents, @"Electrical\Single Line Diagram") ?? throw new InvalidOperationException("Documents\\Electrical\\Single Line Diagram 경로가 없습니다.");
            templates = project.SheetTemplatesFolder;
            favorites = _repository.FindFolderPath(templates, "Favorites") ?? throw new InvalidOperationException("Templates\\Sheets\\Favorites 경로가 없습니다.");
            IAucSheet baseSheet = FindSheet(favorites, "A*L_SLD");
            string baseSheetId;
            try { baseSheetId = baseSheet.ID; }
            finally { Release(baseSheet); }

            TypicalDrawing GetDrawing(string name)
            {
                if (drawings.TryGetValue(name, out TypicalDrawing drawing)) return drawing;
                IAucSheet source = FindSheet(typicalDocuments, name, true);
                drawing = new TypicalDrawing { Sheet = source, WasOpen = source.IsOpened };
                drawings.Add(name, drawing);
                return drawing;
            }

            void FinishSheet()
            {
                stage = "CurrentSheet 저장·닫기";
                progress.Report(new ProgressInfo(stage, sheetId));
                sheet!.Store();
                sheet.Close();
                Release(sheet, unit);
                sheet = null;
                unit = null;
                result.SheetCount++;
            }

            double ReadPageWidth()
            {
                Visio.Page page = OpenPage(sheet!);
                Visio.Shape pageShape = page.PageSheet;
                Visio.Cell widthCell = pageShape.CellsU["PageWidth"];
                try { return widthCell.Result[Visio.VisUnitCodes.visMillimeters]; }
                finally { Release(widthCell, pageShape, page); }
            }

            LayoutCursor CreateLayout(TypicalDrawing incomer, double left, double top, DrawingGroupLayoutOptions groupLayout)
            {
                return new LayoutCursor(left, top, incomer.Width, incomer.Height, ReadPageWidth(), options, groupLayout);
            }

            double ReadPageHeight()
            {
                Visio.Page page = OpenPage(sheet!);
                Visio.Shape pageShape = page.PageSheet;
                Visio.Cell heightCell = pageShape.CellsU["PageHeight"];
                try { return heightCell.Result[Visio.VisUnitCodes.visMillimeters]; }
                finally { Release(heightCell, pageShape, page); }
            }

            void CreateDestinationSheet(string groupSheetName)
            {
                stage = "Sheet / Unit 생성";
                string name = UniqueSheetName(destination, groupSheetName);
                progress.Report(new ProgressInfo(stage, name));
                unit = GetOrCreateUnit(unitParent, name);
                IAucSheet template = (IAucSheet)_repository.GetObjectById(baseSheetId);
                try { sheet = (IAucSheet)template.CopyTo2(destination, false, false, true); }
                finally { Release(template); }
                sheetId = sheet.ID;
                SetDesignation((IAucObject)sheet, name);
            }

            LayoutCursor StartSheet(TypicalDrawing incomer, LoadDeviceDto device, string groupSheetName, DrawingGroupLayoutOptions groupLayout)
            {
                CreateDestinationSheet(groupSheetName);
                stage = "Incomer 복사/붙여넣기";
                progress.Report(new ProgressInfo(stage, incomer.Sheet.Name));
                CopyDrawing(incomer);                
                double left = options.LeftMargin, top = ReadPageHeight() - options.TopMargin;
                PasteDrawing(sheet!, equipment, unit!, incomer, left, top, device, result, true);
                return CreateLayout(incomer, left, top, groupLayout);
            }

            LayoutCursor StartContinuationSheet(string groupSheetName, DrawingGroupLayoutOptions groupLayout)
            {
                CreateDestinationSheet(groupSheetName);
                stage = "Feeder 연속 Sheet 배치";
                progress.Report(new ProgressInfo(stage, sheet!.Name));
                return LayoutCursor.CreateContinuation(ReadPageWidth(), ReadPageHeight(), options, groupLayout);
            }

            // 해당 부분만 고치면 Sheet이름이 변경됨
            string SLD_Name = "SLD";
            int maxGroupNumber = 0;
            var namePattern = new Regex("^" + Regex.Escape(SLD_Name) + @"(?:_(\d+))?(?:\(\d+\))?$", RegexOptions.IgnoreCase);
            IAucObjects existingSheets = destination.FindObjects(AucObjectKind.aucObjSheet, AucVbSearchType.aucSearchFlat);
            try
            {
                for (int index = 1; index <= existingSheets.Count; index++)
                {
                    IAucObject existingSheet = existingSheets.Item(index);
                    try
                    {
                        Match match = namePattern.Match(existingSheet.Name);
                        if (!match.Success) continue;

                        int number = 1; // SLD 또는 SLD(1)은 그룹 번호 1
                        if (match.Groups[1].Success && !int.TryParse(match.Groups[1].Value, out number)) continue;

                        maxGroupNumber = Math.Max(maxGroupNumber, number);
                    }
                    finally { Release(existingSheet); }
                }
            }
            finally { Release(existingSheets); }
            int groupNumber = maxGroupNumber + 1;

            // Load DTO의 SwitchGear가 Incomer이고 각 Load가 하단에 반복 배치할 Feeder다.
            foreach (var group in devices.Select((device, index) => new { Device = device, Index = index }).GroupBy(item => string.IsNullOrWhiteSpace(item.Device.SwitchGearId) ? item.Device.SwitchGear : item.Device.SwitchGearId, StringComparer.OrdinalIgnoreCase))
            {
                var groupItems = group.ToList();
                LoadDeviceDto incomerDevice = groupItems.Select(item => item.Device).FirstOrDefault(item => !string.IsNullOrWhiteSpace(item.SwitchGearType));
                if (incomerDevice == null) continue;

                string groupSheetName = groupNumber == 1 ? $"{SLD_Name}" : $"{SLD_Name}_{groupNumber}";
                TypicalDrawing incomer = GetDrawing(incomerDevice.SwitchGearType);
                DrawingGroupLayoutOptions groupLayout = options.ResolveGroupLayout(incomerDevice.SwitchGearType, groupItems.Select(item => item.Device.FeederType));
                groupItems = groupItems.OrderBy(item => groupLayout.IsLastFeederType(item.Device.FeederType) ? 1 : 0).ThenBy(item => item.Index).ToList();
                currentDeviceIndex = groupItems[0].Index;
                Debug.WriteLine($"[배치 설정] {groupLayout.Name} / Incomer={incomerDevice.SwitchGearType} / Feeder={string.Join(",", groupItems.Select(item => item.Device.FeederType).Distinct(StringComparer.OrdinalIgnoreCase))} / Mode={groupLayout.PlacementMode}");
                LayoutCursor layout = StartSheet(incomer, incomerDevice, groupSheetName, groupLayout);
                foreach (var pendingItem in groupItems)
                {
                    LoadDeviceDto device = pendingItem.Device;
                    currentDeviceIndex = pendingItem.Index;
                    stage = "Feeder 복사/붙여넣기";
                    progress.Report(new ProgressInfo(stage, $"{device.FeederType} / {device.Designation} / Load {result.FeederCount + 1} / {devices.Count}"));
                    TypicalDrawing feeder = GetDrawing(device.FeederType);
                    CopyDrawing(feeder);
                    if (!layout.TryPlace(feeder.Width, feeder.Height, out double left, out double top))
                    {
                        FinishSheet();
                        layout = StartContinuationSheet(groupSheetName, groupLayout);
                        if (!layout.TryPlace(feeder.Width, feeder.Height, out left, out top)) throw new InvalidOperationException($"Feeder '{device.FeederType}'를 빈 연속 Sheet에도 배치할 수 없습니다.");
                    }
                    stage = "Feeder 붙여넣기·Load 연결·Unit 이동";
                    progress.Report(new ProgressInfo(stage, device.Designation));
                    PasteDrawing(sheet!, equipment, unit!, feeder, left, top, device, result);
                    result.FeederCount++;
                }
                FinishSheet();
            }
            completed = true;
            return result;
        }
        catch (Exception exception)
        {
            var drawingException = new InvalidOperationException($"Auto Drawing 실패 [{stage}], Sheet ID={sheetId}. 작성 중인 Sheet와 Unit은 남겨두었습니다.\n{exception.Message}", exception);
            if (currentDeviceIndex >= 0) drawingException.Data["FailedDeviceIndex"] = currentDeviceIndex;
            throw drawingException;
        }
        finally
        {
            foreach (TypicalDrawing drawing in drawings.Values)
            {
                // 실패/대화창 대기 중에는 추가 Close 호출로 EB를 건드리지 않는다.
                if (completed && !drawing.WasOpen && drawing.Sheet.IsOpened) drawing.Sheet.Close();
                Release(drawing.Sheet);
            }
            Release(sheet, unit, favorites, templates, destination, documents, typicalDocuments, unitParent, equipment, project, _visio);
            _visio = null;
        }
    }

    private void CopyDrawing(TypicalDrawing drawing)
    {
        EnsureTypicalSheetOpen(drawing);
        Visio.Page page = OpenPage(drawing.Sheet);
        Visio.Window window = _visio!.ActiveWindow;
        Visio.Selection? selection = null;
        try
        {
            window.SelectAll();
            selection = window.Selection;
            if (!drawing.HasBounds)
            {
                selection.BoundingBox((short)(Visio.VisBoundingBoxArgs.visBBoxUprightWH | Visio.VisBoundingBoxArgs.visBBoxDrawingCoords), out double left, out double bottom, out double right, out double top);
                drawing.Left = left * 25.4; drawing.Top = top * 25.4;
                drawing.Width = (right - left) * 25.4; drawing.Height = (top - bottom) * 25.4;
                drawing.HasBounds = true;
            }
            selection.Copy();
        }
        finally { Release(selection, window, page); }
    }

    private void EnsureTypicalSheetOpen(TypicalDrawing drawing)
    {
        if (drawing.Sheet.IsOpened) return;
        while (_openTypicalDrawings.Count >= MaxOpenTypicalSheets)
        {
            TypicalDrawing oldest = _openTypicalDrawings.Dequeue();
            if (!oldest.Sheet.IsOpened) continue;
            Debug.WriteLine($"[Typical Sheet FIFO] 가장 먼저 연 Sheet 닫기: {oldest.Sheet.Name}");
            oldest.Sheet.Close();
        }
        drawing.Sheet.Open(AucSheetOpenFlags.aucSheetOpenSheet | AucSheetOpenFlags.aucSheetOpenAutosave);
        _openTypicalDrawings.Enqueue(drawing);
        Debug.WriteLine($"[Typical Sheet FIFO] Sheet 열기: {drawing.Sheet.Name} / Auto Drawing이 연 Typical Sheet={_openTypicalDrawings.Count}/{MaxOpenTypicalSheets}");
    }

    private void PasteDrawing(IAucSheet sheet, IAucObject equipment, IAucObject unit, TypicalDrawing drawing, double left, double top, LoadDeviceDto device, AutoDrawingResult result, bool isIncomer = false)
    {
        Visio.Page page = OpenPage(sheet);
        System.Threading.Thread.Sleep(1000);
        double centerX = left + drawing.Width / 2, centerY = top - drawing.Height / 2;
        Debug.WriteLine($"[배치 좌표] {(isIncomer ? "Incomer" : device.FeederType)} / 요청 Left={left:F2}, Top={top:F2}mm / 원본 Left={drawing.Left:F2}, Top={drawing.Top:F2}, Width={drawing.Width:F2}, Height={drawing.Height:F2}mm");
        try
        {
            page.PasteToLocation(centerX / 25.4, centerY / 25.4, 0);
            System.Threading.Thread.Sleep(250);
            DebugPasteBounds("PasteToLocation 직후", page, left, top, centerX, centerY);
            IAucApplication application = _access.GetApplication();
            try { new EbPasteSpecialUiService().Complete(page, centerX, centerY, new IntPtr(application.ActiveWindow)); }
            finally { Release(application); }
            DebugPasteBounds("Paste Special 완료 후", page, left, top, centerX, centerY);
            sheet.Store();
            CompletePastedEquipment(equipment, unit, device, result, isIncomer);
        }
        finally { Release(page); }
    }

    private void CompletePastedEquipment(IAucObject equipment, IAucObject unit, LoadDeviceDto device, AutoDrawingResult result, bool isIncomer)
    {
        IReadOnlyList<IAucObject> copied = ReadPastedEquipment(equipment);
        try
        {
            // 이 한 묶음만 연결하고 이동을 끝낸 다음 호출자(다음 복사)로 돌아간다.
            _pastedEquipment = copied;
            _pastingIncomer = isIncomer;
            Debug.WriteLine($"[붙여넣기 시작] {(isIncomer ? "Incomer" : "Feeder")} / Load={device.Designation} / LoadId={device.LoadId} / SwitchGearId={device.SwitchGearId} / 새 객체={copied.Count}");
            if (_pastingIncomer)
            {
                ReplacePastedObjectAssociations(device, "SwitchGear", result);
            }
            else
            {
                ReplacePastedObjectAssociations(device, "Load", result);
                ReplacePastedObjectAssociations(device, "BUSBAR", result);
            }
            System.Threading.Thread.Sleep(250);
            result.MovedEquipmentCount += MovePastedEquipment(equipment, unit, copied);
            Debug.WriteLine($"[붙여넣기 완료] {(isIncomer ? "Incomer" : "Feeder")} / {device.Designation} / Unit 이동 {copied.Count}개 완료");
        }
        finally { _pastedEquipment = Array.Empty<IAucObject>(); _pastingIncomer = false; foreach (IAucObject item in copied) Release(item); }
    }

    [Conditional("DEBUG")]
    private void DebugPasteBounds(string stage, Visio.Page page, double requestedLeftMm, double requestedTopMm, double requestedCenterXmm, double requestedCenterYmm)
    {
        Visio.Window? window = null;
        Visio.Selection? selection = null;
        Visio.Shape? pageShape = null;
        Visio.Cell? xRulerOriginCell = null, yRulerOriginCell = null;
        try
        {
            window = _visio!.ActiveWindow;
            selection = window.Selection;
            selection.BoundingBox((short)(Visio.VisBoundingBoxArgs.visBBoxUprightWH | Visio.VisBoundingBoxArgs.visBBoxDrawingCoords), out double left, out double bottom, out double right, out double top);
            double actualLeftMm = left * 25.4, actualTopMm = top * 25.4;
            double actualCenterXmm = (left + right) / 2 * 25.4, actualCenterYmm = (bottom + top) / 2 * 25.4;
            pageShape = page.PageSheet;
            xRulerOriginCell = pageShape.CellsU["XRulerOrigin"];
            yRulerOriginCell = pageShape.CellsU["YRulerOrigin"];
            double xRulerOriginMm = xRulerOriginCell.Result[Visio.VisUnitCodes.visMillimeters];
            double yRulerOriginMm = yRulerOriginCell.Result[Visio.VisUnitCodes.visMillimeters];
            Debug.WriteLine($"[배치 실측:{stage}] Selection={selection.Count}개 / 요청 Left={requestedLeftMm:F3}, Top={requestedTopMm:F3}, Center=({requestedCenterXmm:F3}, {requestedCenterYmm:F3})mm / 실제 Left={actualLeftMm:F3}, Top={actualTopMm:F3}, Center=({actualCenterXmm:F3}, {actualCenterYmm:F3})mm / 오차 X={actualLeftMm - requestedLeftMm:+0.000;-0.000;0.000}, Y={actualTopMm - requestedTopMm:+0.000;-0.000;0.000}mm / RulerOrigin=({xRulerOriginMm:F3}, {yRulerOriginMm:F3})mm");
        }
        catch (Exception exception)
        {
            Debug.WriteLine($"[배치 실측:{stage}] BoundingBox 측정 실패: {exception.Message}");
        }
        finally { Release(yRulerOriginCell, xRulerOriginCell, pageShape, selection, window); }
    }

    private IReadOnlyList<IAucObject> ReadPastedEquipment(IAucObject equipment, int timeoutMs = 60000)
    {
        var timer = Stopwatch.StartNew();
        const int stableDurationMs = 1000;
        int previousCount = -1;
        long stableSinceMs = 0;
        while (true)
        {
            equipment.Refresh();
            IAucObjects children = equipment.FindObjects(AucObjectKind.aucObjUnspecified, AucVbSearchType.aucSearchFlat);
            var copied = new List<IAucObject>();
            var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            bool ready = false;
            try
            {
                int count = children.Count;
                for (int index = 1; index <= count; index++)
                {
                    IAucObject item = children.Item(index);
                    string id = item.ID;
                    if (!_originalEquipmentIds.Contains(id)) { copied.Add(item); ids.Add(id); }
                    else Release(item);
                }
                if (ids.Count != previousCount)
                {
                    previousCount = ids.Count;
                    stableSinceMs = timer.ElapsedMilliseconds;
                }
                long stableForMs = timer.ElapsedMilliseconds - stableSinceMs;
                Debug.WriteLine($"[Equipment 생성 대기] 새 객체={ids.Count} / 개수 유지={stableForMs}ms / 경과={timer.ElapsedMilliseconds}ms");
                // ID는 이동 대상을 구분할 때만 사용하고, 생성 완료는 새 객체 수가 일정 시간 증가하지 않는 것으로 판단한다.
                ready = ids.Count > 0 && stableForMs >= stableDurationMs;
                if (ready) return copied;
            }
            finally { Release(children); if (!ready) foreach (IAucObject item in copied) Release(item); }
            if (timer.ElapsedMilliseconds >= timeoutMs) throw new InvalidOperationException($"Special Paste 이후 Equipment root의 새 객체 생성이 {timeoutMs / 1000}초 안에 안정되지 않았습니다(마지막 조회 {previousCount}개). association/MoveTo를 완료로 처리하지 않고 다음 붙여넣기를 중단합니다.");
            System.Threading.Thread.Sleep(250);
        }
    }

    private static int MovePastedEquipment(IAucObject equipment, IAucObject unit, IReadOnlyList<IAucObject> copied)
    {
        if (copied.Count == 0) throw new InvalidOperationException("이동할 새 Equipment 객체가 없습니다. 다음 붙여넣기를 중단합니다.");
        foreach (IAucObject item in copied)
        {
            if (!item.MoveTo(unit)) throw new InvalidOperationException($"'{item.Name}'의 Unit 이동에 실패했습니다.");
            item.Store();
        }
        System.Threading.Thread.Sleep(250);
        HashSet<string> rootIds = ReadEquipmentIds(equipment);
        string[] remaining = copied.Select(item => item.ID).Where(rootIds.Contains).ToArray();
        if (remaining.Length > 0) throw new InvalidOperationException($"Unit 이동 후에도 Equipment root에 {remaining.Length}개가 남아 있습니다. 다음 붙여넣기를 중단합니다. ID={string.Join(", ", remaining)}");
        return copied.Count;
    }

    private void ReplacePastedObjectAssociations(LoadDeviceDto device, string searchName, AutoDrawingResult result)
    {
        string associationType = searchName.Replace(" ", "").Trim().ToLowerInvariant();
        string targetId = associationType switch
        {
            "switchgear" => device.SwitchGearId,
            "load" => device.LoadId,
            "busbar" => device.CableId,
            _ => throw new ArgumentException($"지원하지 않는 MasterShape 검색 이름입니다: {searchName}", nameof(searchName))
        };
        foreach (IAucObject child in _pastedEquipment)
        {
            IAucAssociations? source = null, links = null, targetLinks = null;
            IAucAssociation? association = null, added = null;
            IAucObject? target = null, shape = null;
            try
            {
                if (!MatchesAssociationObject(child, associationType)) continue;
                source = child.SourceAssociations;
                links = source.Filter((AucRole)86);
                if (links.Count == 0)
                    continue;
                association = links.Item(1);
                shape = association.RelatedObject;
                if (shape == null)
                    continue;
                target = _repository.GetObjectById(targetId);
                targetLinks = target.SourceAssociations;
                Debug.WriteLine($"[Association] {searchName}={targetId} / Type={associationType} / Symbol={shape.ID} / 복사 객체={child.ID}");
                if (!source.Delete((AucRole)86, shape)) throw new InvalidOperationException("복사 객체의 심볼 연결 해제에 실패했습니다.");
                result.RemovedDummyCount++;
                child.Store();
                added = targetLinks.New((AucRole)86, shape);
                result.ReassociatedSymbolCount++;
                target.Store();

            }
            finally { Release(added, targetLinks, target, shape, association, links, source); }
        }
    }

    private static bool MatchesAssociationObject(IAucObject child, string associationType)
    {
        switch (associationType)
        {
            case "switchgear":
                child.ExecuteFormula("El2;A328;", out string switchType);
                return !string.IsNullOrWhiteSpace(switchType) && string.Equals(switchType.Replace(" ", "").Trim(), "switchgear", StringComparison.OrdinalIgnoreCase);
            case "load":
                child.ExecuteFormula("El2;A101619;", out string loadType);
                return !string.IsNullOrWhiteSpace(loadType);
            case "busbar":
                child.ExecuteFormula("El2;A105379;", out string masterShapeName);
                return !string.IsNullOrWhiteSpace(masterShapeName) && masterShapeName.Replace(" ", "").Trim().StartsWith("busbar", StringComparison.OrdinalIgnoreCase);
            default:
                throw new ArgumentException($"지원하지 않는 Association 객체 유형입니다: {associationType}", nameof(associationType));
        }
    }

    private Visio.Page OpenPage(IAucSheet sheet)
    {
        if (!sheet.IsOpened) sheet.Open(AucSheetOpenFlags.aucSheetOpenSheet | AucSheetOpenFlags.aucSheetOpenAutosave);
        sheet.Activate();
        _visio ??= (Visio.Application)Marshal.GetActiveObject("Visio.Application");
        return _visio.ActivePage;
    }

    private static IAucSheet FindSheet(IAucObject root, string name, bool excludeCopyBin = false)
    {
        IAucObjects found = root.FindObjects(AucObjectKind.aucObjSheet, AucVbSearchType.aucSearchDeep, AucType.aucTypeUnspecified, AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, name);
        try
        {
            string rootId = excludeCopyBin ? root.ID : string.Empty;
            for (int index = 1; index <= found.Count; index++)
            {
                IAucSheet candidate = (IAucSheet)found.Item(index);
                bool inCopyBin = false;
                try
                {
                    // 전체 트리를 읽지 않고 이름이 일치한 Sheet의 경로만 확인한다.
                    IAucObject? parent = excludeCopyBin ? candidate.Parent : null;
                    try
                    {
                        while (parent != null && !string.Equals(parent.ID, rootId, StringComparison.OrdinalIgnoreCase))
                        {
                            if (string.Equals(parent.Name, "Copy bin", StringComparison.OrdinalIgnoreCase)) { inCopyBin = true; break; }
                            IAucObject previous = parent;
                            parent = parent.Parent;
                            Release(previous);
                        }
                    }
                    finally { Release(parent); }
                    if (!inCopyBin) return candidate;
                }
                catch { Release(candidate); throw; }
                Release(candidate);
            }
            throw new InvalidOperationException($"'{root.Name}'에서 Sheet '{name}'을 찾지 못했습니다.");
        }
        finally { Release(found); }
    }

    private static IAucObject GetOrCreateUnit(IAucObject parent, string name)
    {
        IAucObjects found = parent.FindObjects(AucObjectKind.aucObjUnit, AucVbSearchType.aucSearchFlat, AucType.aucTypeUnspecified, AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, name);
        try { if (found.Count > 0) return found.Item(1); }
        finally { Release(found); }
        IAucObject unit = parent.NewChild(AucObjectKind.aucObjUnit, AucType.aucUnit, false);
        try { SetDesignation(unit, name); return unit; }
        catch { Release(unit); throw; }
    }

    private static void SetDesignation(IAucObject item, string name)
    {
        IAucAttributes attributes = item.Attributes;
        IAucAttribute? designation = null;
        try
        {
            if (!attributes.ItemByIDSave(AucAttribute.aucAttrDesignation, out designation) || designation == null) designation = attributes.Add(AucAttribute.aucAttrDesignation);
            designation.Value = name;
            item.Store();
        }
        finally { Release(designation, attributes); }
    }

    private static HashSet<string> ReadEquipmentIds(IAucObject root)
    {
        var ids = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        root.Refresh();
        IAucObjects children = root.FindObjects(AucObjectKind.aucObjUnspecified, AucVbSearchType.aucSearchFlat);
        try
        {
            for (int index = 1; index <= children.Count; index++)
            {
                IAucObject item = children.Item(index);
                try { ids.Add(item.ID); }
                finally { Release(item); }
            }
        }
        finally { Release(children); }
        return ids;
    }

    private static string UniqueSheetName(IAucObject destination, string name)
    {
        string candidate = name;
        for (int suffix = 1; ; suffix++)
        {
            IAucObjects found = destination.FindObjects(AucObjectKind.aucObjSheet, AucVbSearchType.aucSearchFlat, AucType.aucTypeUnspecified, AucAttribute.aucAttrDesignation, AucVbFindCondition.aucCondEqual, candidate);
            try { if (found.Count == 0) return candidate; }
            finally { Release(found); }
            candidate = name + "(" + suffix + ")";
        }
    }

    private static void Release(params object?[] items)
    {
        foreach (object? item in items)
            if (item != null && Marshal.IsComObject(item)) Marshal.ReleaseComObject(item);
    }

    private sealed class TypicalDrawing
    {
        public IAucSheet Sheet = null!;
        public bool WasOpen, HasBounds;
        public double Left, Top, Width, Height;
    }

    internal sealed class LayoutCursor
    {
        private readonly double _incomerLeft, _incomerBottom, _right, _bottom, _columnGap, _rowGap, _feederOffsetX, _feederOffsetY;
        private readonly DrawingGroupLayoutOptions _groupLayout;
        private readonly bool _isContinuation;
        private double _x, _y, _rowBottom;
        private int _rowNumber = 1, _rowPlacementCount, _placementCount;
        private bool _rowContainsIncomer;

        public LayoutCursor(double incomerLeft, double incomerTop, double incomerWidth, double incomerHeight, double pageWidth, DrawingLayoutOptions options, DrawingGroupLayoutOptions groupLayout)
        {
            _groupLayout = groupLayout;
            _isContinuation = false;
            _incomerLeft = incomerLeft;
            _incomerBottom = incomerTop - incomerHeight;
            _right = pageWidth - options.RightMargin;
            _bottom = options.BottomMargin;
            _columnGap = groupLayout.ColumnGap ?? options.ColumnGap;
            _rowGap = groupLayout.RowGap ?? options.RowGap;
            _feederOffsetX = groupLayout.FeederOffsetX ?? options.FeederOffsetX;
            _feederOffsetY = groupLayout.FeederOffsetY ?? options.FeederOffsetY;
            if (groupLayout.PlacementMode == FeederPlacementMode.RightOfIncomer)
            {
                StartRow(incomerLeft + incomerWidth + _columnGap, incomerTop, true);
            }
            else
            {
                StartRow(incomerLeft, _incomerBottom - _rowGap, false);
            }
        }

        private LayoutCursor(double pageWidth, double pageHeight, DrawingLayoutOptions options, DrawingGroupLayoutOptions groupLayout)
        {
            _groupLayout = groupLayout;
            _isContinuation = true;
            _incomerLeft = options.LeftMargin;
            _incomerBottom = pageHeight - options.TopMargin;
            _right = pageWidth - options.RightMargin;
            _bottom = options.BottomMargin;
            _columnGap = groupLayout.ColumnGap ?? options.ColumnGap;
            _rowGap = groupLayout.RowGap ?? options.RowGap;
            _feederOffsetX = 0;
            _feederOffsetY = 0;
            StartRow(options.LeftMargin, pageHeight - options.TopMargin, false);
        }

        public static LayoutCursor CreateContinuation(double pageWidth, double pageHeight, DrawingLayoutOptions options, DrawingGroupLayoutOptions groupLayout)
        {
            return new LayoutCursor(pageWidth, pageHeight, options, groupLayout);
        }

        public bool TryPlace(double width, double height, out double left, out double top)
        {
            if (width <= 0 || height <= 0) throw new InvalidOperationException("Feeder 크기가 올바르지 않습니다.");
            if (_x + width > _right + 0.01)
            {
                if (_rowPlacementCount == 0 && !_rowContainsIncomer) throw new InvalidOperationException($"{_groupLayout.Name} {_rowNumber}행의 LeftMargin이 커서 Feeder를 배치할 수 없습니다.");
                MoveToNextRow();
            }
            if (_x + width > _right + 0.01) throw new InvalidOperationException($"{_groupLayout.Name} {_rowNumber}행의 LeftMargin이 커서 Feeder를 배치할 수 없습니다.");
            left = _x; top = _y;
            if (_y - height < _bottom - 0.01)
            {
                if (_placementCount == 0 && _isContinuation) throw new InvalidOperationException("Feeder가 빈 연속 Sheet의 배치 영역에도 들어가지 않습니다.");
                return false;
            }
            _x += width + _columnGap;
            _rowBottom = Math.Min(_rowBottom, _y - height);
            _rowPlacementCount++;
            _placementCount++;
            return true;
        }

        private void MoveToNextRow()
        {
            double nextRowTop = _rowBottom - _rowGap;
            _rowNumber++;
            StartRow(_incomerLeft, nextRowTop, false);
        }

        private void StartRow(double baseLeft, double baseTop, bool containsIncomer)
        {
            FeederRowMargin margin = _isContinuation ? new FeederRowMargin { RowNumber = _rowNumber } : _groupLayout.GetRowMargin(_rowNumber);
            _x = baseLeft + _feederOffsetX + margin.LeftMargin;
            _y = baseTop + (_rowNumber == 1 ? _feederOffsetY : 0) + margin.TopMargin;
            _rowBottom = containsIncomer ? _incomerBottom : double.PositiveInfinity;
            _rowPlacementCount = 0;
            _rowContainsIncomer = containsIncomer;
            Debug.WriteLine($"[Feeder {_rowNumber}행] {_groupLayout.Name} / Sheet={(_isContinuation ? "Continuation" : "Incomer")} / Mode={_groupLayout.PlacementMode} / Base=({baseLeft:F2}, {baseTop:F2})mm / Margin=({margin.LeftMargin:+0.00;-0.00;0.00}, {margin.TopMargin:+0.00;-0.00;0.00})mm / 배치=({_x:F2}, {_y:F2})mm");
        }
    }
}
