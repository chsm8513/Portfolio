using Aucotec.EngineeringBase.Interop;
using Auto_SLD.Models;
using EngineeringBase;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;

namespace Auto_SLD.Service;

internal sealed class DataService
{
    private readonly IEbRepository _repository;
    public List<LoadDeviceDto> loadDevices = new List<LoadDeviceDto>();
    private const string FindCableAndDestination1Formula = "R110;P;A5;\"$\";A1;\"$\";R1106;A5;\"$\";A1;";
    private const string FindCableAndDestination2Formula = "R110;P;A5;\"$\";A1;\"$\";R1107;A5;\"$\";A1;";

    public DataService(IEbRepository repository)
    {
        _repository = repository ?? throw new ArgumentNullException(nameof(repository));
    }

    public IReadOnlyList<LoadDeviceDto> GetLoadDevices(IAucProject project, IAucObject equipmentFolder, IProgress<ProgressInfo> progress)
    {
        IAucObject? worksheetFolder = null;
        IAucObject? loadDeviceConfiguration = null;
        IAucWorksheet? worksheet = null;

        try
        {
            progress.Report(new ProgressInfo("Load Device 조회 중", "Worksheet 설정을 찾는 중..."));

            worksheetFolder = project.WorksheetTemplatesFolder;
            IReadOnlyList<IAucObject> worksheetConfigurations = _repository.FindObjects(worksheetFolder, new EbObjectFilter(null, null, new Dictionary<int, string> { [(int)AucAttribute.aucAttrDesignation] = "Load Device SLD" }), true);

            if (worksheetConfigurations.Count > 0)
                loadDeviceConfiguration = worksheetConfigurations[0];

            for (int index = 1; index < worksheetConfigurations.Count; index++)
                EbComHelper.ReleaseComObject(worksheetConfigurations[index]);

            if (loadDeviceConfiguration == null)
                throw new InvalidOperationException("WorksheetTemplatesFolder 아래에서 이름(Designation)이 'Load Device SLD'인 객체를 찾을 수 없습니다.");

            progress.Report(new ProgressInfo("Load Device 조회 중", "Worksheet 실행 중..."));
            worksheet = equipmentFolder.OpenWorksheet(loadDeviceConfiguration, false);

            if (worksheet == null)
                throw new InvalidOperationException("Equipment 폴더에서 Load Device Worksheet를 실행하지 못했습니다.");

            loadDevices = new List<LoadDeviceDto>(worksheet.RowCount);
            progress.Report(new ProgressInfo("Load Device 조회 중", $"0 / {worksheet.RowCount:N0}"));

            void ReportRowProgress(int rowIndex)
            {
                if (rowIndex == worksheet.RowCount || rowIndex % 25 == 0) progress.Report(new ProgressInfo("Load Device 조회 중", $"{rowIndex:N0} / {worksheet.RowCount:N0}"));
            }

            for (int rowIndex = 1; rowIndex <= worksheet.RowCount; rowIndex++)
            {
                IAucObject item = worksheet.GetObjectItem(rowIndex);
                if (item == null)
                    continue;
                Debug.WriteLine($"DataServiceDto[{rowIndex}]: {item.Name}");
                item.ExecuteFormula(FindCableAndDestination1Formula, out string destination1Result);
                item.ExecuteFormula(FindCableAndDestination2Formula, out string destination2Result);
                if (!TryParseCableDestination(destination1Result, out string cable, out string cableId, out string switchGear, out string switchGearId)
                    || !TryParseCableDestination(destination2Result, out string destination2Cable, out string destination2CableId, out string destination2, out string destination2Id))
                {
                    Debug.WriteLine($"DataServiceDto[{item.Name}] 제외: Cable Destination 1 또는 Destination 2가 비어 있습니다. Dest1={destination1Result}, Dest2={destination2Result}");
                    ReportRowProgress(rowIndex);
                    continue;
                }

                string fullCableId = GetFullID(cableId);
                string fullDestination2CableId = GetFullID(destination2CableId);
                string fullSwitchGearId = GetFullID(switchGearId);
                string fullDestination2Id = GetFullID(destination2Id);
                if (string.IsNullOrEmpty(fullCableId) || string.IsNullOrEmpty(fullSwitchGearId) || string.IsNullOrEmpty(fullDestination2Id)
                    || !string.Equals(fullCableId, fullDestination2CableId, StringComparison.OrdinalIgnoreCase)
                    || !string.Equals(fullDestination2Id, item.ID, StringComparison.OrdinalIgnoreCase))
                {
                    Debug.WriteLine($"DataServiceDto[{item.Name}] 제외: Cable 관계가 일치하지 않습니다. Cable={cable}/{cableId}, Dest2Cable={destination2Cable}/{destination2CableId}, Dest2={destination2}/{destination2Id}, LoadId={item.ID}");
                    ReportRowProgress(rowIndex);
                    continue;
                }

                IAucObject switchGearObject = _repository.GetObjectById(fullSwitchGearId);
                string switchGearType;
                try { switchGearType = _repository.GetAttributeValue(switchGearObject, 104622) ?? string.Empty; }
                finally { EbComHelper.ReleaseComObject(switchGearObject); }
                Debug.WriteLine($"DataServiceDto[{item.Name}]: Dest1={destination1Result}, Dest2={destination2Result}");

                loadDevices.Add(new LoadDeviceDto
                {
                    SwitchGear = switchGear,
                    Cable = cable,
                    Designation = worksheet.GetText(rowIndex, 1),
                    Comment = worksheet.GetText(rowIndex, 2),
                    Type = worksheet.GetText(rowIndex, 3),
                    LoadType = worksheet.GetText(rowIndex, 4),
                    SwitchGearType = switchGearType,
                    FeederType = worksheet.GetText(rowIndex, 5),
                    SwitchGearId = fullSwitchGearId,
                    CableId = fullCableId,
                    LoadId = item.ID,
                });

                ReportRowProgress(rowIndex);
            }

            return loadDevices;
        }
        finally
        {
            if (worksheet != null)
            {
                try
                {
                    worksheet.Close();
                }
                catch (COMException)
                {
                    // EB 버전에 따라 임시 Worksheet Close를 허용하지 않을 수 있다.
                }
            }

            EbComHelper.ReleaseComObject(worksheet);
            EbComHelper.ReleaseComObject(loadDeviceConfiguration);
            EbComHelper.ReleaseComObject(worksheetFolder);
        }
    }

    public string GetFullID(string shortID)
    {
        if (string.IsNullOrWhiteSpace(shortID)) return string.Empty;
        if (Guid.TryParse(shortID, out Guid id)) return id == Guid.Empty ? string.Empty : id.ToString();
        if (shortID.Length < 8) return "";
        return "00000000-0000-0000-0000-0000" + shortID.Substring(shortID.Length - 8);
    }

    private static bool TryParseCableDestination(string result, out string cable, out string cableId, out string destination, out string destinationId)
    {
        cable = string.Empty;
        cableId = string.Empty;
        destination = string.Empty;
        destinationId = string.Empty;
        if (string.IsNullOrWhiteSpace(result)) return false;
        string[] values = result.Split('$');
        if (values.Length < 4) return false;
        cable = values[0].Trim();
        cableId = values[1].Trim();
        destination = values[2].Trim();
        destinationId = values[3].Trim();
        return !string.IsNullOrEmpty(cable) && !string.IsNullOrEmpty(cableId) && !string.IsNullOrEmpty(destination) && !string.IsNullOrEmpty(destinationId);
    }

}
