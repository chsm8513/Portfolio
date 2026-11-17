namespace Auto_SLD.Models;

internal enum FeederPlacementMode
{
    BelowIncomer,
    RightOfIncomer
}

internal sealed class FeederRowMargin
{
    public int RowNumber { get; set; }
    // 양수 LeftMargin은 오른쪽, 양수 TopMargin은 위쪽으로 이동한다.
    public double LeftMargin { get; set; }
    public double TopMargin { get; set; }
}

internal sealed class DrawingGroupLayoutOptions
{
    public string Name { get; set; } = string.Empty;
    public string[] IncomerTypes { get; set; } = Array.Empty<string>();
    public string[] FeederTypes { get; set; } = Array.Empty<string>();
    public string[] LastFeederTypes { get; set; } = Array.Empty<string>();
    public FeederPlacementMode PlacementMode { get; set; } = FeederPlacementMode.BelowIncomer;
    public double? ColumnGap { get; set; }
    public double? RowGap { get; set; }
    public double? FeederOffsetX { get; set; }
    public double? FeederOffsetY { get; set; }
    public FeederRowMargin[] RowMargins { get; set; } = Array.Empty<FeederRowMargin>();

    public bool Matches(string incomerType, IEnumerable<string> feederTypes)
    {
        if (!IncomerTypes.Contains(incomerType, StringComparer.OrdinalIgnoreCase)) return false;
        if (FeederTypes.Length == 0) return true;
        return feederTypes.All(feederType => FeederTypes.Contains(feederType, StringComparer.OrdinalIgnoreCase));
    }

    public FeederRowMargin GetRowMargin(int rowNumber)
    {
        return RowMargins.FirstOrDefault(row => row.RowNumber == rowNumber) ?? new FeederRowMargin { RowNumber = rowNumber };
    }

    public bool IsLastFeederType(string feederType)
    {
        return LastFeederTypes.Contains(feederType?.Trim() ?? string.Empty, StringComparer.OrdinalIgnoreCase);
    }
}

internal sealed class DrawingLayoutOptions
{
    // Incomer BoundingBox의 왼쪽 상단을 Sheet 가장자리에서 떨어뜨릴 거리(mm).
    public double LeftMargin { get; set; } = 50;
    public double TopMargin { get; set; } = 50;
    public double RightMargin { get; set; } = 50;
    public double BottomMargin { get; set; } = 0;
    public double ColumnGap { get; set; } = 0;
    public double RowGap { get; set; } = 0;
    // 음수 X는 왼쪽, 양수 Y는 위쪽. 그룹 설정에 값이 있으면 그룹 설정을 우선한다.
    public double FeederOffsetX { get; set; } = 0;
    public double FeederOffsetY { get; set; } = 0;

    public DrawingGroupLayoutOptions DefaultGroupLayout { get; } = new DrawingGroupLayoutOptions { Name = "Default", PlacementMode = FeederPlacementMode.BelowIncomer };

    // Incomer/Feeder 조합별 배치 설정. 위에서 먼저 일치하는 설정을 사용한다.
    public List<DrawingGroupLayoutOptions> GroupLayouts { get; } = new List<DrawingGroupLayoutOptions>
    {
        new DrawingGroupLayoutOptions
        {
            Name = "MVI + MVF",
            IncomerTypes = new[] { "MV1I", "MV2I-N" },
            FeederTypes = new[] { "MVF", "MVF-F", "MVF-M1", "MVF-M2", "MVF-SP-F", "MVF-SP-M", "MV2I-B", "MV2I-B2" },
            LastFeederTypes = new[] { "MV2I-B", "MV2I-B2" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = 15 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        /*new DrawingGroupLayoutOptions
        {
            Name = "MV2I-B + MVF",
            IncomerTypes = new[] { "MV2I-B" },
            FeederTypes = new[] { "MVF", "MVF-F", "MVF-M1", "MVF-M2", "MVF-SP-F", "MVF-SP-M" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },*/
        new DrawingGroupLayoutOptions
        {
            Name = "LVI/LV1I + LVF/LVM",
            IncomerTypes = new[] { "LVI", "LV1I", "LV2I-B" },
            FeederTypes = new[] { "LVF1", "LVF2", "LVM1", "LVM2", "LVSP-F", "LVSP-M" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = 8 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        new DrawingGroupLayoutOptions
        {
            Name = "LV2I + LVF/LVM",
            IncomerTypes = new[] { "LV2I", "LV2I-A" },
            FeederTypes = new[] { "LVF1", "LVF2", "LVM1", "LVM2", "LVSP-F", "LVSP-M" },
            PlacementMode = FeederPlacementMode.RightOfIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = -171 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        new DrawingGroupLayoutOptions
        {
            Name = "DC-M + DC-F",
            IncomerTypes = new[] { "DC-M" },
            FeederTypes = new[] { "FEEDER", "DC-F1", "DC-F2" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 57, TopMargin = 24 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        new DrawingGroupLayoutOptions
        {
            Name = "UPS-M + UPS-F",
            IncomerTypes = new[] { "UPS-M" },
            FeederTypes = new[] { "UPS-F1", "UPS-F2" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 22, TopMargin = 33 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        new DrawingGroupLayoutOptions
        {
            Name = "MCI + MCF/MCM",
            IncomerTypes = new[] { "MCI" },
            FeederTypes = new[] { "MCF1", "MCF2", "MCM1", "MCM2", "MCM3", "MCSP-F", "MCSP-M", "MCI2", "MCAUX" },
            LastFeederTypes = new[] { "MCI2", "MCAUX" },
            PlacementMode = FeederPlacementMode.RightOfIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = -65 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        /*new DrawingGroupLayoutOptions
        {
            Name = "MCI2 + MCF/MCM",
            IncomerTypes = new[] { "MCI2" },
            FeederTypes = new[] { "MCF1", "MCF2", "MCM1", "MCM2", "MCM3", "MCSP-F", "MCSP-M", "MCAUX" },
            PlacementMode = FeederPlacementMode.RightOfIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },*/
        new DrawingGroupLayoutOptions
        {
            Name = "PDPI + F1/F2/FEEDER",
            IncomerTypes = new[] { "PDPI", "PDP1I", "PDP2I" },
            FeederTypes = new[] { "F1", "F2", "FEEDER", "PDPF-SP" },
            PlacementMode = FeederPlacementMode.RightOfIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 0, TopMargin = -33 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        },
        new DrawingGroupLayoutOptions
        {
            Name = "PDPNM + F1/F2/FEEDER",
            IncomerTypes = new[] { "PDPN", "PDPM-ELD", "PDPM-GFR" },
            FeederTypes = new[] { "F1", "F2", "FEEDER", "PDPF-SP" },
            PlacementMode = FeederPlacementMode.BelowIncomer,
            RowMargins = new[]
            {
                new FeederRowMargin { RowNumber = 1, LeftMargin = 10, TopMargin = 68 },
                new FeederRowMargin { RowNumber = 2, LeftMargin = 0, TopMargin = 0 },
                new FeederRowMargin { RowNumber = 3, LeftMargin = 0, TopMargin = 0 }
            }
        }        
    };

    public DrawingGroupLayoutOptions ResolveGroupLayout(string incomerType, IEnumerable<string> feederTypes)
    {
        string[] types = feederTypes.Distinct(StringComparer.OrdinalIgnoreCase).ToArray();
        return GroupLayouts.FirstOrDefault(group => group.Matches(incomerType, types)) ?? DefaultGroupLayout;
    }
}
