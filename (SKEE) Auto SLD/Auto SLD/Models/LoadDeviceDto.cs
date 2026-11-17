using System.ComponentModel;

namespace Auto_SLD.Models;

public sealed class LoadDeviceDto
{   
    [DisplayName("Switch Gear")]
    public string SwitchGear { get; set; } = string.Empty;

    [DisplayName("Cable")]
    public string Cable { get; set; } = string.Empty;

    [DisplayName("Load")]
    public string Designation { get; set; } = string.Empty;

    [DisplayName("Comment")]
    public string Comment { get; set; } = string.Empty;

    [DisplayName("Type")]
    public string Type { get; set; } = string.Empty;

    [DisplayName("Load Type")]
    public string LoadType { get; set; } = string.Empty;

    [DisplayName("SwitchGearType")]
    public string SwitchGearType { get; set; } = string.Empty;

    [DisplayName("Feeder Type")]
    public string FeederType { get; set; } = string.Empty;

    [Browsable(false)]
    [DisplayName("SwitchGearID")]
    public string SwitchGearId { get; set; } = string.Empty;

    [Browsable(false)]
    [DisplayName("CableID")]
    public string CableId { get; set; } = string.Empty;

    [Browsable(false)]
    [DisplayName("LoadID")]
    public string LoadId { get; set; } = string.Empty;
}
