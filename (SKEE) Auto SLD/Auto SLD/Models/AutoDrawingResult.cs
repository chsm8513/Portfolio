namespace Auto_SLD.Models;

internal sealed class AutoDrawingResult
{
    public int SheetCount { get; set; }
    public int FeederCount { get; set; }
    public int ReassociatedSymbolCount { get; set; }
    public int RemovedDummyCount { get; set; }
    public int MovedEquipmentCount { get; set; }
}
