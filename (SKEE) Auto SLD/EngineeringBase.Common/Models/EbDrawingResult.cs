namespace EngineeringBase.Common.Models;

public sealed class EbDrawingResult
{
    public EbDrawingResult(int requestedCount, int placedCount, IReadOnlyList<string> errors)
    {
        RequestedCount = requestedCount;
        PlacedCount = placedCount;
        Errors = errors;
    }

    public int RequestedCount { get; }
    public int PlacedCount { get; }
    public IReadOnlyList<string> Errors { get; }
    public bool IsSuccess => Errors.Count == 0;
}
