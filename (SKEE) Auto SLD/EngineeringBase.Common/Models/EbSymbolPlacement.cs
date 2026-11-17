namespace EngineeringBase.Common.Models;

/// <summary>
/// EB 객체 관계 및 배치 규칙에서 변환된 심볼 배치 명령입니다. 좌표 단위는 mm입니다.
/// </summary>
public sealed class EbSymbolPlacement
{
    public EbSymbolPlacement(int sourceIndex, string symbolName, double x, double y)
    {
        SourceIndex = sourceIndex;
        SymbolName = symbolName;
        X = x;
        Y = y;
    }

    public int SourceIndex { get; }
    public string SymbolName { get; }
    public double X { get; }
    public double Y { get; }
}
