using System.Runtime.InteropServices;

namespace EngineeringBase.Common;

internal static class EbCom
{
    public static void Release(object? instance)
    {
        if (instance is not null && Marshal.IsComObject(instance))
        {
            Marshal.FinalReleaseComObject(instance);
        }
    }
}
