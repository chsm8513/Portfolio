using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Visio = Microsoft.Office.Interop.Visio;

namespace Auto_SLD.Service;

// Visio의 OLE PasteSpecial이 아니라 EB의 붉은 붙여넣기 대기 영역을 확정한다.
internal sealed class EbPasteSpecialUiService
{
    internal const int PasteSpecialOffsetX = 50;
    internal const int PasteSpecialOffsetY = 30;

    public string Complete(Visio.Page page, double centerXmm, double centerYmm, IntPtr ebWindow)
    {
        Visio.Application app = page.Application;
        Visio.Window? drawing = null;
        IntPtr previousDpi = SetThreadDpiAwarenessContext(new IntPtr(-4));
        string stage = "대상 Visio 창 확인";
        try
        {
            drawing = app.ActiveWindow;
            IntPtr main = new IntPtr(app.WindowHandle32), canvas = new IntPtr(drawing.WindowHandle32);
            var processIds = new HashSet<int> { ProcessId(main), ProcessId(ebWindow) };
            processIds.Remove(0);
            if (!IsWindowEnabled(main)) throw new InvalidOperationException("Visio에 다른 처리창이 열려 있습니다. 해당 처리가 끝난 뒤 다시 실행하세요.");
            if (FindDialog(processIds) != IntPtr.Zero) throw new InvalidOperationException("이미 열린 Paste Special 창이 있습니다. 먼저 기존 작업을 마무리하세요.");

            stage = "Visio 창 최대화";
            if (IsIconic(main) || !IsZoomed(main))
            {
                if (!ShowWindowAsync(main, 3)) throw new InvalidOperationException("Visio 창에 최대화 요청을 전달하지 못했습니다."); // SW_MAXIMIZE
                WaitFor(() => IsZoomed(main) && !IsIconic(main) ? new object() : null, 3000, "Visio 창이 최대화되지 않았습니다.");
                Thread.Sleep(250); // 크기 변경 후 도면 화면 배치가 반영된 다음 좌표를 읽는다.
            }
            if (GetAncestor(GetForegroundWindow(), 2) != GetAncestor(main, 2) && !SetForegroundWindow(main))
                throw new InvalidOperationException("Visio를 전면으로 활성화하지 못했습니다. EB, Visio, Auto SLD의 실행 권한이 같은지 확인하세요.");

            stage = "붉은 영역 중앙 우클릭";
            drawing.GetViewRect(out double left, out double top, out double width, out double height);
            double x = centerXmm / 25.4, y = centerYmm / 25.4;
            // 도형은 이동하지 않는다. 중앙이 화면 밖인 경우 보기 영역만 이동한다.
            if (x <= left || x >= left + width || y >= top || y <= top - height)
            {
                drawing.SetViewRect(x - width / 2, y + height / 2, width, height);
                drawing.GetViewRect(out left, out top, out width, out height);
            }
            if (!GetClientRect(canvas, out Rect client)) throw new InvalidOperationException("Visio 도면 창의 화면 영역을 읽지 못했습니다.");
            var point = ToClientPoint(x, y, left, top, width, height, client.Right - client.Left, client.Bottom - client.Top);
            if (!ClientToScreen(canvas, ref point)) throw new InvalidOperationException("도면 중앙을 화면 좌표로 변환하지 못했습니다.");
            Click(point.X, point.Y, true);
            System.Threading.Thread.Sleep(250);

            stage = "Paste Special 메뉴 선택";
            // 메뉴 핸들/상태/이름은 찾지 않는다. 우클릭한 화면 좌표에서 고정 오프셋만큼 이동한다.
            Point menuItem = GetPasteSpecialClickPoint(point, GetDpiForWindow(main));
            Click(menuItem.X, menuItem.Y, false);
            System.Threading.Thread.Sleep(250);

            stage = "Paste Special 확인창 찾기";
            IntPtr dialogHandle = IntPtr.Zero;
            WaitFor(() => (dialogHandle = FindDialog(processIds)) != IntPtr.Zero ? new object() : null, 15000, "EB 'Paste Special' 확인창이 열리지 않았습니다.");
            stage = "Paste Special 확인창 활성화";
            if (GetForegroundWindow() != dialogHandle && !SetForegroundWindow(dialogHandle))
                throw new InvalidOperationException("Paste Special 확인창을 활성화하지 못했습니다.");
            stage = "기본값 OK 및 붙여넣기 확정 대기";
            // 라디오 옵션은 변경하지 않는다. 사용자가 요청한 기본 선택 그대로 OK만 누른다.
            PressKey(0x0D);
            System.Threading.Thread.Sleep(1000);
            WaitFor(() => !IsWindow(dialogHandle) ? new object() : null, 90000, "OK 이후에도 Paste Special 확인창이 닫히지 않았습니다. Visio의 처리 상태를 확인하세요.");
            return $"기본 선택 (옵션 변경 없음). 우클릭 위치 기준 오프셋 X={PasteSpecialOffsetX}, Y={PasteSpecialOffsetY} (96 DPI 기준)로 Paste Special 클릭, 확인창에서 Enter로 OK 실행.";
        }
        catch (Exception exception)
        {
            throw new EbPasteSpecialException($"EB 붙여넣기 확정 실패 [{stage}]: {exception.Message} 붉은 대기 상태를 완료로 처리하지 않고 중단합니다. 추가 EB 호출로 대기하지 않도록 Sheet 닫기/활성화와 Equipment 자동 정리는 생략합니다.", exception);
        }
        finally
        {
            if (previousDpi != IntPtr.Zero) SetThreadDpiAwarenessContext(previousDpi);
            if (drawing != null && Marshal.IsComObject(drawing)) Marshal.ReleaseComObject(drawing);
            if (Marshal.IsComObject(app)) Marshal.ReleaseComObject(app);
        }
    }

    internal static Point ToClientPoint(double x, double y, double left, double top, double width, double height, int clientWidth, int clientHeight)
    {
        if (new[] { x, y, left, top, width, height }.Any(value => double.IsNaN(value) || double.IsInfinity(value)) || width <= 0 || height <= 0 || clientWidth <= 0 || clientHeight <= 0)
            throw new InvalidOperationException("Visio 화면 좌표 변환 값이 올바르지 않습니다.");
        double px = (x - left) / width * clientWidth, py = (top - y) / height * clientHeight;
        if (px <= 0 || px >= clientWidth || py <= 0 || py >= clientHeight) throw new InvalidOperationException("붙여넣기 영역 중앙이 Visio 화면 밖입니다.");
        return new Point { X = (int)Math.Round(px), Y = (int)Math.Round(py) };
    }

    internal static bool IsPasteSpecialName(string name)
    {
        return string.Equals(NormalizeName(name), "Paste Special", StringComparison.OrdinalIgnoreCase);
    }

    private static string NormalizeName(string name) => name.Replace("&", "").Replace("…", "").Replace("...", "").Trim();

    internal static Point GetPasteSpecialClickPoint(Point anchor, uint dpi)
    {
        // 100% 화면 배율에서 우클릭 위치의 오른쪽 50px, 아래쪽 30px가 두 번째 메뉴 항목이다.
        double scale = dpi == 0 ? 1.0 : dpi / 96.0;
        return new Point { X = anchor.X + (int)Math.Round(PasteSpecialOffsetX * scale), Y = anchor.Y + (int)Math.Round(PasteSpecialOffsetY * scale) };
    }

    private static void PressKey(ushort key)
    {
        var inputs = new[]
        {
            new KeyboardEvent { Type = 1, Data = new InputData { Keyboard = new KeyboardInput { VirtualKey = key } } },
            new KeyboardEvent { Type = 1, Data = new InputData { Keyboard = new KeyboardInput { VirtualKey = key, Flags = 0x0002 } } }
        };
        if (SendKeyboardInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(KeyboardEvent))) != inputs.Length)
            throw new InvalidOperationException("Paste Special 키보드 입력을 전달하지 못했습니다. EB, Visio, Auto SLD의 실행 권한이 같은지 확인하세요.");
    }

    private static IntPtr FindDialog(HashSet<int> processIds)
    {
        IntPtr found = IntPtr.Zero;
        EnumWindows((window, _) =>
        {
            if (!IsWindowVisible(window) || !processIds.Contains(ProcessId(window))) return true;
            var title = new StringBuilder(256);
            GetWindowText(window, title, title.Capacity);
            if (!IsPasteSpecialName(title.ToString())) return true;
            found = window;
            return false;
        }, IntPtr.Zero);
        return found;
    }

    private static T WaitFor<T>(Func<T?> read, int timeoutMs, string message) where T : class
    {
        var timer = Stopwatch.StartNew();
        do
        {
            T? result = read();
            if (result != null) return result;
            Thread.Sleep(100);
        } while (timer.ElapsedMilliseconds < timeoutMs);
        throw new InvalidOperationException(message);
    }

    private static void Click(int x, int y, bool rightButton)
    {
        if (!SetCursorPos(x, y)) throw new InvalidOperationException("마우스를 붙여넣기 위치로 이동하지 못했습니다.");
        var inputs = new[]
        {
            new Input { Type = 0, Mouse = new MouseInput { Flags = rightButton ? 0x0008u : 0x0002u } },
            new Input { Type = 0, Mouse = new MouseInput { Flags = rightButton ? 0x0010u : 0x0004u } }
        };
        if (SendInput((uint)inputs.Length, inputs, Marshal.SizeOf(typeof(Input))) != inputs.Length) throw new InvalidOperationException("마우스 입력을 전달하지 못했습니다. EB, Visio, Auto SLD의 실행 권한이 같은지 확인하세요.");
    }

    private static int ProcessId(IntPtr window)
    {
        if (window == IntPtr.Zero) return 0;
        GetWindowThreadProcessId(window, out uint id);
        return (int)id;
    }

    [StructLayout(LayoutKind.Sequential)]
    internal struct Point { public int X, Y; }
    [StructLayout(LayoutKind.Sequential)]
    private struct Rect { public int Left, Top, Right, Bottom; }
    [StructLayout(LayoutKind.Sequential)]
    private struct MouseInput { public int X, Y; public uint Data, Flags, Time; public UIntPtr ExtraInfo; }
    [StructLayout(LayoutKind.Sequential)]
    private struct Input { public uint Type; public MouseInput Mouse; }

    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardInput { public ushort VirtualKey, ScanCode; public uint Flags, Time; public UIntPtr ExtraInfo; }
    [StructLayout(LayoutKind.Explicit)]
    private struct InputData
    {
        [FieldOffset(0)] public MouseInput Mouse;
        [FieldOffset(0)] public KeyboardInput Keyboard;
    }
    [StructLayout(LayoutKind.Sequential)]
    private struct KeyboardEvent { public uint Type; public InputData Data; }
    [DllImport("user32.dll")] private static extern IntPtr SetThreadDpiAwarenessContext(IntPtr context);
    [DllImport("user32.dll")] private static extern bool SetForegroundWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")] private static extern IntPtr GetAncestor(IntPtr window, uint flags);
    [DllImport("user32.dll")] private static extern bool IsWindowEnabled(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsWindowVisible(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsZoomed(IntPtr window);
    [DllImport("user32.dll")] private static extern bool IsIconic(IntPtr window);
    [DllImport("user32.dll")] private static extern bool ShowWindowAsync(IntPtr window, int command);
    [DllImport("user32.dll")] private static extern bool GetClientRect(IntPtr window, out Rect rect);
    [DllImport("user32.dll", CharSet = CharSet.Unicode)] private static extern int GetWindowText(IntPtr window, StringBuilder text, int count);
    private delegate bool EnumWindowCallback(IntPtr window, IntPtr parameter);
    [DllImport("user32.dll")] private static extern bool EnumWindows(EnumWindowCallback callback, IntPtr parameter);
    [DllImport("user32.dll", EntryPoint = "SendInput", SetLastError = true)] private static extern uint SendKeyboardInput(uint count, KeyboardEvent[] inputs, int size);
    [DllImport("user32.dll")] private static extern uint GetDpiForWindow(IntPtr window);
    [DllImport("user32.dll")] private static extern bool ClientToScreen(IntPtr window, ref Point point);
    [DllImport("user32.dll")] private static extern uint GetWindowThreadProcessId(IntPtr window, out uint processId);
    [DllImport("user32.dll")] private static extern bool SetCursorPos(int x, int y);
    [DllImport("user32.dll", SetLastError = true)] private static extern uint SendInput(uint count, Input[] inputs, int size);
}

internal sealed class EbPasteSpecialException : InvalidOperationException
{
    public EbPasteSpecialException(string message, Exception innerException) : base(message, innerException) { }
}
