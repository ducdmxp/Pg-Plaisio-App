using System;
using System.Runtime.InteropServices;
using System.Text;

namespace Convert2DTo3D.Helper.CSHelper.UserManipulation
{
    public class Dialog
    {
        protected delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowText(IntPtr hWnd, StringBuilder strText, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        protected static extern int GetWindowTextLength(IntPtr hWnd);

        [DllImport("user32.dll")]
        protected static extern bool EnumWindows(EnumWindowsProc enumProc, IntPtr lParam);

        [DllImport("user32.dll")]
        protected static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // Callback for examining the window
        protected static bool EnumTheWindows(IntPtr hWnd, IntPtr lParam)
        {
            int size = GetWindowTextLength(hWnd);
            if (size++ > 0 && IsWindowVisible(hWnd))
            {
                StringBuilder sb = new StringBuilder(size);
                GetWindowText(hWnd, sb, size);
                if (sb.ToString().Equals("Save Print Output As", StringComparison.Ordinal))
                {
                    uint procId = 0;
                    GetWindowThreadProcessId(hWnd, out procId);
                    Console.WriteLine($"Found it! ProcID: {procId}");
                }
            }
            return true;
        }
    }
}