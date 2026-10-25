// ---------------------------------------------------------------------------------------------------------------
// Name:     NativeMethods.cs
// Purpose:  The only place Windows API calls are declared. Handles are IntPtr so one Any CPU DLL works in
//           32-bit and 64-bit Office.
// Revision: 17
// DTG:      261024233000
// SPDX-License-Identifier: MIT
// ---------------------------------------------------------------------------------------------------------------

using System;
using System.Runtime.InteropServices;

namespace OpenTools
{
    internal static class NativeMethods
    {
        // ----- Constants (names match the Windows headers) -----

        public const uint SWP_NOSIZE = 0x0001;
        public const uint SWP_NOMOVE = 0x0002;
        public const uint SWP_NOZORDER = 0x0004;
        public const uint SWP_NOACTIVATE = 0x0010;
        public const uint SWP_SHOWWINDOW = 0x0040;

        public const int SW_HIDE = 0;
        public const int SW_SHOWNORMAL = 1;
        public const int SW_SHOWMINIMIZED = 2;
        public const int SW_SHOWMAXIMIZED = 3;
        public const int SW_SHOWNOACTIVATE = 4;
        public const int SW_RESTORE = 9;

        public const uint GA_PARENT = 1;
        public const uint GA_ROOT = 2;
        public const uint GW_HWNDNEXT = 2;
        public const uint GW_CHILD = 5;

        public const uint RDW_INVALIDATE = 0x0001;
        public const uint RDW_ERASE = 0x0004;
        public const uint RDW_ALLCHILDREN = 0x0080;
        public const uint RDW_UPDATENOW = 0x0100;
        public const uint RDW_FRAME = 0x0400;
        public const int WM_CLOSE = 0x0010;
        public const int WM_SETREDRAW = 0x000B;

        public const int WH_CALLWNDPROC = 4;
        public const int WH_CBT = 5;
        public const int HCBT_CREATEWND = 3;
        public const int HCBT_MINMAX = 1;
        public const int HCBT_ACTIVATE = 5;
        public const int WM_SHOWWINDOW = 0x0018;
        public const int DWMWA_TRANSITIONS_FORCEDISABLED = 3;
        public const int DWMWA_CLOAK = 13;

        public const int WM_WINDOWPOSCHANGING = 0x0046;
        public const int WM_WINDOWPOSCHANGED = 0x0047;
        public const int WS_CLIPSIBLINGS = 0x04000000;

        public const int WS_EX_TOOLWINDOW = 0x00000080;
        public const int WS_EX_TRANSPARENT = 0x00000020;
        public const int WS_EX_LAYERED = 0x00080000;
        public const int GWL_EXSTYLE = -20;
        public const uint LWA_ALPHA = 0x00000002;
        public const int WS_EX_NOACTIVATE = 0x08000000;

        public const uint SHGFI_ICON = 0x00000100;
        public const uint SHGFI_SMALLICON = 0x00000001;
        public const uint SHGFI_USEFILEATTRIBUTES = 0x00000010;

        public const uint FILE_ATTRIBUTE_NORMAL = 0x00000080;

        // ----- Structures -----

        [StructLayout(LayoutKind.Sequential)]
        public struct RECT
        {
            public int Left;
            public int Top;
            public int Right;
            public int Bottom;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPOS
        {
            public IntPtr Hwnd;
            public IntPtr HwndInsertAfter;
            public int X;
            public int Y;
            public int Cx;
            public int Cy;
            public uint Flags;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CREATESTRUCT
        {
            public IntPtr CreateParams;
            public IntPtr Instance;
            public IntPtr Menu;
            public IntPtr Parent;
            public int Cy;
            public int Cx;
            public int Y;
            public int X;
            public int Style;
            public IntPtr Name;
            public IntPtr ClassName;
            public int ExStyle;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct CBT_CREATEWND
        {
            public IntPtr CreateStruct;
            public IntPtr InsertAfter;
        }

        public delegate IntPtr HookProc(int code, IntPtr wParam, IntPtr lParam);

        public delegate bool EnumWindowsProc(IntPtr hWnd, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        public struct CWPSTRUCT
        {
            public IntPtr LParam;
            public IntPtr WParam;
            public uint Message;
            public IntPtr Hwnd;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WINDOWPLACEMENT
        {
            public uint Length;
            public uint Flags;
            public uint ShowCmd;
            public POINT MinPosition;
            public POINT MaxPosition;
            public RECT NormalPosition;
        }

        [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
        public struct SHFILEINFO
        {
            public IntPtr HIcon;
            public int IIcon;
            public uint Attributes;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string DisplayName;

            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 80)]
            public string TypeName;
        }

        // ----- Window functions -----

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool EnumWindows(EnumWindowsProc callback, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowRect(IntPtr hWnd, out RECT rect);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetWindowPos(IntPtr hWnd, IntPtr hWndInsertAfter, int x, int y, int cx, int cy,
            uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool PostMessage(IntPtr hWnd, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsWindowVisible(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsIconic(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool IsZoomed(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetWindow(IntPtr hWnd, uint command);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, EntryPoint = "GetWindowTextW", SetLastError = true)]
        public static extern int GetWindowText(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern int GetClassName(IntPtr hWnd, System.Text.StringBuilder text, int maxCount);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool ScreenToClient(IntPtr hWnd, ref POINT point);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetAncestor(IntPtr hWnd, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT placement);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool DestroyIcon(IntPtr hIcon);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool RedrawWindow(IntPtr hWnd, IntPtr updateRect, IntPtr updateRegion, uint flags);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr SendMessage(IntPtr hWnd, int message, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        public static extern IntPtr SetWindowsHookEx(int hookId, HookProc callback, IntPtr module, uint threadId);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool UnhookWindowsHookEx(IntPtr hook);

        [DllImport("user32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr CallNextHookEx(IntPtr hook, int code, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", EntryPoint = "GetWindowLong", SetLastError = true)]
        private static extern int GetWindowLong32(IntPtr hWnd, int index);

        [DllImport("user32.dll", EntryPoint = "GetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr GetWindowLong64(IntPtr hWnd, int index);

        [DllImport("user32.dll", EntryPoint = "SetWindowLong", SetLastError = true)]
        private static extern int SetWindowLong32(IntPtr hWnd, int index, int value);

        [DllImport("user32.dll", EntryPoint = "SetWindowLongPtr", SetLastError = true)]
        private static extern IntPtr SetWindowLong64(IntPtr hWnd, int index, IntPtr value);

        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        public static extern bool SetLayeredWindowAttributes(IntPtr hWnd, uint colorKey, byte alpha, uint flags);

        public static int GetExStyle(IntPtr window)
        {
            return IntPtr.Size == 8
                ? (int)GetWindowLong64(window, GWL_EXSTYLE).ToInt64()
                : GetWindowLong32(window, GWL_EXSTYLE);
        }

        public static void SetExStyle(IntPtr window, int style)
        {
            if (IntPtr.Size == 8)
            {
                SetWindowLong64(window, GWL_EXSTYLE, new IntPtr(style));
            }
            else
            {
                SetWindowLong32(window, GWL_EXSTYLE, style);
            }
        }

        [DllImport("dwmapi.dll")]
        public static extern int DwmSetWindowAttribute(IntPtr hWnd, int attribute, ref int value, int size);

        [DllImport("dwmapi.dll")]
        public static extern int DwmGetWindowAttribute(IntPtr hWnd, int attribute, ref int value, int size);

        [DllImport("kernel32.dll")]
        public static extern uint GetCurrentThreadId();

        // ----- Shell functions -----

        [DllImport("shell32.dll", CharSet = CharSet.Unicode)]
        public static extern IntPtr SHGetFileInfo(string path, uint fileAttributes, ref SHFILEINFO info,
            uint infoSize, uint flags);

        // ----- Accessibility (MSAA), used to read and select the ribbon tab -----

        public const uint OBJID_CLIENT = 0xFFFFFFFC;
        public const int ROLE_SYSTEM_PAGETAB = 0x25;
        public const int STATE_SYSTEM_SELECTED = 0x2;
        public const int STATE_SYSTEM_INVISIBLE = 0x8000;
        public static readonly Guid IID_IAccessible = new Guid("618736E0-3C3D-11CF-810C-00AA00389B71");

        [DllImport("oleacc.dll")]
        public static extern int AccessibleObjectFromWindow(IntPtr hWnd, uint objectId, ref Guid riid,
            [MarshalAs(UnmanagedType.Interface)] out IAccessible accessible);

        [DllImport("oleacc.dll")]
        public static extern int AccessibleChildren(IAccessible container, int childStart, int children,
            [Out, MarshalAs(UnmanagedType.LPArray, SizeParamIndex = 2)] object[] found, out int obtained);
    }

    // The standard IAccessible interface from oleacc. Methods are in vtable order, after IDispatch.
    [ComImport]
    [Guid("618736E0-3C3D-11CF-810C-00AA00389B71")]
    [InterfaceType(ComInterfaceType.InterfaceIsDual)]
    internal interface IAccessible
    {
        [return: MarshalAs(UnmanagedType.IDispatch)]
        object get_accParent();

        int get_accChildCount();

        [return: MarshalAs(UnmanagedType.IDispatch)]
        object get_accChild(object child);

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accName(object child);

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accValue(object child);

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accDescription(object child);

        object get_accRole(object child);

        object get_accState(object child);

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accHelp(object child);

        int get_accHelpTopic([MarshalAs(UnmanagedType.BStr)] out string helpFile, object child);

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accKeyboardShortcut(object child);

        object get_accFocus();

        object get_accSelection();

        [return: MarshalAs(UnmanagedType.BStr)]
        string get_accDefaultAction(object child);

        void accSelect(int flags, object child);

        void accLocation(out int left, out int top, out int width, out int height, object child);

        object accNavigate(int direction, object start);

        object accHitTest(int x, int y);

        void accDoDefaultAction(object child);

        void put_accName(object child, [MarshalAs(UnmanagedType.BStr)] string name);

        void put_accValue(object child, [MarshalAs(UnmanagedType.BStr)] string value);
    }
}
