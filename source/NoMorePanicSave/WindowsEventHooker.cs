using System;
using System.Runtime.InteropServices;

namespace NoMorePanicSave
{
    public class WindowsEventHooker
    {
        public delegate void WinEventDelegate(
                IntPtr hWinEventHook, uint eventType,
                IntPtr hWnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        [DllImport("user32.dll")]
        public static extern IntPtr SetWinEventHook(
              uint eventMin, uint eventMax, IntPtr hmodWinEventProc, WinEventDelegate lpfnWinEventProc,
              uint idProcess, uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        public static extern bool UnhookWinEvent(IntPtr hWinEventHook);
    }

    public static class SetWinEventHookFlags
    {
        public const uint WINEVENT_INCONTEXT = 4;
        public const uint WINEVENT_OUTOFCONTEXT = 0;
        public const uint WINEVENT_SKIPOWNPROCESS = 2;
        public const uint WINEVENT_SKIPOWNTHREAD = 1;
    }
}