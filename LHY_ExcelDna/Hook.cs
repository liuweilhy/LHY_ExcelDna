using System;
using System.ComponentModel;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;

namespace Hook
{
    public class KeyboardHook
    {
        #region Define

        private const Int32 HC_ACTION = 0;

        private const Int32 WM_KEYDOWN = 0x100;
        private const Int32 WM_KEYUP = 0x101;
        private const Int32 WM_SYSKEYDOWN = 0x0104;
        private const Int32 WM_SYSKEYUP = 0x105;

        private const Int32 WH_JOURNALRECORD = 0;
        private const Int32 WH_JOURNALPLAYBACK = 1;
        private const Int32 WH_KEYBOARD = 2;
        private const Int32 WH_GETMESSAGE = 3;
        private const Int32 WH_CALLWNDPROC = 4;
        private const Int32 WH_CBT = 5;
        private const Int32 WH_SYSMSGFILTER = 6;
        private const Int32 WH_MOUSE = 7;
        private const Int32 WH_HARDWARE = 8;
        private const Int32 WH_DEBUG = 9;
        private const Int32 WH_SHELL = 10;
        private const Int32 WH_FOREGROUNDIDLE = 11;
        private const Int32 WH_CALLWNDPROCRET = 12;
        private const Int32 WH_KEYBOARD_LL = 13;
        private const Int32 WH_MOUSE_LL = 14;

        #endregion Define

        #region Hook

        /// <summary>
        /// An application-defined or library-defined callback function used with the SetWindowsHookEx function.
        /// </summary>
        /// <param name="code">A code the hook procedure uses to determine how to process the message.</param>
        /// <param name="wParam">virtual-key code</param>
        /// <param name="lParam">The repeat count, scan code, extended-key flag, context code, previous key-state flag, and transition-state flag. </param>
        /// <returns>If code is less than zero, the hook procedure must return the value returned by CallNextHookEx. If code is greater than or equal to zero, and the hook procedure did not process the message, it is highly recommended that you call CallNextHookEx and return the value it returns; otherwise bad stuff.</returns>
        private delegate int HookProc(int code, Int32 wParam, Int32 lParam);

        /// <summary>
        /// Passes the hook information to the next hook procedure in the current hook chain. A hook procedure can call this function either before or after processing the hook information.
        /// </summary>
        /// <param name="hhk">This parameter is ignored.</param>
        /// <param name="nCode">The hook code passed to the current hook procedure. The next hook procedure uses this code to determine how to process the hook information.</param>
        /// <param name="wParam">The wParam value passed to the current hook procedure. The meaning of this parameter depends on the type of hook associated with the current hook chain.</param>
        /// <param name="lParam">The lParam value passed to the current hook procedure. The meaning of this parameter depends on the type of hook associated with the current hook chain.</param>
        /// <returns>This value is returned by the next hook procedure in the chain. The current hook procedure must also return this value. The meaning of the return value depends on the hook type. For more information, see the descriptions of the individual hook procedures.</returns>
        [DllImport("user32.dll")]
        private static extern int CallNextHookEx(IntPtr hhk, int nCode, Int32 wParam, Int32 lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        /// <summary>
        /// Get Current Thread Id
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        /// <summary>
        /// Installs an application-defined hook procedure into a hook chain.
        /// </summary>
        /// <param name="idHook">The type of hook procedure to be installed.</param>
        /// <param name="lpfn">Reference to the hook callback method.</param>
        /// <param name="hMod">A handle to the DLL containing the hook procedure pointed to by the lpfn parameter. The hMod parameter must be set to NULL if the dwThreadId parameter specifies a thread created by the current process and if the hook procedure is within the code associated with the current process.</param>
        /// <param name="dwThreadId">The identifier of the thread with which the hook procedure is to be associated. If this parameter is zero, the hook procedure is associated with all existing threads running in the same desktop as the calling thread.</param>
        /// <returns>If the function succeeds, the return value is the handle to the hook procedure. If the function fails, the return value is NULL. To get extended error information, call GetLastError.</returns>
        [DllImport("user32.dll")]
        private static extern IntPtr SetWindowsHookEx(Int32 idHook, HookProc lpfn, IntPtr hMod, int dwThreadId);

        /// <summary>
        /// Removes a hook procedure installed in a hook chain by the SetWindowsHookEx function.
        /// </summary>
        /// <param name="hhk">A handle to the hook to be removed. This parameter is a hook handle obtained by a previous call to SetWindowsHookEx.</param>
        /// <returns>If the function succeeds, the return value is nonzero. If the function fails, the return value is zero. To get extended error information, call GetLastError.</returns>
        [DllImport("user32.dll")]
        private static extern int UnhookWindowsHookEx(IntPtr hhk);

        #endregion Hook

        /// <summary>
        /// 钩子类名
        /// </summary>
        public string Name { get; private set; }

        /// <summary>
        /// 暂停标识
        /// </summary>
        private bool ispaused = false;
        public bool IsPaused
        {
            get
            {
                return ispaused;
            }
            set
            {
                if (value != ispaused && value == true)
                {
                    StopHook();
                }

                if (value != ispaused && value == false)
                {
                    StartHook();
                }

                ispaused = value;
            }
        }

        /// <summary>
        /// 钩子处理函数的句柄
        /// </summary>
        private IntPtr hhook;

        /// <summary>
        /// 钩子处理函数的指针
        /// </summary>
        private HookProc hookproc;

        /// <summary>
        /// 键盘事件委托声明
        /// </summary>
        /// <param name="e"></param>
        public delegate void KeyEventDelegate(KeyboardHookEventArgs e);

        /// <summary>
        /// 键盘按下事件委托
        /// </summary>
        public KeyEventDelegate KeyDownEvent = delegate { };

        /// <summary>
        /// 键盘松开事件委托
        /// </summary>
        public KeyEventDelegate KeyUpEvent = delegate { };

        public KeyboardHook(string name = "")
        {
            Name = name;
            StartHook();
        }

        ~KeyboardHook()
        {
            StopHook();
        }

        /// <summary>
        /// KeyboardProc callback function
        /// </summary>
        /// <param name="code"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        private int HookCallback(int code, Int32 wParam, Int32 lParam)
        {
            // https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644984(v=vs.85)

            int result = 0;
            try
            {
                // Debug.WriteLine("code: " + code.ToString());
                if (code == HC_ACTION && !ispaused)
                {
                    // 键盘按下
                    if (lParam > 0)
                    {
                        KeyDownEvent(new KeyboardHookEventArgs(wParam, lParam));
                    }
                    // 键盘松开
                    else if (lParam < 0)
                    {
                        KeyUpEvent(new KeyboardHookEventArgs(wParam, lParam));
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }
            finally
            {
                result = CallNextHookEx(IntPtr.Zero, code, wParam, lParam);
            }

            return result;
        }

        private void StartHook()
        {
            Debug.WriteLine(string.Format("Starting hook '{0}'...", Name), string.Format("Hook.StartHook [{0}]", Thread.CurrentThread.Name));

            hookproc = new HookProc(HookCallback);
            hhook = SetWindowsHookEx(WH_KEYBOARD, hookproc,
                IntPtr.Zero, GetCurrentThreadId());
            if (hhook == null || hhook == IntPtr.Zero)
            {
                Win32Exception LastError = new Win32Exception(Marshal.GetLastWin32Error());
            }
        }

        private void StopHook()
        {
            Debug.WriteLine(string.Format("Stopping hook '{0}'...", Name), string.Format("Hook.StartHook [{0}]", Thread.CurrentThread.Name));

            UnhookWindowsHookEx(hhook);
        }
    }

    public class KeyboardHookEventArgs
    {
        private const int KEY_PRESSED = 0x8000;

        private enum VirtualKeyStates : int
        {
            VK_LWIN = 0x5B,
            VK_RWIN = 0x5C,
            VK_LSHIFT = 0xA0,
            VK_RSHIFT = 0xA1,
            VK_LCONTROL = 0xA2,
            VK_RCONTROL = 0xA3,
            VK_LALT = 0xA4,
            VK_RALT = 0xA5
        };

        [DllImport("user32.dll")]
        private static extern short GetKeyState(VirtualKeyStates nVirtKey);

        private Int32 wparam;
        private Int32 lparam;

        /// <summary>
        /// 线程键盘钩子事件参数，注意与全局键盘钩子完全不同！
        /// </summary>
        /// <param name="wParam">virtual-key code</param>
        /// <param name="lParam">The repeat count, scan code, extended-key flag, context code, previous key-state flag, and transition-state flag.</param>
        internal KeyboardHookEventArgs(Int32 wParam, Int32 lParam)
        {
            // KeyboardProc callback function
            // https://docs.microsoft.com/en-us/previous-versions/windows/desktop/legacy/ms644984(v=vs.85)
            // wParam: Virtual-Key Codes
            // https://docs.microsoft.com/zh-cn/windows/win32/inputdev/virtual-key-codes?redirectedfrom=MSDN
            // lParam: About Keyboard Input
            // https://docs.microsoft.com/zh-cn/windows/win32/inputdev/about-keyboard-input?redirectedfrom=MSDN

            this.wparam = wParam;
            this.lparam = lParam;

            //Control.ModifierKeys doesn't capture alt/win, and doesn't have r/l granularity
            this.isLAltPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_LALT) & KEY_PRESSED);
            this.isRAltPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_RALT) & KEY_PRESSED);

            this.isLCtrlPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_LCONTROL) & KEY_PRESSED);
            this.isRCtrlPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_RCONTROL) & KEY_PRESSED);

            this.isLShiftPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_LSHIFT) & KEY_PRESSED);
            this.isRShiftPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_RSHIFT) & KEY_PRESSED);

            this.isLWinPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_LWIN) & KEY_PRESSED);
            this.isRWinPressed = Convert.ToBoolean(GetKeyState(VirtualKeyStates.VK_RWIN) & KEY_PRESSED);
        }

        public bool isAltPressed { get { return isLAltPressed || isRAltPressed; } }
        public bool isCtrlPressed { get { return isLCtrlPressed || isRCtrlPressed; } }
        public bool isLAltPressed { get; private set; }
        public bool isLCtrlPressed { get; private set; }
        public bool isLShiftPressed { get; private set; }
        public bool isLWinPressed { get; private set; }
        public bool isRAltPressed { get; private set; }
        public bool isRCtrlPressed { get; private set; }
        public bool isRShiftPressed { get; private set; }
        public bool isRWinPressed { get; private set; }
        public bool isShiftPressed { get { return isLShiftPressed || isRShiftPressed; } }
        public bool isWinPressed { get { return isLWinPressed || isRWinPressed; } }

        public override string ToString()
        {
            //return string.Format("Key={0}; Win={1}; Alt={2}; Ctrl={3}; Shift={4}", new object[] { Key, isWinPressed, isAltPressed, isCtrlPressed, isShiftPressed });
            if ((wparam >= 0x30 && wparam <= 0x39) || (wparam >= 0x41 && wparam <= 0x5A))
            {
                return new string((char)wparam, 1);
            }
            return "";
        }

        public string GetKeystrokeMessageFlag()
        {
            return Convert.ToString(lparam, 2);
        }
    }
}