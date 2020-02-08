using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace Hook
{
    public class KeyboardHook
    {
        /// <summary>
        /// 声明钩子委托
        /// </summary>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        public delegate int HookProc(int nCode, Int32 wParam, IntPtr lParam);

        /// <summary>
        /// 装载钩子
        /// </summary>
        /// <param name="idHook"></param>
        /// <param name="lpfn"></param>
        /// <param name="hInstance"></param>
        /// <param name="threadId"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);

        /// <summary>
        /// 卸载钩子
        /// </summary>
        /// <param name="idHook"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern bool UnhookWindowsHookEx(int idHook);

        /// <summary>
        /// 传递钩子
        /// </summary>
        /// <param name="idHook"></param>
        /// <param name="nCode"></param>
        /// <param name="wParam"></param>
        /// <param name="lParam"></param>
        /// <returns></returns>
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern int CallNextHookEx(int idHook, int nCode, Int32 wParam, IntPtr lParam);

        /// <summary>
        /// 获取全部按键状态
        /// </summary>
        /// <param name="pbKeyState"></param>
        /// <returns>非0表示成功</returns>
        [DllImport("user32.dll")]
        public static extern int GetKeyboardState(byte[] pbKeyState);

        /// <summary>
        /// 错误代码
        /// </summary>
        /// <returns></returns>
        [DllImport("Kernel32.dll")]
        public static extern int GetLastError();

        /// <summary>
        /// 获取程序集模块句柄
        /// </summary>
        /// <param name="lpModuleName"></param>
        /// <returns></returns>
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern IntPtr GetModuleHandle(string lpModuleName);

        /// <summary>
        /// 获取线程ID
        /// </summary>
        /// <returns></returns>
        [DllImport("kernel32.dll")]
        public static extern int GetCurrentThreadId();

        #region 变量
        private HookProc KeyboardHookProcedure; // 键盘钩子委托实例
        private static int hKeyboardHook = 0;   // 键盘钩子句柄
        public const int WH_KEYBOARD = 2;       // 键盘消息钩子(线程钩子)
        #endregion

        /// <summary>
        /// 事件
        /// </summary>
        public event KeyEventHandler OnKeyDownEvent;

        public KeyboardHook()
        {
            Start();
        }

        ~KeyboardHook()
        {
            Stop();
        }

        /// <summary>
        /// 安装钩子
        /// </summary>
        public void Start()
        {
            //安装一个线程键盘钩子
            if (hKeyboardHook == 0)
            {
                KeyboardHookProcedure = new HookProc(KeyboardHookProc);
                hKeyboardHook = SetWindowsHookEx(WH_KEYBOARD, KeyboardHookProcedure,
                    IntPtr.Zero, GetCurrentThreadId());
                if (hKeyboardHook == 0)
                {
                    Stop();
                    Console.WriteLine(GetLastError());
                    throw new Exception("SetWindowsHookEx ist failed.");
                }
            }
        }

        /// <summary>
        /// 卸载钩子
        /// </summary>
        public void Stop()
        {
            bool retKeyboard = true;
            if (hKeyboardHook != 0)
            {
                retKeyboard = UnhookWindowsHookEx(hKeyboardHook);
                hKeyboardHook = 0;
            }
            //如果卸下钩子失败
            if (!(retKeyboard))
            {
                Console.WriteLine(GetLastError());
                throw new Exception("UnhookWindowsHookEx failed.");
            }
        }

        private int KeyboardHookProc(int nCode, Int32 wParam, IntPtr lParam)
        {
            // lParam的第31位为0表示按下，为1表示松开
            try
            {
                if (nCode >= 0)
                {
                    Keys keyData = (Keys)wParam;
                    Int32 lparam = 0;
                    Marshal.PtrToStructure(lParam, lparam);
                    MessageBox.Show(nCode.ToString() + ":" + keyData.ToString() + "," + lparam.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

            return CallNextHookEx(hKeyboardHook, nCode, wParam, lParam);
        }
    }
}