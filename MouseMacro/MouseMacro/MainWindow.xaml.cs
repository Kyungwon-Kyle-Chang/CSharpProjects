using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace MouseMacro
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static System.Timers.Timer aTimer;
        private int _mouseClickNum = 1;
        private float _mouseClickDelay = 0.1f;
        private int _coordListIndex = 0;

        public MainWindow()
        {
            InitializeComponent();

            MouseHook.Start();
        }


        private void MouseHook_MouseLClick(object sender, EventArgs e)
        {
            if (StopButton.IsMouseOver) return;

            Mouse.Capture(this);
            Point pointToWindow = Mouse.GetPosition(this);
            Point pointToScreen = PointToScreen(pointToWindow);
            Mouse.Capture(null);

            List<MouseCoordInfo> items = new List<MouseCoordInfo>();

            items.Add(new MouseCoordInfo() { Num = _mouseClickNum++, X = pointToScreen.X, Y = pointToScreen.Y });
            SavedPointsView.Items.Add(items);
        }

        private void Event(object sender, EventArgs e)
        {
            Console.WriteLine("Left mouse click!");

            Mouse.Capture(this);
            Point pointToWindow = Mouse.GetPosition(this);
            Point pointToScreen = PointToScreen(pointToWindow);
            Mouse.Capture(null);
        }

        private void SaveButton_Click(object sender, RoutedEventArgs e)
        {
            MouseHook.MouseLClick += new EventHandler(MouseHook_MouseLClick);
            SavedPointsView.Items.Clear();
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            MouseHook.MouseLClick -= MouseHook_MouseLClick;
            _mouseClickNum = 1;
        }

        private void ExecuteButton_Click(object sender, RoutedEventArgs e)
        {
            _coordListIndex = 0;

            aTimer = new System.Timers.Timer(_mouseClickDelay * 1000);
            aTimer.Elapsed += timerFunction;
            aTimer.Enabled = true;
        }

        private void Window_KeyDown(object sender, KeyEventArgs e)
        {
            //if(e.Key == Key.E)
            //{
            //    _coordListIndex = 0;

            //    aTimer = new System.Timers.Timer(_mouseClickDelay * 1000);
            //    aTimer.Elapsed += timerFunction;
            //    aTimer.Enabled = true;
            //}
        }

        private void textBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            string sDelay = textBox.Text;

            float.TryParse(sDelay, out _mouseClickDelay);
        }

        private void timerFunction(object sender, System.Timers.ElapsedEventArgs e)
        {
            if (_coordListIndex >= SavedPointsView.Items.Count)
            {
                aTimer.Close();
                aTimer.Enabled = false;
                aTimer.Elapsed -= timerFunction; 
                return;
            }
            else
            {
                int index = _coordListIndex++;
                int x = (int)((List<MouseCoordInfo>)SavedPointsView.Items[index])[0].X;
                int y = (int)(((List<MouseCoordInfo>)SavedPointsView.Items[index]))[0].Y;
                SetCursorPos(x, y);
                mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)x, (uint)y, 0, 0);
            }
        }

        public class MouseCoordInfo
        {
            public int Num { get; set; }
            public double X { get; set; }
            public double Y { get; set; }
        }

        public static class MouseHook
        {
            public static event EventHandler MouseMove = delegate { };
            public static event EventHandler MouseLClick = delegate { };

            public static void Start()
            {
                _hookID = SetHook(_proc);
            }

            public static void stop()
            {
                UnhookWindowsHookEx(_hookID);
            }

            private static LowLevelMouseProc _proc = HookCallback;
            private static IntPtr _hookID = IntPtr.Zero;

            private static IntPtr SetHook(LowLevelMouseProc proc)
            {
                using (Process curProcess = Process.GetCurrentProcess())
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return SetWindowsHookEx(WH_MOUSE_LL, proc,
                        GetModuleHandle(curModule.ModuleName), 0);
                }
            }

            private delegate IntPtr LowLevelMouseProc(int nCode, IntPtr wParam, IntPtr lParam);

            private static IntPtr HookCallback(
                int nCode, IntPtr wParam, IntPtr lParam)
            {
                if (nCode >= 0 && MouseMessages.WM_MOUSEMOVE == (MouseMessages)wParam)
                {
                    MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    MouseMove(null, new EventArgs());
                }
                else if (nCode >= 0 && MouseMessages.WM_LBUTTONDOWN == (MouseMessages)wParam)
                {
                    MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    MouseLClick(null, new EventArgs());
                }
                return CallNextHookEx(_hookID, nCode, wParam, lParam);
            }

            private const int WH_MOUSE_LL = 14;

            private enum MouseMessages
            {
                WM_LBUTTONDOWN = 0x0201,
                WM_LBUTTONUP = 0x0202,
                WM_MOUSEMOVE = 0x0200,
                WM_MOUSEWHEEL = 0x020A,
                WM_RBUTTONDOWN = 0x0204,
                WM_RBUTTONUP = 0x0205
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct POINT
            {
                public int x;
                public int y;
            }

            [StructLayout(LayoutKind.Sequential)]
            private struct MSLLHOOKSTRUCT
            {
                public POINT pt;
                public uint mouseData;
                public uint flags;
                public uint time;
                public IntPtr dwExtraInfo;
            }

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr SetWindowsHookEx(int idHook,
                LowLevelMouseProc lpfn, IntPtr hMod, uint dwThreadId);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            [return: MarshalAs(UnmanagedType.Bool)]
            private static extern bool UnhookWindowsHookEx(IntPtr hhk);

            [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr CallNextHookEx(IntPtr hhk, int nCode,
                IntPtr wParam, IntPtr lParam);

            [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
            private static extern IntPtr GetModuleHandle(string lpModuleName);
        }

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        [DllImport("User32", EntryPoint = "SetCursorPos")]
        public static extern void SetCursorPos(int x, int y);
    }
}
