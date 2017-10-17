using System;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.Media;
using System.Runtime.InteropServices;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace ColorDetectMacro
{
    /// <summary>
    /// MainWindow.xaml에 대한 상호 작용 논리
    /// </summary>
    public partial class MainWindow : Window
    {
        private TextBox[] _textBoxArray = new TextBox[8];
        private int _textBoxArrayIndex = 0;
        private float msDelay = 1000f;

        private int refreshX, refreshY;
        private int vipStartX, vipStartY;
        private int vipEndX, vipEndY;
        private int proceedX, proceedY;
        private int targetX, targetY;
        private readonly int[] vipRGB = new int[3] { 149, 42, 224 };
        private Bitmap screenPixel;
        DispatcherTimer dispatcherTimer = new DispatcherTimer();

        public MainWindow()
        {
            InitializeComponent();

            _textBoxArray[0] = RefreshX;
            _textBoxArray[1] = RefreshY;
            _textBoxArray[2] = VIP_StartX;
            _textBoxArray[3] = VIP_StartY;
            _textBoxArray[4] = VIP_EndX;
            _textBoxArray[5] = VIP_EndY;
            _textBoxArray[6] = ProceedX;
            _textBoxArray[7] = ProceedY;

            dispatcherTimer.Tick += dispatcherTimer_Tick;
            MouseHook.Start();
        }

        private void MouseHook_MouseLClick(object sender, EventArgs e)
        {
            Win32Point winPoint = new Win32Point();
            GetCursorPos(ref winPoint);

            int x = winPoint.X;
            int y = winPoint.Y;
            _textBoxArray[_textBoxArrayIndex].Text = x.ToString();
            _textBoxArray[_textBoxArrayIndex + 1].Text = y.ToString();

            _textBoxArrayIndex += 2;
            if (_textBoxArrayIndex >= 8)
            {
                MouseHook.MouseLClick -= MouseHook_MouseLClick;
                _textBoxArrayIndex = 0;
            }
        }

        private void MouseHook_MouseRClick(object sender, EventArgs e)
        {
            FinishTimer();
        }

        private void dispatcherTimer_Tick(object sender, EventArgs e)
        {
            Stopwatch sw = new Stopwatch();
            sw.Start();
            
            SetCursorPos(refreshX, refreshY);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)refreshX, (uint)refreshY, 0, 0);

            POINT point;
            if (!CheckColor(out point))
            {
                sw.Stop();
                Elapsed.Text = sw.ElapsedMilliseconds.ToString();
                return;
            }

            SetCursorPos(vipStartX + point.X, vipStartY + point.Y);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)point.X, (uint)point.Y, 0, 0);

            SetCursorPos(proceedX, proceedY);
            mouse_event(MOUSEEVENTF_LEFTDOWN | MOUSEEVENTF_LEFTUP, (uint)proceedX, (uint)proceedY, 0, 0);

            sw.Stop();
            Elapsed.Text = sw.ElapsedMilliseconds.ToString();

            if((bool)AlarmCheckBox.IsChecked) SystemSounds.Beep.Play();
            FinishTimer();
        }

        private bool CheckColor(out POINT point)
        {
            GetColorAt(vipStartX, vipStartY);

            for (int h = 0; h < vipEndY - vipStartY; h++)
            {
                for (int w = 0; w < vipEndX - vipStartX; w++) 
                {
                    Color pixelColor = screenPixel.GetPixel(w, h);
                    if(CompareColor(pixelColor))
                    {
                        point = new POINT(w, h);
                        return true;
                    }
                }
            }

            point = new POINT();
            return false;
        }

        private bool CompareColor(Color pixelColor)
        {
            return (pixelColor.R < vipRGB[0] + 10 && pixelColor.R > vipRGB[0] - 10) &&
                    (pixelColor.G < vipRGB[1] + 10 && pixelColor.G > vipRGB[1] - 10) &&
                    (pixelColor.B < vipRGB[2] + 10 && pixelColor.B > vipRGB[2] - 10);
        }

        private void FinishTimer()
        {
            dispatcherTimer.Stop();

            MouseHook.MouseRClick -= MouseHook_MouseRClick;
        }
        //------------------------------------ UI Functions ----------------------------------
        private void SettingButton_Click(object sender, RoutedEventArgs e)
        {
            for (int i = 0; i < _textBoxArray.Length; ++i)
                _textBoxArray[i].Clear();

            MouseHook.MouseLClick += new EventHandler(MouseHook_MouseLClick);
        }

        private void PlayButton_Click(object sender, RoutedEventArgs e)
        {
            refreshX = Convert.ToInt32(_textBoxArray[0].Text);
            refreshY = Convert.ToInt32(_textBoxArray[1].Text);
            vipStartX = Convert.ToInt32(_textBoxArray[2].Text);
            vipStartY = Convert.ToInt32(_textBoxArray[3].Text);
            vipEndX = Convert.ToInt32(_textBoxArray[4].Text);
            vipEndY = Convert.ToInt32(_textBoxArray[5].Text);
            proceedX = Convert.ToInt32(_textBoxArray[6].Text);
            proceedY = Convert.ToInt32(_textBoxArray[7].Text);
            targetX = Convert.ToInt32(TargetRow.Text);
            targetY = Convert.ToInt32(TargetCol.Text);

            screenPixel = new Bitmap(vipEndX - vipStartX, vipEndY - vipStartY, PixelFormat.Format32bppArgb);

            MouseHook.MouseRClick += new EventHandler(MouseHook_MouseRClick);
            
            dispatcherTimer.Interval = new TimeSpan(0, 0, 0, 0, (int)msDelay);
            dispatcherTimer.Start();
        }

        private void Delay_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (float.TryParse(Delay.Text, out msDelay)) msDelay *= 1000;
        }

        private void StopButton_Click(object sender, RoutedEventArgs e)
        {
            FinishTimer();
        }
        //-------------------------------------------------------------------------------------
        #region MouseHook Class
        public static class MouseHook
        {
            public static event EventHandler MouseMove = delegate { };
            public static event EventHandler MouseLClick = delegate { };
            public static event EventHandler MouseRClick = delegate { };

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
                else if (nCode >= 0 && MouseMessages.WM_RBUTTONDOWN == (MouseMessages)wParam)
                {
                    MSLLHOOKSTRUCT hookStruct = (MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(MSLLHOOKSTRUCT));
                    MouseRClick(null, new EventArgs());
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
        #endregion

        #region Native Methods
        [StructLayout(LayoutKind.Sequential)]
        internal struct Win32Point
        {
            public Int32 X;
            public Int32 Y;
        };

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetCursorPos(ref Win32Point pt);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        [DllImport("User32", EntryPoint = "SetCursorPos")]
        public static extern void SetCursorPos(int x, int y);


        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        public struct POINT
        {
            public int X;
            public int Y;

            public POINT(int x, int y)
            {
                X = x;
                Y = y;
            }
        }

        [DllImport("gdi32.dll", CharSet = CharSet.Auto, SetLastError = true, ExactSpelling = true)]
        public static extern int BitBlt(IntPtr hDC, int x, int y, int nWidth, int nHeight, IntPtr hSrcDC, int xSrc, int ySrc, int dwRop);
        
        public void GetColorAt(int x, int y)
        {
            using (Graphics gdest = Graphics.FromImage(screenPixel))
            {
                using (Graphics gsrc = Graphics.FromHwnd(IntPtr.Zero))
                {
                    IntPtr hSrcDC = gsrc.GetHdc();
                    IntPtr hDC = gdest.GetHdc();
                    int retval = BitBlt(hDC, 0, 0, screenPixel.Width, screenPixel.Height, hSrcDC, x, y, (int)CopyPixelOperation.SourceCopy);
                    gdest.ReleaseHdc();
                    gsrc.ReleaseHdc();
                }
            }
        }
        #endregion
    }
    

}