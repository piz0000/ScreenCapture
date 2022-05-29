using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace ScreenCapture
{
    /// <summary>
    /// 화면 저장 <para></para>
    /// Main()의 Application.SetCompatibleTextRenderingDefault(); 이후에 생성해야 한다.
    /// </summary>
    public partial class FormScreenCapture : Form
    {
        //핫키등록
        [DllImport("user32.dll")] static extern int RegisterHotKey(IntPtr hwnd, int id, KeyModifiers fsModifiers, Keys vk);

        //핫키제거
        [DllImport("user32.dll")] static extern int UnregisterHotKey(IntPtr hwnd, int id);

        enum KeyModifiers { None = 0, Alt = 1, Control = 2, Shift = 4, Windows = 8 }
        const int WM_HOTKEY = 0x0312;   //WndProc Message.Msg
        const int HOTKEY_ID1 = 31197;   //Any number to use to [ctrl + shift + s]
        const int HOTKEY_ID2 = 31198;   //Any number to use to [ctrl + shift + a]
        const int HOTKEY_ID3 = 31199;   //Any number to use to [ctrl + shift + q]

        //윈도우 화면 좌표 값으로 위치 정보 반환한다.
        [DllImport("user32.dll")] static extern bool GetWindowRect(IntPtr hWnd, ref Rect rectangle);
        //윈도우 클라이언트 영역에서의 좌표값 반환한다.
        [DllImport("user32.dll")] static extern bool GetClientRect(IntPtr hWnd, out Rect rect);
        struct Rect
        {
            public int Left { get; set; }
            public int Top { get; set; }
            public int Right { get; set; }
            public int Bottom { get; set; }
        }

        //현재 활성창 가져오기
        [DllImport("user32.dll")] static extern IntPtr GetForegroundWindow();

        //윈도우 비스타 이상은 이게 정확하다
        //윈도우 비스타 이상에서는 화면에 그릴때 그림자와 화면 이동 등
        //여러가지 원인때문에 좀더 크거나 작게 나오는데
        //윈도우 차이 때문에 GetWindowRect가 제대로 작동하지 않는 것처럼 보인다
        [DllImport("dwmapi.dll")] static extern int DwmGetWindowAttribute(IntPtr hwnd, DWMWINDOWATTRIBUTE dwAttribute, out RECT pvAttribute, int cbAttribute);
        enum DWMWINDOWATTRIBUTE : uint
        {
            NCRenderingEnabled = 1,
            NCRenderingPolicy,
            TransitionsForceDisabled,
            AllowNCPaint,
            CaptionButtonBounds,
            NonClientRtlLayout,
            ForceIconicRepresentation,
            Flip3DPolicy,
            ExtendedFrameBounds,
            HasIconicBitmap,
            DisallowPeek,
            ExcludedFromPeek,
            Cloak,
            Cloaked,
            FreezeRepresentation
        }


        #region 현재 실행중인 프로그램의 좌표값 얻기
        [DllImport("user32.dll", SetLastError = true)] [return: MarshalAs(UnmanagedType.Bool)] static extern bool GetWindowPlacement(IntPtr hWnd, ref WINDOWPLACEMENT lpwndpl);
        [Serializable]
        [StructLayout(LayoutKind.Sequential)]
        struct WINDOWPLACEMENT
        {
            public int length;
            public int flags;
            public ShowWindowCommands showCmd;
            public Point ptMinPosition;
            public Point ptMaxPosition;
            public Rectangle rcNormalPosition;
        }
        enum ShowWindowCommands : int
        {
            Hide = 0,
            Normal = 1,
            Minimized = 2,
            Maximized = 3,
        }
        WINDOWPLACEMENT GetPlacement(IntPtr hwnd)
        {
            WINDOWPLACEMENT placement = new WINDOWPLACEMENT();
            placement.length = Marshal.SizeOf(placement);
            GetWindowPlacement(hwnd, ref placement);
            return placement;
        }
        #endregion

        //응용 프로그램 이름 가져오기
        [DllImport("user32.dll")] static extern int GetWindowText(int hWnd, StringBuilder text, int count);



        /// <summary>
        /// 핫키 등록 시작
        /// </summary>
        public FormScreenCapture()
        {
            InitializeComponent();

            RegisterHotKey(Handle, HOTKEY_ID1, KeyModifiers.Control | KeyModifiers.Shift, Keys.S);
            RegisterHotKey(Handle, HOTKEY_ID2, KeyModifiers.Control | KeyModifiers.Shift, Keys.A);
            RegisterHotKey(Handle, HOTKEY_ID3, KeyModifiers.Control | KeyModifiers.Shift, Keys.Q);
        }

        /// <summary>
        /// 핫키 해제
        /// </summary>
        ~FormScreenCapture()
        {
            UnregisterHotKey(Handle, HOTKEY_ID1);
            UnregisterHotKey(Handle, HOTKEY_ID2);
        }

        protected override void WndProc(ref Message m)
        {
            switch (m.Msg)
            {
                case WM_HOTKEY:
                    Keys _keys = (Keys)(((int)m.LParam >> 16) & 0xFFFF); //눌러진 단축키의 키
                    KeyModifiers _keyModifiers = (KeyModifiers)((int)m.LParam & 0xFFFF);   //눌러진 단축키의 수식어
                    if (6 == (int)_keyModifiers && Keys.S == _keys)
                    {
                        // 활성 저장
                        IntPtr hwnd = GetForegroundWindow();
                        SavePng(GetProcessWindowText(hwnd), GetRectangleToImage(GetRectangle(hwnd)));
                    }
                    else if (6 == (int)_keyModifiers && Keys.A == _keys)
                    {
                        // 화면 저장
                        SavePng(GetRectangleToImage(Screen.FromPoint(Cursor.Position).WorkingArea));
                    }
                    else if (6 == (int)_keyModifiers && Keys.Q == _keys)
                    {
                        Rectangle rect = new Rectangle();
                        rect.X = (int)System.Windows.SystemParameters.VirtualScreenLeft;
                        rect.Y = (int)System.Windows.SystemParameters.VirtualScreenTop;
                        rect.Width = (int)System.Windows.SystemParameters.VirtualScreenWidth;
                        rect.Height = (int)System.Windows.SystemParameters.VirtualScreenHeight;
                        SavePng(GetRectangleToImage(rect));
                    }
                    break;
            }

            base.WndProc(ref m);
        }

        /// <summary>
        /// Window Text 가져오기
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        string GetProcessWindowText(IntPtr hwnd)
        {
            StringBuilder Buff = new StringBuilder(256);

            if (GetWindowText(hwnd.ToInt32(), Buff, 256) > 0)
            {
                return Buff.ToString();
            }
            else
            {
                return string.Empty;
            }
        }

        /// <summary>
        /// 화면 영역 가져오기
        /// </summary>
        /// <param name="rect"></param>
        /// <returns></returns>
        Bitmap GetRectangleToImage(Rectangle rect)
        {
            if (rect.IsEmpty)
            {
                return null;
            }

            Bitmap bmp = new Bitmap(rect.Width, rect.Height, PixelFormat.Format32bppArgb);
            Graphics graphics = Graphics.FromImage(bmp);
            graphics.CopyFromScreen(rect.Left, rect.Top, 0, 0, rect.Size, CopyPixelOperation.SourceCopy);
            return bmp;
        }

        /// <summary>
        /// Png 저장하기
        /// </summary>
        /// <param name="fileName"></param>
        /// <param name="bitmap"></param>
        void SavePng(string fileName, Bitmap bitmap)
        {
            if (bitmap == null)
            {
                return;
            }

            string folder = Application.StartupPath + "\\ScreenCapture\\";
            Directory.CreateDirectory(folder);

            if (string.IsNullOrEmpty(fileName))
            {
                fileName = "ScreenCapture";
            }
            fileName = InvalidFileName(fileName);
            fileName = DuplicationFileName(folder, fileName);

            try
            {
                bitmap.Save(folder + fileName + ".png", ImageFormat.Png);
            }
            catch (Exception exe)
            {
                MessageBox.Show(exe.ToString());
            }
        }

        /// <summary>
        /// Png 저장하기
        /// </summary>
        /// <param name="bitmap"></param>
        void SavePng(Bitmap bitmap)
        {
            SavePng("ScreenCapture", bitmap);
        }

        /// <summary>
        /// 파일 이름으로 사용할수 없는 문자열 제거
        /// </summary>
        /// <param name="fileName"></param>
        /// <returns></returns>
        string InvalidFileName(string fileName)
        {
            string invalidFileNameChars = new string(Path.GetInvalidFileNameChars());
            string invalidPathChars = new string(Path.GetInvalidPathChars());

            string regexSearch = invalidFileNameChars + invalidPathChars;
            Regex regex = new Regex(string.Format("[{0}]", Regex.Escape(regexSearch)));
            string _fileName = regex.Replace(fileName, "");

            if (_fileName.Length >= 260)
            {
                _fileName = _fileName.Substring(0, 259);
            }

            return _fileName;
        }

        /// <summary>
        /// 파일 이름 중복되면 fileName 뒤에 (숫자) 붙인다
        /// </summary>
        /// <param name="folder"></param>
        /// <param name="fileName"></param>
        /// <returns></returns>
        string DuplicationFileName(string folder, string fileName)
        {
            //존재시 괄호(숫자) 붙임
            if (File.Exists(folder + fileName + ".png"))
            {
                string[] count = Directory.GetFiles(folder, fileName + "*", SearchOption.TopDirectoryOnly);
                if (count != null && count.Length >= 1)
                {
                    fileName = fileName + " (" + count.Length + ")";
                }
            }
            return fileName;
        }

        /// <summary>
        /// Window 영역 가져오기
        /// </summary>
        /// <param name="hwnd"></param>
        /// <returns></returns>
        Rectangle GetRectangle(IntPtr hwnd)
        {
            //화면 상태 알아내기
            WINDOWPLACEMENT windowPlacement = GetPlacement(hwnd);
            if (windowPlacement.showCmd == ShowWindowCommands.Maximized)
            {
                //전체 저장
                SavePng(GetRectangleToImage(Screen.PrimaryScreen.WorkingArea));
            }
            else if (windowPlacement.showCmd == ShowWindowCommands.Normal)
            {
                Rectangle convertRectangle = new Rectangle();
                if (Environment.OSVersion.Version.Major < 6)
                {
                    //윈도우 비스타 이전...

                    //clientRect.Right [가로] clientRect.Bottom [세로] 윈도우(폼) 사이즈는 이게 정확하다
                    GetClientRect(hwnd, out Rect clientRect);
                    Size winSize = new Size(clientRect.Right, clientRect.Bottom);

                    Rect rect = new Rect();
                    bool result = GetWindowRect(hwnd, ref rect);

                    if (result)
                    {
                        convertRectangle = ConvertRectToRectangle(rect);

                        Point point = new Point();
                        if (rect.Left == -7)
                        {
                            //윈도우 자석기능 사용시 -7 나오니 0으로 변경, 우측은 괜찮음
                            point.X = 0;
                        }
                        else
                        {
                            //위치 보정
                            point.X = convertRectangle.X + ((convertRectangle.Width - winSize.Width) / 2);
                        }
                        point.Y = convertRectangle.Y;

                        convertRectangle.Location = point;
                        convertRectangle.Size = new Size(winSize.Width, convertRectangle.Height);
                    }
                }
                else
                {
                    //윈도우 비스타 이후..
                    RECT rcFrame;
                    int result = DwmGetWindowAttribute(hwnd, DWMWINDOWATTRIBUTE.ExtendedFrameBounds, out rcFrame, Marshal.SizeOf(typeof(RECT)));
                    if (result != 0)
                    {
                        //Handle for failure
                    }
                    else
                    {
                        convertRectangle.Location = rcFrame.Location;
                        convertRectangle.Size = rcFrame.Size;
                    }
                }

                return convertRectangle;
            }

            return new Rectangle();
        }

        /// <summary>
        /// Rect To Rectangle
        /// </summary>
        /// <param name="rect"></param>
        /// <returns></returns>
        Rectangle ConvertRectToRectangle(Rect rect)
        {
            Size size = new Size();
            size.Width = rect.Right - rect.Left;
            size.Height = rect.Bottom - rect.Top;

            Point point = new Point();
            point.X = rect.Left;
            point.Y = rect.Top;

            Rectangle rectangle = new Rectangle(point, size);
            return rectangle;
        }



        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            public int Left, Top, Right, Bottom;

            public RECT(int left, int top, int right, int bottom)
            {
                Left = left;
                Top = top;
                Right = right;
                Bottom = bottom;
            }

            public RECT(Rectangle r) : this(r.Left, r.Top, r.Right, r.Bottom) { }

            public int X
            {
                get { return Left; }
                set { Right -= (Left - value); Left = value; }
            }

            public int Y
            {
                get { return Top; }
                set { Bottom -= (Top - value); Top = value; }
            }

            public int Height
            {
                get { return Bottom - Top; }
                set { Bottom = value + Top; }
            }

            public int Width
            {
                get { return Right - Left; }
                set { Right = value + Left; }
            }

            public Point Location
            {
                get { return new Point(Left, Top); }
                set { X = value.X; Y = value.Y; }
            }

            public Size Size
            {
                get { return new Size(Width, Height); }
                set { Width = value.Width; Height = value.Height; }
            }

            public static implicit operator Rectangle(RECT r)
            {
                return new Rectangle(r.Left, r.Top, r.Width, r.Height);
            }

            public static implicit operator RECT(System.Drawing.Rectangle r)
            {
                return new RECT(r);
            }

            public static bool operator ==(RECT r1, RECT r2)
            {
                return r1.Equals(r2);
            }

            public static bool operator !=(RECT r1, RECT r2)
            {
                return !r1.Equals(r2);
            }

            public bool Equals(RECT r)
            {
                return r.Left == Left && r.Top == Top && r.Right == Right && r.Bottom == Bottom;
            }

            public override bool Equals(object obj)
            {
                if (obj is RECT)
                    return Equals((RECT)obj);
                else if (obj is Rectangle)
                    return Equals(new RECT((Rectangle)obj));
                return false;
            }

            public override int GetHashCode()
            {
                return ((Rectangle)this).GetHashCode();
            }

            public override string ToString()
            {
                return string.Format(System.Globalization.CultureInfo.CurrentCulture, "{{Left={0},Top={1},Right={2},Bottom={3}}}", Left, Top, Right, Bottom);
            }
        }
    }
}
