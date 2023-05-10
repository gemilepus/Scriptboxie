using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Media;
using System.Windows.Forms;
using System.Drawing;
using System.Collections;
using System.Runtime.InteropServices;
using System.Threading;
using System.Diagnostics;

using OpenCvSharp;
using OpenCvSharp.Extensions;

using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

using GameOverlay.Drawing;
using GameOverlay.Windows;

using Gma.System.MouseKeyHook;

using IniParser;
using IniParser.Model;

using WindowsInput;
using WindowsInput.Native;
using static Keyboard;
using System.Windows.Navigation;
using System.Security.Principal;
using System.Management;
using System.Windows.Media;
using System.Net;
using System.Text.Json;
using System.Windows.Interop;

namespace Metro
{
    public partial class MainWindow : MetroWindow
    {
        // **************************************** Motion ******************************************
        #region Motion
        // Mouse move
        [DllImport("user32")]
        private static extern int SetCursorPos(int x, int y);
        [DllImport("user32")]
        private static extern bool GetCursorPos(ref System.Drawing.Point lpPoint);
        // VkKeyScan Char to 0x00
        [DllImport("user32.dll")]
        private static extern byte VkKeyScan(char ch);
        #endregion
        // **************************************** Window ******************************************
        #region Window

        // GetActiveWindowTitle
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        private string GetActiveWindowTitle()
        {
            const int nChars = 256;
            StringBuilder Buff = new StringBuilder(nChars);
            IntPtr handle = GetForegroundWindow();

            if (GetWindowText(handle, Buff, nChars) > 0)
            {
                return Buff.ToString();
            }
            return null;
        }

        [DllImport("USER32.DLL")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);

        [DllImport("user32.dll", SetLastError = true)]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        // PostMessageA
        [DllImport("User32.Dll", EntryPoint = "PostMessageA")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        private const int WM_KEYDOWN = 0x100;
        private const int WM_KEYUP = 0x101;
        private const int WM_CHAR = 0x0102;

        //const int WM_LBUTTONDOWN = 0x201;
        //const int WM_LBUTTONUP = 0x202;

        // SendMessage
        [DllImport("user32.dll")]
        private static extern int SendMessage(int hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPStr)] string lParam);

        #endregion
        // **************************************** Image ******************************************
        #region Image
        private Bitmap makeScreenshot()
        {
            Bitmap screenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);

            gfxScreenshot.Dispose();

            return screenshot;
        }

        private Bitmap makeScreenshot_clip(int x, int y, int height, int width)
        {
            Bitmap screenshot = new Bitmap(width, height);

            System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(x, y, 0, 0, screenshot.Size);

            gfxScreenshot.Dispose();

            return screenshot;
        }

        private String RunTemplateMatch(Mat rec, Mat template,string Mode,double mThreshold)
        {
            string ResponseStr = "";

            using (Mat refMat = rec)
            using (Mat tplMat = template)
            using (Mat res = new Mat(refMat.Rows - tplMat.Rows + 1, refMat.Cols - tplMat.Cols + 1, MatType.CV_32FC1))
            {
                Mat gref, gtpl;
                if (Mode.Equals("RGB"))
                {
                    gref = refMat.CvtColor(ColorConversionCodes.BGR2HLS);
                    gtpl = tplMat.CvtColor(ColorConversionCodes.BGR2HLS);
                }
                else
                {
                    gref = refMat.CvtColor(ColorConversionCodes.BGR2GRAY);
                    gtpl = tplMat.CvtColor(ColorConversionCodes.BGR2GRAY);
                }

                Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.CCoeffNormed);
                //Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.SqDiffNormed);
                Cv2.Threshold(res, res, 0.8, 1.0, ThresholdTypes.Tozero);

                while (true)
                {
                    double minval, maxval, threshold = mThreshold;
                    OpenCvSharp.Point minloc, maxloc;
                    Cv2.MinMaxLoc(res, out minval, out maxval, out minloc, out maxloc);

                    if (maxval >= threshold)
                    {
                        //Setup the rectangle to draw
                        //OpenCvSharp.Rect r = new OpenCvSharp.Rect(new OpenCvSharp.Point(maxloc.X, maxloc.Y), new OpenCvSharp.Size(tplMat.Width, tplMat.Height));

                        //Draw a rectangle of the matching area
                        //Cv2.Rectangle(refMat, r, Scalar.LimeGreen, 2);

                        //Fill in the res Mat so you don't find the same area again in the MinMaxLoc
                        OpenCvSharp.Rect outRect;
                        Cv2.FloodFill(res, maxloc, new Scalar(0), out outRect, new Scalar(0.1), new Scalar(1.0), FloodFillFlags.Link4);

                        ResponseStr = ResponseStr + maxloc.X.ToString() + "," + maxloc.Y.ToString() + ",";
                    }
                    else
                    {
                        break;
                    }

                    gref.Dispose();
                    gtpl.Dispose();
                }

                refMat.Dispose();
                tplMat.Dispose();

                gref.Dispose();
                gtpl.Dispose();

                res.Dispose();

                rec.Dispose();
                template.Dispose();

                return ResponseStr;
            }
        }

        private Mat DetectFace_Mat(CascadeClassifier cascade, Mat src) // input Mat 
        {
            Mat result;
            using (var gray = new Mat())
            {
                result = src.Clone();
                Cv2.CvtColor(src, gray, ColorConversionCodes.BGR2GRAY);
                // Detect faces
                OpenCvSharp.Rect[] faces = cascade.DetectMultiScale(gray, 1.08, 2, HaarDetectionTypes.ScaleImage, new OpenCvSharp.Size(30, 30));
                //Render all detected faces
                foreach (OpenCvSharp.Rect face in faces)
                {
                    var center = new OpenCvSharp.Point // get x,y
                    {
                        X = (int)(face.X + face.Width * 0.5),
                        Y = (int)(face.Y + face.Height * 0.5)
                    };
                    var axes = new OpenCvSharp.Size
                    {
                        Width = (int)(face.Width * 0.5),
                        Height = (int)(face.Height * 0.5)
                    };
                    Cv2.Ellipse(result, center, axes, 0, 0, 360, new Scalar(255, 0, 255), 4);

                }
            }
            return result;
        }

        private void MatchBySift(Mat src1, Mat src2)
        {
            #region MatchBySift
            var gray1 = new Mat();
            var gray2 = new Mat();

            Cv2.CvtColor(src1, gray1, ColorConversionCodes.BGR2GRAY);
            Cv2.CvtColor(src2, gray2, ColorConversionCodes.BGR2GRAY);

            //var sift = SIFT.Create();

            // Detect the keypoints and generate their descriptors using SIFT
            //KeyPoint[] keypoints1, keypoints2;
            //var descriptors1 = new MatOfFloat();
            //var descriptors2 = new MatOfFloat();
            //sift.DetectAndCompute(gray1, null, out keypoints1, descriptors1);
            //sift.DetectAndCompute(gray2, null, out keypoints2, descriptors2);

            //// Match descriptor vectors
            //var bfMatcher = new BFMatcher(NormTypes.L2, false);
            //var flannMatcher = new FlannBasedMatcher();
            //DMatch[] bfMatches = bfMatcher.Match(descriptors1, descriptors2);
            //DMatch[] flannMatches = flannMatcher.Match(descriptors1, descriptors2);

            //// Draw matches
            //var bfView = new Mat();
            //Cv2.DrawMatches(gray1, keypoints1, gray2, keypoints2, bfMatches, bfView);
            //var flannView = new Mat();
            //Cv2.DrawMatches(gray1, keypoints1, gray2, keypoints2, flannMatches, flannView);

            //using (new OpenCvSharp.Window("SIFT matching (by BFMather)", WindowMode.AutoSize, bfView))
            //using (new OpenCvSharp.Window("SIFT matching (by FlannBasedMatcher)", WindowMode.AutoSize, flannView))
            //{
            //    Cv2.WaitKey();
            //}
            #endregion
        }
        #endregion

        private DependencyProperty UnitIsCProperty = DependencyProperty.Register("IsActive", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));
        public static bool IsAdmin => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
        
        private new bool IsActive
        {
            get { return (bool)this.GetValue(UnitIsCProperty); }
            set { this.SetValue(UnitIsCProperty, value); }
        }

        // GameOverlay .Net
        private OverlayWindow _window;
        private GameOverlay.Drawing.Graphics _graphics;
        private GameOverlay.Drawing.SolidBrush _red;
        private GameOverlay.Drawing.Font _font;
        private GameOverlay.Drawing.SolidBrush _black;
        private GameOverlay.Drawing.Graphics gfx;
        // NotifyIcon
        private static NotifyIcon mNotifyIcon = new NotifyIcon();
        // DataGrid
        private List<MainTable> mDataTable = new List<MainTable>();
        private List<EditTable> eDataTable = new List<EditTable>();

        private List<Thread> _workerThreads = new List<Thread>();
        private SoundPlayer mAlertSound = new SoundPlayer(Metro.Properties.Resources.sound);
        private SettingHelper mSettingHelper = new SettingHelper();
        private bool TestMode = false;

        #region Globalmousekeyhook
        private static IKeyboardMouseEvents m_GlobalHook, Main_GlobalHook, Main_GlobalKeyUpHook;

        private int now_x, now_y;

        private void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress += GlobalHookKeyPress;
            m_GlobalHook.MouseMove += HookManager_MouseMove;
        }

       
        private void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            m_GlobalHook.Dispose();
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            if (Btn_Toggle.IsOn == true && Btn_Toggle.IsMouseOver == false)
            {

            }
            Console.WriteLine("KeyPress: \t{0}", e.KeyChar);
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            if (Btn_Toggle.IsOn == true && Btn_Toggle.IsMouseOver == false)
            {
                //if (e.Button.Equals("")) { }
                mDataGrid.DataContext = null;
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "500", Event = "", Note = "" });
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Move", Action = now_x.ToString() + "," + now_y.ToString(), Event = "", Note = "" });
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "500", Event = "", Note = "" });
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Click", Action = e.Button.ToString() + "", Event = "", Note = "" });
                mDataGrid.DataContext = mDataTable;
            }
            Console.WriteLine("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp);

            // ScrollToBottom
            if (mDataGrid.Items.Count > 0)
            {
                var border = VisualTreeHelper.GetChild(mDataGrid, 0) as Decorator;
                if (border != null)
                {
                    var scroll = border.Child as ScrollViewer;
                    if (scroll != null) scroll.ScrollToBottom();
                }
            }

            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        private void HookManager_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            now_x = e.X;
            now_y = e.Y;
        }

        private void Main_GlobalHookKeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            // for Hotkey setting
            if (TextBox_OnOff_Hotkey.IsFocused)
            {
                TextBox_OnOff_Hotkey.Text = e.KeyCode.ToString();
            }
            else if (TextBox_Run_Hotkey.IsFocused)
            {
                TextBox_Run_Hotkey.Text = e.KeyCode.ToString();
            }
            else if (TextBox_Stop_Hotkey.IsFocused)
            {
                TextBox_Stop_Hotkey.Text = e.KeyCode.ToString();
            }

            if (Btn_Toggle.IsOn == true)
            {
                if (!e.KeyCode.ToString().Equals(""))
                {
                    string mKeyCode = e.KeyCode.ToString();

                    // Number
                    if (mKeyCode.Length == 2 && mKeyCode.IndexOf("D") != -1)
                    {
                        mKeyCode = mKeyCode.Replace("D", "");
                    }
                    mKeyCode = ConvertHelper.ConvertKeyCode(mKeyCode);

                    if (mKeyCode.IndexOf("Oem") == -1)
                    {
                        mDataGrid.DataContext = null;
                        
                        if (mKeyCode.Equals("WIN") || mKeyCode.Equals("Apps") || mKeyCode.Equals("SNAPSHOT") || mKeyCode.Equals("Scroll") || mKeyCode.Equals("Pause"))
                        {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "500", Event = "", Note = "" });
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Key", Action = mKeyCode, Event = "", Note = "" });
                        }
                        else {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "500", Event = "", Note = "" });
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "SendKeyDown", Action = mKeyCode.ToUpper(), Event = "", Note = "" });
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "200", Event = "", Note = "" });
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "SendKeyUp", Action = mKeyCode.ToUpper(), Event = "", Note = "" });
                        }
                        
                        mDataGrid.DataContext = mDataTable;

                        // ScrollToBottom
                        if (mDataGrid.Items.Count > 0)
                        {
                            var border = VisualTreeHelper.GetChild(mDataGrid, 0) as Decorator;
                            if (border != null)
                            {
                                var scroll = border.Child as ScrollViewer;
                                if (scroll != null) scroll.ScrollToBottom();
                            }
                        }
                    }
                }
            }

            Console.WriteLine("KeyUp: \t{0}", e.KeyCode);
        }
        #endregion

        public MainWindow()
        {
            InitializeComponent();

            for (int i=0; i< LangSplitButton.Items.Count; i++){
                System.Windows.Controls.Label mLabel = (System.Windows.Controls.Label)LangSplitButton.Items[i];
                if (mLabel.Tag.Equals(mSettingHelper.Language)) {
                    LangSplitButton.SelectedIndex = i;
                }
            }

            MSG_SHOW = RegisterWindowMessage("Scriptboxie Message");

            Timestamp = (double)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds;

            // Combobox List
            List<string> mList = new List<string>() {
                "Move","Offset","Click", "Match","Match RGB","Match&Draw",
                "Key","ModifierKey","SendKeyDown","SendKeyUp","Delay",
                "Jump","Goto","Loop","RemoveEvent","Clear Screen",
                "Run .exe","PlaySound","WriteClipboard"
                //"ScreenClip", "Sift Match", "GetPoint","PostMessage","Color Test","FindWindow",
            };
            mComboBoxColumn.ItemsSource = mList;

            #region Load user.ini

            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");

            // Load Form location
            Left = double.Parse(data["Def"]["x"]);
            Top = double.Parse(data["Def"]["y"]);

            // Load WindowTitle setting
            TextBox_Title.Text = data["Def"]["WindowTitle"];

            // Load ScaleX, ScaleY, OffsetX, OffsetY
            ScaleX = float.Parse(data["Def"]["ScaleX"]);
            ScaleY = float.Parse(data["Def"]["ScaleY"]);
            OffsetX = int.Parse(data["Def"]["OffsetX"]);
            OffsetY = int.Parse(data["Def"]["OffsetY"]);

            // Load Script
            if (!data["Def"]["Script"].Equals(""))
            {
                Load_Script(data["Def"]["Script"]);
            }
            else {
                mDataGrid.DataContext = null;
                mDataTable.Add(new MainTable() { Enable = true, Mode = "", Action = "", Event = "", Note = "" });
                mDataGrid.DataContext = mDataTable;
            }

            #endregion

            // Load Script setting
            LoadScriptSetting();

            // Load setting
            if (!mSettingHelper.OnOff_CrtlKey.Equals("0"))
            {
                OnOff_CrtlKey_Chk.IsChecked = true;
            }
            if (!mSettingHelper.Run_CrtlKey.Equals("0"))
            {
                Run_CrtlKey_Chk.IsChecked = true;
            }
            if (!mSettingHelper.Stop_CrtlKey.Equals("0"))
            {
                Stop_CrtlKey_Chk.IsChecked = true;
            }
            TextBox_OnOff_Hotkey.Text = mSettingHelper.OnOff_Hotkey;
            TextBox_Run_Hotkey.Text = mSettingHelper.Run_Hotkey;
            TextBox_Stop_Hotkey.Text = mSettingHelper.Stop_Hotkey;

            // Load HideOnSatrt setting
            if (!mSettingHelper.HideOnSatrt.Equals("0"))
            {
                HideOnSatrt_Toggle.IsOn = true;
            }

            // Load TestMode setting
            if (!mSettingHelper.TestMode.Equals("0"))
            {
                TestMode = true;
                TestMode_Toggle.IsOn = true;
            }
            else {
                TestMode = false;
                TestMode_Toggle.IsOn = false;
            }

            #region OverlayWindow

            // it is important to set the window to visible (and topmost) if you want to see it!
            _window = new OverlayWindow(0, 0, Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height)
            {
                IsTopmost = true,
                IsVisible = true
            };
            // handle this event to resize your Graphics surface
            //_window.SizeChanged += _window_SizeChanged;

            // initialize a new Graphics object
            // set everything before you call _graphics.Setup()
            _graphics = new GameOverlay.Drawing.Graphics
            {
                //MeasureFPS = true,
                Height = _window.Height,
                PerPrimitiveAntiAliasing = true,
                TextAntiAliasing = true,
                UseMultiThreadedFactories = false,
                VSync = true,
                Width = _window.Width,
                WindowHandle = IntPtr.Zero
            };

            _window.Create();
            _graphics.WindowHandle = _window.Handle; // set the target handle before calling Setup()         
            _graphics.Setup();

            _red = _graphics.CreateSolidBrush(GameOverlay.Drawing.Color.Red); // those are the only pre defined Colors                                                              
            _font = _graphics.CreateFont("Arial", 25); // creates a simple font with no additional style
            _black = _graphics.CreateSolidBrush(GameOverlay.Drawing.Color.Transparent);

            gfx = _graphics; // little shortcut
            #endregion

            KListener();

            Main_GlobalKeyUpHook = Hook.GlobalEvents();
            Main_GlobalKeyUpHook.KeyUp += Main_GlobalHookKeyUp;

            // NotifyIcon
            mNotifyIcon.Icon = System.Drawing.Icon.ExtractAssociatedIcon(System.Reflection.Assembly.GetExecutingAssembly().Location);
            mNotifyIcon.ContextMenuStrip = new System.Windows.Forms.ContextMenuStrip();
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("Visit_Website").ToString(), null, this.notifyIcon_Visit_Click);
            mNotifyIcon.ContextMenuStrip.Items.Add("-");
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("HideShow").ToString(), null, this.notifyIcon_DoubleClick);
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("Exit").ToString(), null, this.notifyIcon_Exit_Click);
            mNotifyIcon.Visible = true;
            mNotifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);
           

            if (!IsAdmin)
            {
                TextBlock_Please.Visibility = Visibility.Visible;
            }

            this.WindowState = WindowState.Minimized;
            this.Show();
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);

            if (mSettingHelper.HideOnSatrt.Equals("0"))
            {
                this.WindowState = WindowState.Normal;
            }
            else {
                this.Hide();
            }

            // test
            ManagementObjectSearcher objvida = new ManagementObjectSearcher("select * from Win32_VideoController ");
            string VC = String.Empty, DI = String.Empty;
            foreach (ManagementObject obj in objvida.Get())
            {
                if (obj["CurrentBitsPerPixel"] != null && obj["CurrentHorizontalResolution"] != null)
                {
                    if (((String)obj["DeviceID"]).IndexOf("VideoController1") != -1)
                    {
                        DI = obj["DeviceID"].ToString();
                        VC = obj["Description"].ToString();

                        Console.WriteLine("DeviceID: " + DI + ",Description: " + VC);
                    }
                }
            }

        }

        #region KListener
        private void KListener()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            Main_GlobalHook = Hook.GlobalEvents();
            Main_GlobalHook.KeyDown += Main_GlobalHookKeyPress;
        }
        private void UnKListener()
        {
            Main_GlobalHook.KeyDown -= Main_GlobalHookKeyPress;
            //It is recommened to dispose it
            Main_GlobalHook.Dispose();
        }

        private string PassCtrlKey = "";
        private double Timestamp;

        SortedList KeyList = new SortedList();

        DateTime StartTime = new DateTime(2001, 1, 1);

        private void Main_GlobalHookKeyPress(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            V V = new V();

            DateTime KeyTime = DateTime.Now;
            TimeSpan KeyTimeSpan = new TimeSpan(KeyTime.Ticks - StartTime.Ticks);
            
            double KeyTimeValue = Math.Floor(KeyTimeSpan.TotalMilliseconds);
            Console.WriteLine("{0} KeyTimeValue.", KeyTimeValue.ToString());

            if (KeyList.IndexOfKey(e.KeyValue) != -1) {

                double KeyListTimeValue = (double)KeyList.GetByIndex(KeyList.IndexOfKey(e.KeyValue));
                Console.WriteLine("{0} KeyListTimeValue.", KeyListTimeValue);
                Console.WriteLine("{0} KeyTimeValue- KeyListTimeValue.", KeyTimeValue- KeyListTimeValue);

                if ((KeyTimeValue- KeyListTimeValue) > 200)
                {
                    KeyList.RemoveAt(KeyList.IndexOfKey(e.KeyValue));
                }
                else {

                    return;
                }
                
            }
            else {
                KeyList.Add(e.KeyValue, KeyTimeValue);
            }

            if (!TextBox_Title.Text.Equals(""))
            {
                if (GetActiveWindowTitle() == null) { return; }
                string ActiveTitle = GetActiveWindowTitle();
                if (ActiveTitle.Length == ActiveTitle.Replace(TextBox_Title.Text, "").Length) { return; }
            }

            // stop KListener
            //UnKListener();

            // ON / OFF
            if (!Script_Toggle.IsOn && e.KeyCode.ToString().Equals(mSettingHelper.OnOff_Hotkey) && !(OnOff_CrtlKey_Chk.IsChecked == true && e.Control == false)) // def "'"
            {
                if (Btn_ON.Content.Equals("ON"))
                {
                    Btn_ON.Content = "OFF";
                    Btn_ON.Foreground = System.Windows.Media.Brushes.Red;

                    for (int i = 0; i < _workerThreads.Count; i++)
                    {
                        if (_workerThreads[i].IsAlive)
                        {
                            _workerThreads[i].Abort();
                            _workerThreads[i] = null;

                            _workerThreads[i] = new Thread(() =>
                            {
                                try
                                {
                                    Script(Load_Script_to_DataTable(eDataTable[i].eTable_Path), "Def");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("{0} Exception caught.", ex);
                                }
                                finally { 
                                    GC.Collect(); 
                                    GC.WaitForPendingFinalizers();
                                }
                            });
                        }
                        eDataTable[i].eTable_State = "Stop";
                    }

                    if (Ring.IsActive == true) {
                        Stop_script();
                    }
                }
                else
                {
                    Btn_ON.Content = "ON";
                    Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                }
                eDataGrid.DataContext = null;
                eDataGrid.DataContext = eDataTable;
            }
            if (!Btn_ON.Content.Equals("ON")) {
                //KListener(); 
                return; }

            Console.WriteLine(e.KeyCode.ToString());


            if (e.KeyCode.ToString().Equals(mSettingHelper.Run_Hotkey) && !(Run_CrtlKey_Chk.IsChecked == true && e.Control == false)) // def "["
            {
                AlertSound();
                ShowBalloon("Run", "...");
                Run_script();
            }
            if (e.KeyCode.ToString().Equals(mSettingHelper.Stop_Hotkey) && !(Stop_CrtlKey_Chk.IsChecked == true && e.Control == false)) // def "]"
            {
                AlertSound();
                Stop_script();
                ShowBalloon("Stop", "...");
            }

            PassCtrlKey = "";
            if (e.Control)
            {
                PassCtrlKey = "Ctrl+";
            }
            else if (e.Alt)
            {
                PassCtrlKey = "Alt+";
            }
            else if (e.Shift)
            {
                PassCtrlKey = "Shift+";
            }

            string KeyCode = PassCtrlKey + e.KeyCode.ToString();

            double TimestampNow = (double)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds;
            if (TimestampNow - Timestamp > 2000)
            {
               // PassCtrlKey = "";
            }

            //switch (e.KeyCode.ToString())
            //{
            //    case "LControlKey":
            //        PassCtrlKey = "Ctrl+";
            //        Timestamp = TimestampNow;
            //        break;
            //    case "LMenu":
            //        PassCtrlKey = "Alt+";
            //        break;
            //    case "LShiftKey":
            //    case "RShiftKey":
            //        PassCtrlKey = "Shift+";
            //        break;
            //    default:
            //        break;
            //}

            // Select Script
            bool IsRun = false;
            for (int i = 0; i < eDataTable.Count; i++)
            {
                if (KeyCode.Equals(eDataTable[i].eTable_Key) && eDataTable[i].eTable_Enable == true)
                {
                    IsRun = true;
                    Console.WriteLine("START " + _workerThreads[i].ThreadState.ToString());

                    //if (_workerThreads[i].ThreadState == System.Threading.ThreadState.WaitSleepJoin){
                    //    break;
                    //}

                    if (!_workerThreads[i].IsAlive)
                    {
                        if (_workerThreads[i].ThreadState != System.Threading.ThreadState.Stopped && _workerThreads[i].ThreadState != System.Threading.ThreadState.Unstarted)
                        {
                            _workerThreads[i].Start();
                        }
                        else
                        {
                            Console.WriteLine(eDataTable[i].eTable_Path);

                            _workerThreads[i] = null;
                            _workerThreads[i] = new Thread(() =>
                            {
                                try
                                {
                                    Script(Load_Script_to_DataTable(eDataTable[i].eTable_Path), "Def");
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine("{0} Exception caught.", ex);

                                }
                                finally
                                {
                                    Console.WriteLine("Memory used before collection:       {0:N0}",GC.GetTotalMemory(false));
                                    GC.Collect();
                                    GC.WaitForPendingFinalizers();
                                    Console.WriteLine("WaitForPendingFinalizers");
                                    Console.WriteLine("Memory used after full collection:   {0:N0}", GC.GetTotalMemory(false));
                                }
                            });
                            _workerThreads[i].Start();
                        }

                        eDataTable[i].eTable_State = "Running";
                        ShowBalloon("Running ", eDataTable[i].eTable_Name);
                        Console.WriteLine(_workerThreads[i].ThreadState.ToString());
                    }
                    else
                    {
                        Console.WriteLine(eDataTable[i].eTable_Path);

                        _workerThreads[i].Abort();
                        _workerThreads[i] = null;

                        _workerThreads[i] = new Thread(() =>
                        {
                            try
                            {
                                Script(Load_Script_to_DataTable(eDataTable[i].eTable_Path), "Def");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("{0} Exception caught.", ex);
                            }
                            finally
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                                Console.WriteLine("WaitForPendingFinalizers");
                            }
                        });

                        eDataTable[i].eTable_State = "Stop";
                        ShowBalloon("Stop ", eDataTable[i].eTable_Name);
                        Console.WriteLine(_workerThreads[i].ThreadState.ToString());
                    }
                }
            }

            // eDataGrid
            if (!eDataGrid.IsKeyboardFocusWithin && IsRun) {
                eDataGrid.DataContext = null;
                eDataGrid.DataContext = eDataTable;
            }
           
            // Restart KListener
            //KListener();
        }
        #endregion

        // For resize x,y
        private float ScaleX, ScaleY;
        private int OffsetX, OffsetY;
        private string copy = "";

        private void Script(List<MainTable> minDataTable,string Mode)
        {
            InputSimulator mInputSimulator = new InputSimulator();
            Keyboard ky = new Keyboard();
            ConvertHelper ConvertHelper = new ConvertHelper();
            SortedList mDoSortedList = new SortedList();
            V V = new V();

            SortedList DefValueList = new SortedList();
            System.Drawing.Point StartPoint = new System.Drawing.Point();
            GetCursorPos(ref StartPoint);
            DefValueList.Add("{StartPostion}", StartPoint.X.ToString() + "," + StartPoint.Y.ToString());

            // key || value ex:
            //mDoSortedList.Add("Point", "0,0");
            //mDoSortedList.Add("Point Array", "0,0,0,0");
            //mDoSortedList.Add("Draw", "");
            //mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey("Draw"));

            int n = 0; int LoopCount = 0;
            //DateTime centuryBegin = new DateTime(2001, 1, 1);
            //DateTime currentDate = DateTime.Now;
            //int timer = 0;
            while (n < minDataTable.Count)
            {

                //long elapsedTicks = currentDate.Ticks - centuryBegin.Ticks;
                //TimeSpan elapsedSpan = new TimeSpan(elapsedTicks);
                //Console.WriteLine(Math.Floor(elapsedSpan.TotalMinutes));

                if (Mode.Equals("Debug") && TestMode)
                {
                    int number = 80000;
                    number += n;
                    CreateMessage(number.ToString());
                    Thread.Sleep(1000);
                }

                string Command = minDataTable[n].Mode;
                string CommandData = minDataTable[n].Action;
                bool CommandEnable = minDataTable[n].Enable;

                if (DefValueList.IndexOfKey(CommandData) != -1)
                {
                    CommandData = DefValueList.GetByIndex(DefValueList.IndexOfKey(CommandData)).ToString();
                }

                string[] Event = minDataTable[n].Event.Split(',');
                if (minDataTable[n].Event == "") { Event = new string[0]; }

                if (CommandEnable)
                {
                  
                    Mat matTemplate = null, matTarget = null;
                    Boolean err = false;
                    try
                    {
                        #region Switch Command
                        switch (Command)
                        {
                            case "Move":

                                if (Event.Length == 0)
                                {
                                    string[] str_move = CommandData.Split(',');
                                    SetCursorPos(V.Get_ValueX(int.Parse(str_move[0]), ScaleX, OffsetX), V.Get_ValueY(int.Parse(str_move[1]), ScaleY, OffsetY));
                                }
                                else
                                {
                                    // Check Key
                                    if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                    {
                                        string[] Event_Data;
                                        if (CommandData.Length > 0)
                                        {
                                            // use CommandData 
                                            Event_Data = CommandData.Split(',');
                                        }
                                        else
                                        {
                                            // Get SortedList Value by Key
                                            Event_Data = mDoSortedList.GetByIndex(mDoSortedList.IndexOfKey(Event[0])).ToString().Split(',');
                                        }
                                        SetCursorPos(int.Parse(Event_Data[0]), int.Parse(Event_Data[1]));
                                    }
                                }

                                break;

                            case "Offset":

                                if (Event.Length == 0)
                                {
                                    string[] mOffset = CommandData.Split(',');
                                    System.Drawing.Point point = System.Windows.Forms.Control.MousePosition;
                                    SetCursorPos(point.X + V.Get_ValueX(int.Parse(mOffset[0]), ScaleX, 0), point.Y + V.Get_ValueY(int.Parse(mOffset[1]), ScaleY, 0));
                                }
                                else
                                {
                                    // Check Key
                                    if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                    {
                                        string[] Event_Data;
                                        if (CommandData.Length > 0)
                                        {
                                            // use CommandData 
                                            Event_Data = CommandData.Split(',');
                                        }
                                        else
                                        {
                                            // TODO: ?
                                            // Get SortedList Value by Key
                                            Event_Data = mDoSortedList.GetByIndex(mDoSortedList.IndexOfKey(Event[0])).ToString().Split(',');
                                        }
                                        System.Drawing.Point point = System.Windows.Forms.Control.MousePosition;
                                        SetCursorPos(point.X + V.Get_ValueX(int.Parse(Event_Data[0]), ScaleX, 0), point.Y + V.Get_ValueY(int.Parse(Event_Data[1]), ScaleY, 0));
                                    }
                                }

                                break;

                            case "Delay":

                                if (Event.Length == 0 || mDoSortedList.IndexOfKey(Event[0]) != -1)
                                {
                                    Thread.Sleep(Int32.Parse(CommandData));
                                }

                                break;

                            case "Click":

                                if (Event.Length == 0 || mDoSortedList.IndexOfKey(Event[0]) != -1)
                                {
                                    CommandData = CommandData.Trim().ToUpper();
                                    if (CommandData.Equals("LEFT"))
                                    {
                                        mInputSimulator.Mouse.LeftButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.LeftButtonUp();
                                    }
                                    if (CommandData.Equals("LEFT_DOWN"))
                                    {
                                        mInputSimulator.Mouse.LeftButtonDown();
                                    }
                                    if (CommandData.Equals("LEFT_UP"))
                                    {
                                        mInputSimulator.Mouse.LeftButtonUp();
                                    }
                                    if (CommandData.Equals("RIGHT"))
                                    {
                                        mInputSimulator.Mouse.RightButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.RightButtonUp();
                                    }
                                    if (CommandData.Equals("RIGHT_DOWN"))
                                    {
                                        mInputSimulator.Mouse.RightButtonDown();
                                    }
                                    if (CommandData.Equals("RIGHT_UP"))
                                    {
                                        mInputSimulator.Mouse.RightButtonUp();
                                    }
                                }

                                break;

                            case "Match&Draw": // get
                            case "Match":
                            case "Match RGB":
                                do
                                {
                                    // img,x,y,x-length,y-length,threshold
                                    string[] MatchArr = V.Get_Split(CommandData);
                                    double threshold = 0.9;

                                    if (MatchArr[0].Equals(""))
                                    {
                                        MatchArr[0] = "example/s.png";
                                    }

                                    if (MatchArr.Length > 2)
                                    {
                                        matTarget = BitmapConverter.ToMat(makeScreenshot_clip(
                                            int.Parse(MatchArr[1]), int.Parse(MatchArr[2]),
                                            int.Parse(MatchArr[4]), int.Parse(MatchArr[3])));

                                        if (MatchArr.Length > 5)
                                        {
                                            threshold = Convert.ToDouble(MatchArr[5]);
                                        }
                                    }
                                    else
                                    {
                                        Bitmap screenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                                        System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);
                                        gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                                        matTarget = BitmapConverter.ToMat(screenshot);

                                        gfxScreenshot.Dispose();
                                        screenshot.Dispose();

                                        if (MatchArr.Length > 1)
                                        {
                                            threshold = Convert.ToDouble(MatchArr[1]);
                                        }
                                    }

                                    matTemplate = new Mat(MatchArr[0], ImreadModes.Color);
                                    int temp_w = matTemplate.Width / 2, temp_h = matTemplate.Height / 2; // center x y
                                   
                                    string return_xy;
                                    if (Command.Equals("Match RGB"))
                                    {
                                        return_xy = RunTemplateMatch(matTarget, matTemplate, "RGB", threshold);
                                    }
                                    else
                                    {
                                        return_xy = RunTemplateMatch(matTarget, matTemplate, "GRAY", threshold);
                                    }

                                    if (!return_xy.Equals(""))
                                    {
                                        // Screenshot clip offset
                                        int x_offset = 0, y_offset = 0;
                                        if (MatchArr.Length > 2)
                                        {
                                            x_offset = int.Parse(MatchArr[1]);
                                            y_offset = int.Parse(MatchArr[2]);
                                        }

                                        string[] xy = return_xy.Split(',');
                                        for (int j = 0; j < xy.Length-1; j += 2) {
                                            xy[j] = (int.Parse(xy[j]) + x_offset).ToString();
                                            xy[j+1] = (int.Parse(xy[j+1]) + y_offset).ToString();
                                        }

                                        if (Command.Equals("Match&Draw"))
                                        {
                                            gfx.BeginScene();
                                            //gfx.DrawTextWithBackground(_font, _red, _black, 10, 10, return_xyd.ToString());
                                            gfx.DrawRoundedRectangle(_red, RoundedRectangle.Create(int.Parse(xy[0]), int.Parse(xy[1]), temp_w * 2, temp_h * 2, 6), 2);
                                            gfx.EndScene();
                                        }
                                    
                                        // Add Key
                                        if (Event.Length != 0)
                                        {
                                            if (Event[0].IndexOf("-") == -1)
                                            {
                                                if (V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                                {
                                                    mDoSortedList.Add(Event[0], string.Join(",", xy));
                                                }
                                            }
                                            else
                                            {
                                                if (mDoSortedList.IndexOfKey(Event[0].Replace("-", "")) == -1)
                                                {
                                                    mDoSortedList.Add(Event[0].Replace("-", ""), string.Join(",", xy));
                                                }
                                            }
                                        }
                                    }
                                    matTarget.Dispose();
                                    matTemplate.Dispose();

                                } while (false);

                                break;

                            case "Sift Match":

                                Mat matTarget_Sift = BitmapConverter.ToMat(makeScreenshot());
                                Mat matTemplate_Sift = new Mat("s.png", ImreadModes.Color);

                                MatchBySift(matTarget_Sift, matTemplate_Sift);

                                break;

                            case "Key":

                                if (Event.Length != 0 && V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                {
                                    break;
                                }

                                #region SendKeys
                                //Char[] mChar = CommandData.ToCharArray();
                                //for (int j = 0; j < mChar.Length; j++)
                                //{
                                //    keybd_event(VkKeyScan(mChar[j]), 0, KEYEVENTF_EXTENDEDKEY, 0);
                                //    keybd_event(VkKeyScan(mChar[j]), 0, KEYEVENTF_KEYUP, 0);
                                //}
                                //alt = % , shift = + , ctrl = ^ , enter = {ENTER}
                                // EX: ^{ v}  
                                //SendKeys.SendWait(CommandData);

                                //KeysConverter kc = new KeysConverter();
                                //string keyChar = kc.ConvertToString("A");
                                #endregion

                                //***************** InputSimulator *****************
                                if (CommandData.Substring(0, 1).Equals(",")){
                                    CommandData = CommandData.Replace(",", "OEM_COMMA");
                                }
                                string[] mKey = CommandData.ToUpper().Split(',');

                                // ConvertKeyString
                                mKey[0] = ConvertHelper.ConvertKeyString(mKey[0].Trim().ToUpper());
                                if (mKey[0].Length == 1)
                                {
                                    mKey[0] = "VK_" + mKey[0];
                                }

                                VirtualKeyCode mKeyCode = (VirtualKeyCode)ConvertHelper.StringToVirtualKeyCode(mKey[0]);
                                if (mKeyCode != 0)
                                {
                                    if (mKey.Length == 2)
                                    {
                                        if (mKey[1].Equals("DOWN"))
                                        {
                                            mInputSimulator.Keyboard.KeyDown(mKeyCode);
                                        }
                                        else {
                                            mInputSimulator.Keyboard.KeyUp(mKeyCode);
                                        }
                                    }
                                    else {
                                        mInputSimulator.Keyboard.KeyDown(mKeyCode);
                                        Thread.Sleep(250);
                                        mInputSimulator.Keyboard.KeyUp(mKeyCode);
                                    }
                                }
                                else {
                                    string str = CommandData;
                                    char[] arr = str.ToCharArray();

                                    foreach (char c in arr)
                                    {
                                        mInputSimulator.Keyboard.KeyDown((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
                                        Thread.Sleep(100);
                                        mInputSimulator.Keyboard.KeyUp((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
                                    }
                                }
                                //mInputSimulator.Keyboard.KeyPress((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
                                //VirtualKeyCode myEnum = (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), "Enter");
                                //***************** InputSimulator *****************

                                break;

                            case "SendKeyDown":
                            case "SendKeyUp":

                                if (Event.Length != 0 && V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                {
                                    break;
                                }
                                else {
                                    string SendKeyStr = ConvertHelper.ConvertKeyString(CommandData.Trim().ToUpper());

                                    if (SendKeyStr.Length == 1)
                                    {
                                        SendKeyStr = "KEY_" + SendKeyStr;
                                    }

                                    short value = 0;
                                    Array enumValueArray = Enum.GetValues(typeof(ScanCodeShort));
                                    foreach (short enumValue in enumValueArray)
                                    {
                                        if (Enum.GetName(typeof(ScanCodeShort), enumValue).Equals(SendKeyStr))
                                        {
                                            value = enumValue;

                                            if (Command.Equals("SendKeyDown"))
                                            {
                                                ky.SendKeyDown(value);
                                            }
                                            else {
                                                ky.SendKeyUp(value);
                                            }
                                        }
                                    }
                                }

                                break;

                            case "ModifierKey":

                                if (Event.Length != 0 && V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                {
                                    break;
                                }

                                // For example CTRL-ALT-SHIFT-ESC-K which is simulated as
                                // CTRL-down, ALT-down, SHIFT-down, press ESC, press K, SHIFT-up, ALT-up, CTRL-up
                                //mInputSimulator.Keyboard.ModifiedKeyStroke(
                                //  new[] { VirtualKeyCode.CONTROL, VirtualKeyCode.MENU, VirtualKeyCode.SHIFT },
                                //  new[] { VirtualKeyCode.ESCAPE, VirtualKeyCode.VK_K });

                                String[] KeyStr = CommandData.Split('|');
                                String[] ModifierKeyArr = KeyStr[0].Split(',');
                                String[] KeyArr = KeyStr[1].Split(',');

                                VirtualKeyCode[] ModifierKeyCodeArr = new VirtualKeyCode[ModifierKeyArr.Length];
                                VirtualKeyCode[] KeyCodeArr = new VirtualKeyCode[KeyArr.Length];
                                for (int i = 0; i < ModifierKeyCodeArr.Length; i++)
                                {
                                    ModifierKeyArr[i] = ConvertHelper.ConvertKeyString(ModifierKeyArr[i]);

                                    if (ModifierKeyArr[i].Length == 1) {
                                        ModifierKeyArr[i] = "VK_"+ModifierKeyArr[i];
                                    }
                                    ModifierKeyCodeArr[i] = (VirtualKeyCode)ConvertHelper.StringToVirtualKeyCode(ModifierKeyArr[i]);
                                   
                                }
                                for (int i = 0; i < KeyCodeArr.Length; i++)
                                {
                                    if (KeyArr[i].Length == 1)
                                    {
                                        KeyArr[i] = "VK_" + KeyArr[i];
                                    }
                                    KeyCodeArr[i] = (VirtualKeyCode)ConvertHelper.StringToVirtualKeyCode(KeyArr[i]);
                                }

                                mInputSimulator.Keyboard.ModifiedKeyStroke(ModifierKeyCodeArr, KeyCodeArr);

                                //mInputSimulator.Keyboard.ModifiedKeyStroke(VirtualKeyCode.MENU, VirtualKeyCode.TAB);
                                //mInputSimulator.Keyboard.ModifiedKeyStroke((VirtualKeyCode)ConvertHelper.StringToVirtualKeyCode("MENU"), (VirtualKeyCode)ConvertHelper.StringToVirtualKeyCode("TAB"));

                                break;

                            case "RemoveEvent":

                                if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                {
                                    string[] TempEvent_Data = mDoSortedList.GetByIndex(mDoSortedList.IndexOfKey(Event[0])).ToString().Split(',');
                                    mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey(Event[0]));

                                    if (CommandData.ToUpper().Equals("PUSH"))
                                    {
                                        if (TempEvent_Data.Length >= 4) {
                                            mDoSortedList.Add(Event[0], string.Join(",", TempEvent_Data.Skip(2).ToArray()));
                                        }
                                    }
                                }

                                break;

                            case "GetPoint":

                                // Add Key
                                if (Event[0].Length > 0)
                                {
                                    System.Drawing.Point point = System.Windows.Forms.Control.MousePosition;
                                    mDoSortedList.Add(Event[0], point.X.ToString() + "," + point.Y.ToString());
                                }

                                break;

                            case "Run .exe":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    try
                                    {
                                        Process.Start(CommandData);
                                    }
                                    catch {

                                    }
                                }
                               
                                break;

                            case "FindWindow":

                                // and window name were obtained using the Spy++ tool.
                                IntPtr calculatorHandle = FindWindow(null, CommandData);

                                // Verify that Calculator is a running process.
                                if (calculatorHandle == IntPtr.Zero)
                                {
                                    //System.Windows.MessageBox.Show("is not running...");
                                    //return;
                                }

                                // Make Calculator the foreground application and send it 
                                // a set of calculations.
                                SetForegroundWindow(calculatorHandle);

                                break;

                            case "ScreenClip":

                                string[] str_clip = CommandData.Split(',');
                                TempBitmap = makeScreenshot_clip(int.Parse(str_clip[0]), int.Parse(str_clip[1]), int.Parse(str_clip[2]), int.Parse(str_clip[3]));

                                break;

                            case "Clear Screen":

                                gfx.BeginScene(); // call before you start any drawing
                                gfx.ClearScene();
                                gfx.EndScene();

                                break;

                            case "PlaySound":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    SoundPlayer mWaveFile = null;
                                    try
                                    {
                                        mWaveFile = new SoundPlayer(CommandData);
                                        mWaveFile.PlaySync();
                                    }
                                    catch
                                    {

                                    }
                                    finally {
                                        mWaveFile?.Dispose();
                                    }
                                }

                                break;

                            case "WriteClipboard":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    copy = CommandData;
                                    CreateMessage("9486");
                                }

                                break;

                            case "Color Test":

                                Mat mat_screen = new Mat();

                                Bitmap mscreenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                                System.Drawing.Graphics mgfxScreenshot = System.Drawing.Graphics.FromImage(mscreenshot);
                                mgfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                                mat_screen = BitmapConverter.ToMat(mscreenshot);

                                Mat mask = new Mat();
                                Scalar low_value = new Scalar(100, 100, 100);
                                Scalar high_value = new Scalar(255, 255, 255);
                                Cv2.InRange(mat_screen, low_value, high_value, mask);
                                Cv2.ImShow("a", mask);
                                Cv2.WaitKey();

                                break;

                            case "PostMessage":

                                if (Event.Length == 0)
                                {
                                    string[] send = CommandData.Split(',');

                                    IntPtr windowHandle;
                                    if (send[0].Equals("T"))
                                    {
                                        windowHandle = FindWindow(null, send[1]);
                                    }
                                    else
                                    {
                                        windowHandle = FindWindow(send[1], null);
                                    }

                                    IntPtr editHandle = windowHandle;


                                    if (editHandle == IntPtr.Zero)
                                    {
                                        System.Windows.MessageBox.Show("is not running...");
                                    }
                                    else
                                    {
                                        gfx.BeginScene(); // call before you start any drawing
                                        gfx.DrawTextWithBackground(_font, _red, _black, 10, 10, windowHandle.ToString());
                                        gfx.EndScene();
                                    }

                                    for (int i = 2; i < send.Length; i++)
                                    {
                                        int value = (int)new System.ComponentModel.Int32Converter().ConvertFromString(send[i]);
                                        SendMessage((int)editHandle, WM_KEYDOWN, 0, "0x014B0001");
                                        Thread.Sleep(50);
                                        SendMessage((int)editHandle, WM_KEYUP, 0, "0xC14B0001");
                                        Thread.Sleep(50);
                                    }


                                    //Process[] processlist = Process.GetProcesses();

                                    //string titleText = "";
                                    //foreach (Process process in processlist)
                                    //{
                                    //    if (!String.IsNullOrEmpty(process.MainWindowTitle))
                                    //    {
                                    //        Console.WriteLine("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle);
                                    //        titleText += process.MainWindowTitle.ToString();
                                    //    }
                                    //}

                                    ////System.Windows.MessageBox.Show(titleText);
                                    //string out_string = titleText;

                                    //System.IO.File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/" + "out" + ".txt", out_string);
                                }
                                else
                                {
                                    // Check Key
                                    if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                    {
                                        string[] Event_Data;
                                        if (CommandData.Length > 0)
                                        {
                                            // use CommandData 
                                            Event_Data = CommandData.Split(',');
                                        }
                                        else
                                        {
                                            // Get SortedList Value by Key
                                            Event_Data = mDoSortedList.GetByIndex(mDoSortedList.IndexOfKey(Event[0])).ToString().Split(',');
                                        }
                                    }
                                }

                                break;

                            case "Jump":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    n += int.Parse(CommandData);
                                }

                                break;

                            case "Goto":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    n = int.Parse(CommandData) -2;
                                }

                                break;

                            case "Loop":
                                do
                                {
                                    if (!CommandData.Equals(""))
                                    {
                                        LoopCount++;
                                        int LoopSum = int.Parse(CommandData);
                                        if (LoopCount != LoopSum)
                                        {
                                            n = -1;
                                        }
                                    }
                                    else
                                    {
                                        n = -1;
                                    }
                                } while (false);

                                break;

                            default:

                                break;
                        }
                        #endregion
                    }
                    catch(Exception e)
                    {
                        Console.WriteLine("{0} Exception caught.", e);

                        if (e.Message.ToString().IndexOf("Thread") == -1 && !e.Source.Equals("System.Drawing"))
                        {
                            if (Mode.Equals("Debug"))
                            {
                                System.Windows.MessageBox.Show("[Error] Line " + (n + 1).ToString() + " \nMessage: " + e.Message);

                                // debug stop msg
                                CreateMessage("9487");

                                break;
                            }
                        }

                        if (e.Source.Equals("System.Drawing"))
                        {
                            Thread.Sleep(1000);
                            n--;
                        }
                        else
                        {
                            mInputSimulator = null;
                            ky = null;
                            V = null;
                            ConvertHelper = null;
                            mDoSortedList.Clear();

                            err = true;
                        }
                    }
                    finally
                    {
                        if (err) { 
                        
                        }

                        matTemplate?.Dispose();
                        matTarget?.Dispose();

                    }
                }

                n++;
            }

            if (Mode.Equals("Debug")){
                CreateMessage("9487");
            }
            // script eng msg
            CreateMessage("1000");

        }

        #region APP
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);
        private uint MSG_SHOW;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
        }

        int LastNumber = 0;
        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            //Console.WriteLine(lParam);
            if (msg == MSG_SHOW)
            {
                if (wParam.ToString().Equals("9487"))
                {
                    Ring.IsActive = false;
                }

                if (wParam.ToString().Equals("9486"))
                {
                    System.Windows.Clipboard.SetText(copy);
                }

                // TestMode
                if (wParam.ToString().Substring(0, 1).Equals("8"))
                {
                    DataGridRow row = (DataGridRow)mDataGrid.ItemContainerGenerator.ContainerFromIndex(LastNumber);
                    if (row != null)
                    {
                        SolidColorBrush brush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 255, 255));
                        row.Background = brush;
                    }

                    int number = int.Parse(wParam.ToString().Substring(1, 4));
                    row = (DataGridRow)mDataGrid.ItemContainerGenerator.ContainerFromIndex(number);
                    if (row != null)
                    {
                        LastNumber = number;
                        SolidColorBrush brush = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 104, 0));
                        row.Background = brush;
                    }
                }

                // Update Thread status
                if (wParam.ToString().Equals("1001"))
                {
                    for (int i = 0; i < _workerThreads.Count; i++)
                    {
                        if (!_workerThreads[i].IsAlive)
                        {
                            eDataTable[i].eTable_State = "Stop";

                            // eDataGrid
                            if (!eDataGrid.IsKeyboardFocusWithin)
                            {
                                eDataGrid.DataContext = null;
                                eDataGrid.DataContext = eDataTable;
                            }
                        }
                    }
                }

                // On ScriptEnd delay 100ms
                if (wParam.ToString().Equals("1000"))
                {
                    System.Windows.Window window = System.Windows.Window.GetWindow(this);
                    var wih = new WindowInteropHelper(window);
                    IntPtr mPrt = wih.Handle;

                    Thread TempThread = new Thread(() =>
                    {
                        Thread.Sleep(100);

                        if (mPrt != IntPtr.Zero)
                        {
                            string mParam = "1001";
                            int WParam = int.Parse(mParam);
                            SendMessage((int)mPrt, (int)MSG_SHOW, WParam, "0x00000001");
                        }

                        GC.Collect();
                        GC.WaitForPendingFinalizers();
                    });
                    TempThread.Start();
                }

                if (wParam.ToString().Substring(0, 2).Equals("22"))
                {
                    switch (wParam.ToString())
                    {
                        case "2200":
                            this.ShowMessageAsync("Load_Script_to_DataTable", "Error!");

                            break;
                        default:

                            break;
                    }
                }

            }
            return IntPtr.Zero;
        }
        private void CreateMessage(string Param)
        {
            IntPtr mPrt = FindWindow(null, "Scriptboxie");
            if (mPrt != IntPtr.Zero)
            {
                int wParam = int.Parse(Param);
                SendMessage((int)mPrt, (int)MSG_SHOW, wParam, "0x00000001");
            }
        }
        private void ShowBalloon(string title, string msg)
        {
            // Hide BalloonTip
            if (mNotifyIcon.Visible)
            {
                mNotifyIcon.Visible = false;
                mNotifyIcon.Visible = true;
            }

            if (mSettingHelper.ShowBalloon)
            {
                mNotifyIcon.ShowBalloonTip(500, title, msg, ToolTipIcon.None);
            }
        }
        private void notifyIcon_DoubleClick(object Sender, EventArgs e)
        {
            if (this.WindowState == WindowState.Normal)
            {
                this.WindowState = WindowState.Minimized;
                this.Hide();
            }
            else
            {
                this.Show();
                this.WindowState = WindowState.Normal;

                // Activate the form.
                this.Activate();
            }
        }

        private void notifyIcon_Visit_Click(object sender, EventArgs e)
        {
            Process.Start("https://github.com/gemilepus/Scriptboxie");
        }
        private void notifyIcon_Exit_Click(object sender, EventArgs e)
        {
            this.Close();
        }
        private void AlertSound()
        {
            try
            {
                mAlertSound.Play();
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }

        }
        private void TextBox_Title_TextChanged(object sender, TextChangedEventArgs e)
        {
            // .ini
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            data["Def"]["WindowTitle"] = TextBox_Title.Text;
            parser.WriteFile("user.ini", data);
        }
        private void Btn_ON_Click(object sender, RoutedEventArgs e)
        {
            if (Script_Toggle.IsOn) {
                return;
            }
            if (Btn_ON.Content.Equals("ON"))
            {
                Btn_ON.Content = "OFF";
                Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
            }
            else
            {
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            }
        }
        private void Btn_About_Click(object sender, RoutedEventArgs e)
        {
            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();

            this.ShowMessageAsync("About",
                "v" + System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion + "\n"
                + "Author: " + "gemilepus" + "\n"
                + "Mail: " + "gemilepus@gmail.com" + "\n"
                );
        }

        private void NewVersion_PreviewMouseDown(object sender, MouseButtonEventArgs e)
        {
            Process.Start("https://github.com/gemilepus/Scriptboxie/releases");
        }

        private void MetroWindow_Loaded(object sender, RoutedEventArgs e)
        {
            string tag_name = CheckUpdate();

            System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
            if (!tag_name.ToString().Equals("v" + System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion))
            {
                if (!tag_name.ToString().Equals(""))
                {
                    NewVersion.Text = NewVersion.Text + " " + tag_name;
                    NewVersion.Visibility = Visibility.Visible;
                }
               
            }
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            mSettingHelper.End(this);

            // Stop all thread
            for (int i = 0; i < _workerThreads.Count; i++)
            {
                if (_workerThreads[i].IsAlive)
                {
                    _workerThreads[i].Abort();
                }
            }

            if (Ring.IsActive)
            {
                mThread.Abort();
            }

            mNotifyIcon.Visible = false;
        }

        #endregion

        #region Script Panel
        private void Load_Script_ini()
        {
            string fileContent = string.Empty;
            StreamReader reader = new StreamReader(System.Windows.Forms.Application.StartupPath + "/" + "Script.ini");

            // read test
            fileContent = reader.ReadToEnd();
            fileContent.Replace(";", "%;");
            string[] SplitStr = fileContent.Split(';');

            // Table Clear
            eDataGrid.DataContext = null;
            eDataTable.Clear();

            for (int i = 0; i < SplitStr.Length - 6; i += 6)
            {
                eDataTable.Add(new EditTable()
                {
                    eTable_Enable = bool.Parse(SplitStr[i].Replace("%", "")),
                    eTable_Name = SplitStr[i + 1].Replace("%", ""),
                    eTable_Key = SplitStr[i + 2].Replace("%", ""),
                    eTable_State = "",
                    eTable_Note = SplitStr[i + 4].Replace("%", ""),
                    eTable_Path = SplitStr[i + 5].Replace("%", "")
                });
            }
            eDataGrid.DataContext = eDataTable;

            reader.Close();
        }
        private void Save_Script() // async
        {
            string out_string = "";
            for (int i = 0; i < eDataTable.Count; i++)
            {
                out_string += eDataTable[i].eTable_Enable + ";"
                    + eDataTable[i].eTable_Name + ";"
                    + eDataTable[i].eTable_Key + ";"
                    + eDataTable[i].eTable_State + ";"
                    + eDataTable[i].eTable_Note + ";"
                    + eDataTable[i].eTable_Path + ";"
                    + "\n";
            }
            System.IO.File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/" + "Script.ini", out_string);
        }

        private void Load_Script(string filePath)
        {
            string[] mPath = filePath.Split('\\');
            ScriptName.Text = mPath[mPath.Length-1];
            ScriptName.ToolTip = filePath;
            // Table Clear
            mDataGrid.DataContext = null;
            mDataTable.Clear();

            mDataTable = Load_Script_to_DataTable(filePath);
            mDataGrid.DataContext = mDataTable;

            // .ini
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile("user.ini");
            data["Def"]["Script"] = filePath;
            parser.WriteFile("user.ini", data);
        }

        private List<MainTable> Load_Script_to_DataTable(string mfilePath)
        {
            List<MainTable> tempDataTable = new List<MainTable>();
            string fileContent = string.Empty;

            try
            {
                StreamReader reader = new StreamReader(mfilePath);

                fileContent = reader.ReadLine();
                int col = fileContent.Length- fileContent.Replace(";", "").Length;

                fileContent += reader.ReadToEnd();
                fileContent.Replace(";", "%;");
                //fileContent = fileContent.Replace(System.Environment.NewLine, "");
                string[] SplitStr = fileContent.Split(';');

                if (!fileContent.Substring(0, 2).Equals("[{"))
                {
                    for (int i = 0; i < SplitStr.Length - col; i += col)
                    {
                        string mMode = SplitStr[i + 1].Replace("%", "");
                        switch (mMode)
                        {
                            case "Draw":
                                mMode = "Match&Draw";
                                break;
                            case "RemoveKey":
                                mMode = "RemoveEvent";
                                break;
                            case "Clean Draw":
                                mMode = "Clear Screen";
                                break;
                            default:
                                break;
                        }

                        if (col == 5)
                        {
                            tempDataTable.Add(new MainTable()
                            {
                                Enable = bool.Parse(SplitStr[i].Replace("%", "")),
                                Mode = mMode,
                                Action = SplitStr[i + 2].Replace("%", ""),
                                Event = SplitStr[i + 3].Replace("%", ""),
                                Note = SplitStr[i + 4].Replace("%", "")
                            });
                        }
                        else
                        {
                            tempDataTable.Add(new MainTable()
                            {
                                Enable = bool.Parse(SplitStr[i].Replace("%", "")),
                                Mode = mMode,
                                Action = SplitStr[i + 2].Replace("%", ""),
                                Event = SplitStr[i + 3].Replace("%", ""),
                                Note = ""
                            });
                        }

                    }
                }
                else {
                    tempDataTable = JsonSerializer.Deserialize<List<MainTable>>(fileContent);
                }

                reader.Close();
            }
            catch
            {
                bool fileExist = File.Exists(mfilePath);
                if (fileExist)
                {
                    this.ShowMessageAsync("Error", "read error occurred!\n" +
                        "File: "+ mfilePath);
                }
            }
            return tempDataTable;
        }
        private void LoadScriptSetting() {
            // Load Script setting
            try
            {
                Load_Script_ini();
            }
            catch (Exception ex)
            {
                Console.WriteLine("{0} Exception caught.", ex);

                if (!File.Exists("Script.ini"))
                {
                    eDataGrid.DataContext = null;
                    eDataTable.Add(new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                    eDataGrid.DataContext = eDataTable;
                }
            }

            _workerThreads = new List<Thread>();
            for (int i = 0; i < eDataTable.Count; i++)
            {
                if (eDataTable[i].eTable_Path.Length > 0)
                {
                    Console.WriteLine(i + " " + eDataTable[i].eTable_Path);
                    string mScript_Local = eDataTable[i].eTable_Path;
                    Thread TempThread = new Thread(() =>
                    {
                        try
                        {
                            Script(Load_Script_to_DataTable(eDataTable[i].eTable_Path), "Def");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("{0} Exception caught.", ex);
                        }
                        finally
                        {
                            GC.Collect();
                            GC.WaitForPendingFinalizers();

                        }
                    });

                    _workerThreads.Add(TempThread);
                }
            }
        }

        private void Script_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
           
            eDataGrid.IsEnabled = !eDataGrid.IsEnabled;

            if (eDataGrid.IsEnabled)
            {
                for (int i = 0; i < _workerThreads.Count; i++)
                {
                    if (_workerThreads[i].IsAlive)
                    {
                        _workerThreads[i].Abort();
                        _workerThreads[i] = null;

                        _workerThreads[i] = new Thread(() =>
                        {
                            try
                            {
                                Script(Load_Script_to_DataTable(eDataTable[i].eTable_Path), "Def");
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine("{0} Exception caught.", ex);
                            }
                            finally
                            {
                                GC.Collect();
                                GC.WaitForPendingFinalizers();
                            }
                        });
                    }
                    eDataTable[i].eTable_State = "Stop";
                }
                eDataGrid.DataContext = null;
                eDataGrid.DataContext = eDataTable;

                Btn_ON.Content = "OFF";
                Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
            }
            else {
                Save_Script();
                LoadScriptSetting();

                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            }
        }
        #endregion

        #region Script Panel DataGrid Event
        private void eDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Save_Script();

            if (e.Column.Header.Equals("Path"))
            {
                int RowIndex = e.Row.GetIndex();
            }
        }

        private void eDataGrid_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key.ToString().Equals("Return"))
            {
                // CommitEdit & Change Focus
                eDataGrid.CommitEdit();
                ScriptGrid.Focus();
            }
        }

        private void eDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int columnIndex = eDataGrid.Columns.IndexOf(eDataGrid.CurrentCell.Column);
            if (columnIndex < 0) { return; }

            if (eDataGrid.Columns[columnIndex].Header.ToString().Equals(" "))
            {
                int tableIndex = eDataGrid.Items.IndexOf(eDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < eDataTable.Count())
                    {
                        eDataGrid.DataContext = null;
                        eDataTable.RemoveAt(tableIndex);
                        eDataGrid.DataContext = eDataTable;
                    }
                }
                catch (Exception err)
                {
                    eDataGrid.DataContext = eDataTable;
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
            else if (eDataGrid.Columns[columnIndex].Header.ToString().Equals("+"))
            {
                // Get index
                int tableIndex = eDataGrid.Items.IndexOf(eDataGrid.CurrentItem);

                try
                {
                    if (tableIndex < eDataTable.Count() - 1)
                    {
                        // Insert Item
                        eDataGrid.DataContext = null;
                        eDataTable.Insert(tableIndex + 1, new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                    else
                    {
                        eDataGrid.DataContext = null;
                        eDataTable.Add(new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
        }

        private void eDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            Save_Script();
        }
        #endregion

        #region Edit Panel

        private Thread mThread = null;
        // data
        private Bitmap TempBitmap;

        private void Btn_Toggle_Click(object sender, RoutedEventArgs e)
        {
            ClearScreen_Btn.Focus();

            if (Btn_Toggle.IsOn == true)
            {
                Subscribe();
            }
            else
            {
                Unsubscribe();
            }
        }
        private void Run_script()
        {
            if (Ring.IsActive == true)
            {
                mThread.Abort();
            }

            Ring.IsActive = true;

            mThread = new Thread(() =>
            {
                Script(mDataTable, "Debug");
            });
            mThread.Start();
        }
        private void Stop_script()
        {
            if (mThread != null)
            {
                mThread.Abort();
            }
            Ring.IsActive = false;
        }

        private void ClearScreen_Click(object sender, RoutedEventArgs e)
        {
            gfx.BeginScene(); // call before you start any drawing
            gfx.ClearScene();
            gfx.EndScene();
        }

        private void Btn_New_Click(object sender, RoutedEventArgs e)
        {
            // Table Clear
            mDataGrid.DataContext = null;
            mDataTable.Clear();
            mDataTable.Add(new MainTable() { Enable = true, Mode = "", Action = "", Event = "", Note = "" });
            mDataGrid.DataContext = mDataTable;

            ScriptName.Text = "";

            // .ini
            var parser = new FileIniDataParser();
            IniData data = parser.ReadFile("user.ini");
            data["Def"]["Script"] = "";
            parser.WriteFile("user.ini", data);
        }
        private void Btn_open_Click(object sender, RoutedEventArgs e)
        {
            string fileContent = string.Empty;
            string filePath = string.Empty;

            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
            openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
            openFileDialog.Filter = "txt files (*.txt)|*.txt"; // "txt files (*.txt)|*.txt|All files (*.*)|*.*"
            openFileDialog.FilterIndex = 2;
            openFileDialog.RestoreDirectory = true;
            openFileDialog.ShowDialog();
           
            try
            {
                // ScrollToTop
                if (mDataGrid.Items.Count > 0)
                {
                    var border = VisualTreeHelper.GetChild(mDataGrid, 0) as Decorator;
                    if (border != null)
                    {
                        var scroll = border.Child as ScrollViewer;
                        if (scroll != null) scroll.ScrollToTop();
                    }
                }

                //Get the path of specified file
                filePath = openFileDialog.FileName;
                if (filePath.Equals(""))
                    return;

                Load_Script(filePath);
            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err);
            }

        }
        private async void Btn_Save_as_Click(object sender, RoutedEventArgs e) // async
        {
            Btn_ON.Content = "OFF";
            Btn_ON.Foreground = System.Windows.Media.Brushes.Red;

            var result = await this.ShowInputAsync(FindResource("Save").ToString(), FindResource("Input_filename").ToString());

            Btn_ON.Content = "ON";
            Btn_ON.Foreground = System.Windows.Media.Brushes.White;

            if (result == null) { return; }

            try
            {
                string JSON_String = JsonSerializer.Serialize(mDataTable);
                JSON_String = JSON_String.Insert(1, "\n");
                JSON_String = JSON_String.Insert(JSON_String.Length - 1, "\n");
                JSON_String = JSON_String.Replace("\"},", "\"},\n");

                //string out_string = "";
                //for (int i = 0; i < mDataTable.Count; i++)
                //{
                //    out_string += mDataTable[i].Enable.ToString() + ";"
                //        + mDataTable[i].Mode + ";"
                //        + mDataTable[i].Action + ";"
                //        + mDataTable[i].Event.ToString() + ";"
                //        + mDataTable[i].Note.ToString().Replace(";", "") + ";"
                //        + "\n";
                //}
                System.IO.File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/" + result + ".txt", JSON_String);

                Load_Script(System.Windows.Forms.Application.StartupPath + "\\" + result + ".txt");
            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err);

                await this.ShowMessageAsync("", FindResource("Save_could_not_be_completed").ToString());
            }
        }
        private void Btn_Save_Click(object sender, RoutedEventArgs e)
        {
            string result = null;

            // Get Script Path
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            if (data["Def"]["Script"] != null || data["Def"]["Script"] != "")
            {
                result = data["Def"]["Script"];
            }

            if (result == null || result == "") {
                this.ShowMessageAsync("", FindResource("File_does_not_exist").ToString());
                return;
            }
            try
            {
                string JSON_String = JsonSerializer.Serialize(mDataTable);
                JSON_String = JSON_String.Insert(1, "\n");
                JSON_String = JSON_String.Insert(JSON_String.Length-1, "\n");
                JSON_String = JSON_String.Replace("\"},", "\"},\n");

                //string out_string = "";
                //for (int i = 0; i < mDataTable.Count; i++)
                //{
                //    out_string += mDataTable[i].Enable.ToString() + ";"
                //        + mDataTable[i].Mode + ";"
                //        + mDataTable[i].Action + ";"
                //        + mDataTable[i].Event.ToString() + ";"
                //        + mDataTable[i].Note.ToString().Replace(";", "") + ";"
                //        + "\n";
                //}
                System.IO.File.WriteAllText(result, JSON_String);

                this.ShowMessageAsync("", FindResource("Done").ToString());
            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err);

                this.ShowMessageAsync("", FindResource("Save_could_not_be_completed").ToString());
            }
        }
        private void Btn_Run_Click(object sender, RoutedEventArgs ee)
        {
            Btn_ON.Content = "ON";
            Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            Run_script();
        }
        private void Btn_Stop_Click(object sender, RoutedEventArgs ee)
        {
            Stop_script();
        }


        #endregion

        #region Edit Panel DataGrid Event
        private void DataGridCell_Selected(object sender, RoutedEventArgs e)
        {
            // Lookup for the source to be DataGridCell
            if (e.OriginalSource.GetType() == typeof(System.Windows.Controls.DataGridCell))
            {
                // Starts the Edit on the row;
                System.Windows.Controls.DataGrid grd = (System.Windows.Controls.DataGrid)sender;
                grd.BeginEdit(e);
            }
        }
        private void mDataGrid_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key.ToString().Equals("Return"))
            {
                ToolBar.Visibility = Visibility.Collapsed;
                // CommitEdit & Change Focus
                mDataGrid.CommitEdit();
                EditGrid.Focus();
                
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            }
        }
        private void mDataGrid_HeaderClick(object sender, RoutedEventArgs e)
        {       
            mDataGrid.DataContext = null;
            mDataTable.Add(new MainTable() { Enable = true, Mode = "", Action = "", Event = "", Note = "" });
            mDataGrid.DataContext = mDataTable;
        }

        private void mDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            Btn_ON.Content = "OFF";
            Btn_ON.Foreground = System.Windows.Media.Brushes.Red;

            ToolBar.Items.Clear();

            // ToolBar
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Action"))
            {
                MainTable row = (MainTable)mDataGrid.CurrentItem;
                string[] btnlist = new string[] { };

                switch (row.Mode)
                {
                    case "Click":
                        btnlist = new string[] { "Left", "Right", "Left_Down", "Left_Up", "Right_Down", "Right_Up" };
                        break;
                    case "RemoveEvent":
                        btnlist = new string[] { "PUSH" };
                        break;
                }

                if (btnlist.Length <= 0 || row.Action.Length > 0)
                {
                    ToolBar.Visibility = Visibility.Collapsed;
                    return;
                }

                for (int i = 0; i < btnlist.Length; i++)
                {
                    System.Windows.Controls.Button btn = new System.Windows.Controls.Button();
                    btn.Height = 30;
                    btn.Content = btnlist[i];
                    btn.Background = new SolidColorBrush(Colors.LightGray);
                    btn.Foreground = new SolidColorBrush(Colors.BlueViolet);
                    btn.Click += new RoutedEventHandler(ToolbarBtn_Click);

                    ToolBar.Items.Add(btn);
                }

                DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
                FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                if (cellContent != null)
                {
                    System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;
                    if (cell != null)
                    {
                        System.Windows.Point screenCoordinates = cell.TransformToAncestor(EditGrid).Transform(new System.Windows.Point(0, 0));
                        ToolBar.Margin = new Thickness(0, screenCoordinates.Y - 30, screenCoordinates.X, 0);
                    }
                }

                ToolBar.Visibility = Visibility.Visible;
            }
        }

        private void ToolbarBtn_Click(object sender, RoutedEventArgs e)
        {
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            if (columnIndex < 0) { return; }

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Action"))
            {
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < mDataTable.Count())
                    {
                        DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
                        FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                        System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;
                        System.Windows.Controls.Button mButton = (System.Windows.Controls.Button)sender;
                        System.Windows.Controls.TextBox mTextBlock = (System.Windows.Controls.TextBox)cell.Content;

                        mTextBlock.Text = mTextBlock.Text + mButton.Content;
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
                finally {
                    ToolBar.Visibility = Visibility.Collapsed;
                }
            }
        }

        private void ComboBox_Loaded(object sender, RoutedEventArgs e)
        {
            ((System.Windows.Controls.ComboBox)sender).DropDownClosed += new EventHandler(ComboBox_DropDownClosed); ;
        }

        private void ComboBox_DropDownClosed(object sender, EventArgs e)
        {
            DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
            FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
            System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;

            // CommitEdit & Change Focus
            cell.IsEditing = false;

            mDataGrid.CommitEdit();
            EditGrid.Focus();

            Btn_ON.Content = "ON";
            Btn_ON.Foreground = System.Windows.Media.Brushes.White;
        }

        private void mDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void mDataGrid_ScrollChanged(object sender, ScrollChangedEventArgs e)
        {
            if (e.VerticalChange != 0)
            {
                ToolBar.Visibility = Visibility.Collapsed;
            }
        }

        private void mDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            if (columnIndex < 0) { return; }

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals(" ") || mDataGrid.Columns[columnIndex].Header.ToString().Equals("+")){
                DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
                FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                if (cellContent != null)
                {
                    System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;

                    // CommitEdit & Change Focus
                    cell.IsEditing = false;
                }

                mDataGrid.CommitEdit();
                EditGrid.Focus();
            }

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals(" "))
            {
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < mDataTable.Count())
                    {
                        mDataGrid.DataContext = null;
                        mDataTable.RemoveAt(tableIndex);
                        mDataGrid.DataContext = mDataTable;

                        // Automatically turned on
                        Btn_ON.Content = "ON";
                        Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                    }
                }
                catch (Exception err)
                {
                    mDataGrid.DataContext = mDataTable;
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
            else if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("+"))
            {
                // Get index
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);

                try
                {
                    if (tableIndex < mDataTable.Count() - 1)
                    {
                        // Insert Item
                        mDataGrid.DataContext = null;
                        mDataTable.Insert(tableIndex + 1, new MainTable() { Enable = true, Mode = "", Action = "", Event = "" , Note = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                    else
                    {
                        mDataGrid.DataContext = null;
                        mDataTable.Add(new MainTable() { Enable = true, Mode = "", Action = "", Event = "", Note = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
        }
        #endregion

        #region Setting Panel
        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            System.Windows.Documents.Hyperlink mHyperlink = (System.Windows.Documents.Hyperlink)sender;
            Process.Start(mHyperlink.NavigateUri.ToString());
        }
        private void TextBox_OnOff_Hotkey_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            TextBox_OnOff_Hotkey.Text = "";
        }

        private void TextBox_Run_Hotkey_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            TextBox_Run_Hotkey.Text = "";
        }

        private void TextBox_Stop_Hotkey_KeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
            TextBox_Stop_Hotkey.Text = "";
        }

        private void OnOff_CrtlKey_Chk_Checked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.OnOff_CrtlKey = "1";
        }

        private void OnOff_CrtlKey_Chk_Unchecked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.OnOff_CrtlKey = "0";
        }

        private void TextBox_OnOff_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            mSettingHelper.OnOff_Hotkey = TextBox_OnOff_Hotkey.Text.ToString();
        }

        private void Run_CrtlKey_Chk_Checked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.Run_CrtlKey = "1";
        }

        private void Run_CrtlKey_Chk_Unchecked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.Run_CrtlKey = "0";
        }

        private void Stop_CrtlKey_Chk_Checked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.Stop_CrtlKey = "1";
        }

        private void Stop_CrtlKey_Chk_Unchecked(object sender, RoutedEventArgs e)
        {
            mSettingHelper.Stop_CrtlKey = "0";
        }

        private void TextBox_Run_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            mSettingHelper.Run_Hotkey = TextBox_Run_Hotkey.Text.ToString();
        }

        private void TextBox_Stop_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            mSettingHelper.Stop_Hotkey = TextBox_Stop_Hotkey.Text.ToString();
        }

        private void HideOnSatrt_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (HideOnSatrt_Toggle.IsOn == true)
            {
                mSettingHelper.HideOnSatrt = "1";
            }
            else
            {
                mSettingHelper.HideOnSatrt = "0";
            }
        }

        private void TestMode_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (TestMode_Toggle.IsOn == true)
            {
                TestMode = true;
                mSettingHelper.TestMode = "1";
            }
            else
            {
                TestMode = false;
                mSettingHelper.TestMode = "0";
            }
        }

        private void LangSplitButton_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            MahApps.Metro.Controls.SplitButton mSplitButton = (MahApps.Metro.Controls.SplitButton)sender;
            System.Windows.Controls.Label ItemLabel = (System.Windows.Controls.Label)mSplitButton.SelectedItem;
            string SelectVal = ItemLabel.Tag.ToString();

            mSettingHelper.Language = SelectVal;
            SetLang(SelectVal);
        }

        private void SetLang(string lang)
        {
            ResourceDictionary dict = new ResourceDictionary();
            dict.Source = new Uri(@"..\Resources\StringResources." + lang + ".xaml", UriKind.Relative);
            System.Windows.Application.Current.Resources.MergedDictionaries.Add(dict);

            mFrame.Navigate(new System.Uri(@"..\Resources\ModeDocument_"+ lang.Replace("-","") + ".xaml", UriKind.RelativeOrAbsolute));
        }

        private string CheckUpdate()
        {
            try
            {
                const string GITHUB_API = "https://api.github.com/repos/{0}/{1}/releases/latest";
                WebClient mWebClient = new WebClient();
                mWebClient.Headers.Add("User-Agent", "Unity web player");
                Uri uri = new Uri(string.Format(GITHUB_API, "gemilepus", "Scriptboxie"));
                string releases = mWebClient.DownloadString(uri);
                Console.WriteLine(releases);

                var deserialize = JsonSerializer.Deserialize<Dictionary<string, object>>(releases);
                deserialize.TryGetValue("tag_name", out var tag_name);

                return tag_name.ToString();
            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err);
                return "";
            }
        }

        private async void Updates_Button_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string tag_name = CheckUpdate();

                System.Reflection.Assembly assembly = System.Reflection.Assembly.GetExecutingAssembly();
                if (tag_name.ToString().Equals("v" + System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion))
                {
                    await this.ShowMessageAsync("", string.Format(FindResource("The_latest_version_is_used").ToString(), tag_name.ToString()));
                }
                else
                {
                    if (!tag_name.ToString().Equals("")) {
                        await this.ShowMessageAsync("", string.Format(FindResource("The_current_version_is").ToString(), tag_name.ToString()));
                        Process.Start("https://github.com/gemilepus/Scriptboxie/releases");
                    }
                }
            }
            catch (Exception err)
            {
                Console.WriteLine("{0} Exception caught.", err);
            }
          
        }
        #endregion

    }
}