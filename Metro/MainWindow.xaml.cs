﻿using System;
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

namespace Metro
{
    public partial class MainWindow : MetroWindow
    {
        // **************************************** Motion ******************************************
        #region Motion
        // key actions
        [DllImport("user32.dll", SetLastError = true)]
        private static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        private const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        private const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        private const int VK_LCONTROL = 0xA2; //Left Control key code
        private const int A = 0x41; //A key code
        private const int C = 0x43; //C key code

        // Mouse move
        [DllImport("user32")]
        private static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        private static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        // Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

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
        #endregion
        // End *********************************************************************************************

        // **************************************** Window Activate ******************************************
        #region Window Activate
        // Set Foreground Window                        
        // Get a handle to an application window.
        [DllImport("USER32.DLL", CharSet = CharSet.Unicode)]
        private static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        // Activate an application window.
        [DllImport("USER32.DLL")]
        private static extern bool SetForegroundWindow(IntPtr hWnd);

        // VkKeyScan Char to 0x00
        [DllImport("user32.dll")]
        private static extern byte VkKeyScan(char ch);
        #endregion
        // End *********************************************************************************************

        // ******************************** PostMessage & SendMessage & FindWindows ********************************
        #region PostMessage & SendMessage & FindWindows 
        // FindWindows EX
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
        // End *********************************************************************************************

        // **************************************** OpenCV & Media ******************************************
        #region OpenCV & Media
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
        // End ***************************************************************************************

        private DependencyProperty UnitIsCProperty = DependencyProperty.Register("IsActive", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));
        private new bool IsActive
        {
            get { return (bool)this.GetValue(UnitIsCProperty); }
            set { this.SetValue(UnitIsCProperty, value); }
        }

        // DataGrid
        private List<MainTable> mDataTable = new List<MainTable>();
        private List<EditTable> eDataTable = new List<EditTable>();

        private List<Thread> _workerThreads = new List<Thread>();

        // Globalmousekeyhook for record
        #region
        private static IKeyboardMouseEvents m_GlobalHook, Main_GlobalHook, Main_GlobalKeyUpHook;

        private int now_x, now_y;
        private void Btn_Toggle_Click(object sender, RoutedEventArgs e)
        {
            if (Btn_Toggle.IsOn == true)
            {
                Subscribe();
            }
            else
            {
                Unsubscribe();

                // remove event
                mDataGrid.DataContext = null;
                mDataTable.RemoveAt(mDataTable.Count - 1);
                mDataTable.RemoveAt(mDataTable.Count - 1);
                mDataGrid.DataContext = mDataTable;
            }
        }
        private void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress += GlobalHookKeyPress;
            m_GlobalHook.MouseMove += HookManager_MouseMove;
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e)
        {
            if (Btn_Toggle.IsOn == true)
            {
                mDataGrid.DataContext = null;
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "SendKeyDown", mTable_Action = e.KeyChar.ToString().ToUpper(), mTable_Event = "" });
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "Delay", mTable_Action = "100", mTable_Event = "" });
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "SendKeyUp", mTable_Action = e.KeyChar.ToString().ToUpper(), mTable_Event = "" });
                mDataGrid.DataContext = mDataTable;
            }
            Console.WriteLine("KeyPress: \t{0}", e.KeyChar);
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e)
        {
            if (Btn_Toggle.IsOn == true)
            {
                //if (e.Button.Equals("")) { }
                mDataGrid.DataContext = null;
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "Move", mTable_Action = now_x.ToString() + "," + now_y.ToString(), mTable_Event = "" });
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "Click", mTable_Action = e.Button.ToString(), mTable_Event = "" });
                mDataGrid.DataContext = mDataTable;
            }
            Console.WriteLine("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp);

            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        private void HookManager_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            now_x = e.X;
            now_y = e.Y;
            //Console.WriteLine("MouseMove: x={0:0000}; y={1:0000}", e.X, e.Y);
        }

        private void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            m_GlobalHook.Dispose();
        }
        #endregion

        // GameOverlay .Net
        private OverlayWindow _window;
        private GameOverlay.Drawing.Graphics _graphics;
        // Brush
        private GameOverlay.Drawing.SolidBrush _red;
        private GameOverlay.Drawing.Font _font;
        private GameOverlay.Drawing.SolidBrush _black;

        private GameOverlay.Drawing.Graphics gfx;

        // NotifyIcon 最小化
        private static NotifyIcon mNotifyIcon = new NotifyIcon();
        //private void MenuItem_Click(object sender, RoutedEventArgs e)
        //{
        //    this.Show();
        //}

        private SettingHelper mSettingHelper = new SettingHelper();

        private string OnOff_Hotkey, Run_Hotkey, Stop_Hotkey;
        public MainWindow()
        {
            InitializeComponent();

            MSG_SHOW = RegisterWindowMessage("Scriptboxie Message");

            Timestamp = (double)DateTime.UtcNow.Subtract(new DateTime(1970, 1, 1)).TotalMilliseconds;

            // Data Binding
            //this.DataContext = this;

            // Combobox List
            List<string> mList = new List<string>() {
                "Move","Offset","Click", "Match","Match RGB","Match&Draw",
                "Key","ModifierKey","SendKeyDown","SendKeyUp","Delay",
                "Jump", "Loop", "RemoveEvent","Clean Draw",
                "Run .exe","PlaySound"
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
                mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "", mTable_Action = "", mTable_Event = "" });
                mDataGrid.DataContext = mDataTable;
            }

            #endregion

            // Load Script setting
            try
            {
                Load_Script_ini();
            }
            catch (Exception ex)
            {
                if (!File.Exists("Script.ini"))
                {
                    eDataGrid.DataContext = null;
                    eDataTable.Add(new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                    eDataGrid.DataContext = eDataTable;
                }
            }
            
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

            // Load setting
            OnOff_Hotkey = data["Def"]["OnOff_Hotkey"];
            TextBox_OnOff_Hotkey.Text = data["Def"]["OnOff_Hotkey"];
            OnOff_Hotkey = data["Def"]["OnOff_Hotkey"];
            TextBox_OnOff_Hotkey.Text = data["Def"]["OnOff_Hotkey"];
            Run_Hotkey = data["Def"]["Run_Hotkey"];
            TextBox_Run_Hotkey.Text = data["Def"]["Run_Hotkey"];
            Stop_Hotkey = data["Def"]["Stop_Hotkey"];
            TextBox_Stop_Hotkey.Text = data["Def"]["Stop_Hotkey"];


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
            mNotifyIcon.Visible = true;
            mNotifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);

            // test
            //ConvertHelper.GetEnumVirtualKeyCodeValues();
        }

        #region APP

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
        private void AlertSound()
        {
            //try
            //{
            //    SoundPlayer mWaveFile = new SoundPlayer("UI211.wav");
            //    mWaveFile.PlaySync();
            //    mWaveFile.Dispose();
            //}
            //catch (Exception e)
            //{
            //    Console.WriteLine(e);
            //}

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
            this.ShowMessageAsync("About", 
                Metro.Properties.Resources.Version + "\n"
                + "Author: " + "gemilepus" + "\n"
                + "Mail: " + "gemilepus@gmail.com" + "\n"
                );
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

        private void Main_GlobalHookKeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {

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
        }

        private string PassCtrlKey = "";
        private double Timestamp;
        private void Main_GlobalHookKeyPress(object sender, System.Windows.Forms.KeyEventArgs e)
        {

            if (!TextBox_Title.Text.Equals(""))
            {
                if (GetActiveWindowTitle() == null) { return; }
                string ActiveTitle = GetActiveWindowTitle();
                if (ActiveTitle.Length == ActiveTitle.Replace(TextBox_Title.Text, "").Length) { return; }
            }

            // stop KListener
            UnKListener();

            // ON / OFF
            if (e.KeyCode.ToString().Equals(OnOff_Hotkey)) // "'"
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
                KListener(); 
                return; }

            Console.WriteLine(e.KeyCode.ToString());


            if (e.KeyCode.ToString().Equals(Run_Hotkey)) //"["
            {
                AlertSound();
                ShowBalloon("Run", "...");
                Run_script();
            }
            if (e.KeyCode.ToString().Equals(Stop_Hotkey)) //"]"
            {
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
            for (int i = 0; i < eDataTable.Count; i++)
            {
                if (KeyCode.Equals(eDataTable[i].eTable_Key) && eDataTable[i].eTable_Enable == true)
                {
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
            if (!eDataGrid.IsKeyboardFocusWithin) {
                eDataGrid.DataContext = null;
                eDataGrid.DataContext = eDataTable;
            }
           
            // Restart KListener
            KListener();
        }
        #endregion

        private void ShowBalloon(string title,string msg) {
            // Hide BalloonTip
            if (mNotifyIcon.Visible)
            {
                mNotifyIcon.Visible = false;
                mNotifyIcon.Visible = true;
            }
            mNotifyIcon.ShowBalloonTip(500, title, msg, ToolTipIcon.None);
        }

        // For resize x,y
        private float ScaleX, ScaleY;
        private int OffsetX, OffsetY;
      
        private void Script(List<MainTable> minDataTable,string Mode)
        {
            InputSimulator mInputSimulator = new InputSimulator();
            Keyboard ky = new Keyboard();
            ConvertHelper ConvertHelper = new ConvertHelper();
            SortedList mDoSortedList = new SortedList();
            V V = new V();
           
            // key || value ex:
            //mDoSortedList.Add("Point", "0,0");
            //mDoSortedList.Add("Point Array", "0,0,0,0");
            //mDoSortedList.Add("Draw", "");
            //mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey("Draw"));

            int n = 0; int LoopCount = 0;
            while (n < minDataTable.Count)
            {
                string Command = minDataTable[n].mTable_Mode;
                string CommandData = minDataTable[n].mTable_Action;
                bool CommandEnable = minDataTable[n].mTable_IsEnable;

                string[] Event = minDataTable[n].mTable_Event.Split(',');
                if (minDataTable[n].mTable_Event == "") { Event = new string[0]; }

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
                                    if (CommandData.Equals("Left"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                        //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                        //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.LeftButton);
                                        mInputSimulator.Mouse.LeftButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.LeftButtonUp();
                                    }
                                    if (CommandData.Equals("Left_Down"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.LeftButtonDown();
                                    }
                                    if (CommandData.Equals("Left_Up"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.LeftButtonUp();
                                    }
                                    if (CommandData.Equals("Right"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                        //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                        //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.RightButton);

                                        mInputSimulator.Mouse.RightButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.RightButtonUp();
                                    }
                                    if (CommandData.Equals("Right_Down"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.RightButtonDown();
                                    }
                                    if (CommandData.Equals("Right_Up"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
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
                                    double threshold = 0.8;

                                    if (MatchArr[0].Equals(""))
                                    {
                                        MatchArr[0] = "example/s.png";
                                    }

                                    if (MatchArr.Length > 2)
                                    {
                                        matTarget = BitmapConverter.ToMat(makeScreenshot_clip(
                                            int.Parse(MatchArr[1]), int.Parse(MatchArr[2]),
                                            int.Parse(MatchArr[3]), int.Parse(MatchArr[4])));

                                        if (MatchArr.Length > 5)
                                        {
                                            threshold = Convert.ToDouble(MatchArr[1]);
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
                                    if (Command.Equals("Match"))
                                    {
                                        return_xy = RunTemplateMatch(matTarget, matTemplate,"GRAY", threshold);
                                    }
                                    else
                                    {
                                        return_xy = RunTemplateMatch(matTarget, matTemplate, "RGB", threshold);
                                    }
                                    //System.Windows.Forms.MessageBox.Show(RunTemplateMatch(matTarget, matTemplate));

                                    if (!return_xy.Equals(""))
                                    {
                                        if (Command.Equals("Match&Draw"))
                                        {
                                            string[] xy = return_xy.Split(',');
                                            //SetCursorPos(int.Parse(xy[0]) + temp_w, int.Parse(xy[1]) + temp_h);

                                            gfx.BeginScene();
                                            //gfx.DrawTextWithBackground(_font, _red, _black, 10, 10, return_xyd.ToString());
                                            gfx.DrawRoundedRectangle(_red, RoundedRectangle.Create(int.Parse(xy[0]), int.Parse(xy[1]), temp_w * 2, temp_h * 2, 6), 2);
                                            gfx.EndScene();
                                        }
                                    
                                        // Add Key
                                        if (Event[0].Length > 0)
                                        {
                                            if (Event[0].IndexOf("-") == -1)
                                            {
                                                if (V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                                {
                                                    mDoSortedList.Add(Event[0], return_xy);
                                                }
                                            }
                                            else
                                            {
                                                if (mDoSortedList.IndexOfKey(Event[0].Replace("-", "")) == -1)
                                                {
                                                    mDoSortedList.Add(Event[0].Replace("-", ""), return_xy);
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
                                //int temp_w = matTemplate.Width/2 , temp_h = matTemplate.Height/2; // center x y

                                //System.Windows.Forms.MessageBox.Show(RunTemplateMatch(matTarget, matTemplate));
                                //string return_xy = RunTemplateMatch(matTarget, matTemplate);
                                //if (!return_xy.Equals("")) {
                                //    string[] xy = return_xy.Split(',');
                                //    SetCursorPos(int.Parse(xy[0]) + temp_w, int.Parse(xy[1]) + temp_h);
                                //}

                                MatchBySift(matTarget_Sift, matTemplate_Sift);


                                break;

                            case "Key":

                                if (Event.Length != 0 && V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                {
                                    break;
                                }

                                //***************** SendKeys *****************
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
                                //***************** SendKeys *****************

                                //***************** InputSimulator *****************
                                string str = CommandData;
                                char[] arr = str.ToCharArray();

                                foreach (char c in arr)
                                {
                                    //mInputSimulator.Keyboard.KeyPress((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));

                                    mInputSimulator.Keyboard.KeyDown((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
                                    Thread.Sleep(100);
                                    mInputSimulator.Keyboard.KeyUp((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
                                }

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
                                    string SendKeyStr = CommandData;
                                    if (SendKeyStr.Length == 1)
                                    {
                                        SendKeyStr = "KEY_" + SendKeyStr.ToUpper();
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
                                    mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey(Event[0]));
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

                            case "Clean Draw":

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
                                        //System.Windows.MessageBox.Show(windowHandle.ToString());

                                        gfx.BeginScene(); // call before you start any drawing
                                                          // Draw
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
                        //Console.WriteLine("{0} Exception caught.", e);
                        Console.WriteLine("Exception caught.");

                        if (e.Message.ToString().IndexOf("Thread") == -1) {
                            if (Mode.Equals("Debug")) {
                                System.Windows.MessageBox.Show("[Error] Line " + (n+1).ToString() + " : " + e.Message);

                                // debug stop msg
                                CreateMessage("9487");

                                break;
                            } 
                        }

                        mInputSimulator = null;
                        ky = null;
                        V = null;
                        ConvertHelper = null;
                        mDoSortedList.Clear();

                        err = true;
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

        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);
        private uint MSG_SHOW;

        protected override void OnSourceInitialized(EventArgs e)
        {
            base.OnSourceInitialized(e);
            System.Windows.Interop.HwndSource source = PresentationSource.FromVisual(this) as System.Windows.Interop.HwndSource;
            source.AddHook(WndProc);
        }

        private IntPtr WndProc(IntPtr hwnd, int msg, IntPtr wParam, IntPtr lParam, ref bool handled)
        {
            //Console.WriteLine(lParam);
            if (msg == MSG_SHOW)
            {
                if (wParam.ToString().Equals("9487")) {
                    Ring.IsActive = false;
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
                    Thread TempThread = new Thread(() =>
                    {
                        Thread.Sleep(100);

                        IntPtr mPrt = FindWindow(null, "Scriptboxie");
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

                if (wParam.ToString().Substring(0,2).Equals("22"))
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
            // Table Clear
            mDataGrid.DataContext = null;
            mDataTable.Clear();

            mDataTable = Load_Script_to_DataTable(filePath);
            mDataGrid.DataContext = mDataTable;
        }

        private List<MainTable> Load_Script_to_DataTable(string mfilePath)
        {
            List<MainTable> tempDataTable = new List<MainTable>();
            string fileContent = string.Empty;

            try
            {
                StreamReader reader = new StreamReader(mfilePath);
                // read test
                fileContent = reader.ReadToEnd();
                fileContent.Replace(";", "%;");
                string[] SplitStr = fileContent.Split(';');

                for (int i = 0; i < SplitStr.Length - 4; i += 4)
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
                        default:
                            break;
                    }

                    tempDataTable.Add(new MainTable()
                    {
                        mTable_IsEnable = bool.Parse(SplitStr[i].Replace("%", "")),
                        mTable_Mode = mMode,
                        mTable_Action = SplitStr[i + 2].Replace("%", ""),
                        mTable_Event = SplitStr[i + 3].Replace("%", ""),
                    });
                }
            }
            catch
            {
                bool fileExist = File.Exists(mfilePath);
                if (fileExist)
                {
                    this.ShowMessageAsync("Load_Script_to_DataTable", "Error!");
                }
                //CreateMessage("2200");
            }
            return tempDataTable;
        }
        private void eDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Save_Script();

            if (e.Column.Header.Equals("Path")) {
                int RowIndex = e.Row.GetIndex();

                // Restart
                //this.Close();
                //Process.Start("Metro.exe", "");
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
                        //Table Clean
                        eDataGrid.DataContext = null;
                        eDataTable.RemoveAt(tableIndex);
                        eDataGrid.DataContext = eDataTable;
                    }
                }
                catch (Exception err)
                {
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

        private void Script_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            eDataGrid.IsEnabled = !eDataGrid.IsEnabled;
            if (eDataGrid.IsEnabled)
            {
                Btn_ON.Content = "OFF";
                Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
            }
            else {
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            }
        }
        #endregion

        #region Edit Panel

        private Thread mThread = null;
        // data
        private Bitmap TempBitmap;
        private void Run_script()
        {
            if (Ring.IsActive == true)
            {
                mThread.Abort();
            }
            //mThread.Abort();
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
                //Get the path of specified file
                filePath = openFileDialog.FileName;
                Load_Script(filePath);

                // .ini
                var parser = new FileIniDataParser();
                IniData data = parser.ReadFile("user.ini");
                data["Def"]["Script"] = filePath;
                parser.WriteFile("user.ini", data);
            }
            catch
            {
                //this.ShowMessageAsync("", "ERROR!");
            }

        }
        private async void Btn_Save_as_Click(object sender, RoutedEventArgs e) // async
        {
            var result = await this.ShowInputAsync("Save", "input filename:");
            if (result == null) { return; }

            string out_string = "";
            for (int i = 0; i < mDataTable.Count; i++)
            {
                out_string += mDataTable[i].mTable_IsEnable.ToString() + ";"
                    + mDataTable[i].mTable_Mode + ";"
                    + mDataTable[i].mTable_Action + ";"
                    + mDataTable[i].mTable_Event.ToString() + ";"
                    + "\n";
            }
            System.IO.File.WriteAllText(System.Windows.Forms.Application.StartupPath + "/" + result + ".txt", out_string);
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

            if (result == null) {
                this.ShowMessageAsync("", "ERROR!");
                return;
            }

            string out_string = "";
            for (int i = 0; i < mDataTable.Count; i++)
            {
                out_string += mDataTable[i].mTable_IsEnable.ToString() + ";"
                    + mDataTable[i].mTable_Mode + ";"
                    + mDataTable[i].mTable_Action + ";"
                    + mDataTable[i].mTable_Event.ToString() + ";"
                    + "\n";
            }
            System.IO.File.WriteAllText(result, out_string);

            // Restart
            this.Close();
            Process.Start("Scriptboxie.exe", "");
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
                // CommitEdit & Change Focus
                mDataGrid.CommitEdit();
                EditGrid.Focus();
                
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            }
        }

        private void mDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            Btn_ON.Content = "OFF";
            Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
        }

        private void mDataGrid_AddingNewItem(object sender, AddingNewItemEventArgs e)
        {
            e.NewItem = new MainTable
            {
                mTable_IsEnable = true,
                mTable_Mode = "",
                mTable_Action = "",
                mTable_Event = ""
            };
        }
        private void mDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex() + 1).ToString();
        }

        private void mDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            if (columnIndex < 0) { return; }

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals(" "))
            {
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < mDataTable.Count())
                    {
                        //Table Clean
                        mDataGrid.DataContext = null;
                        mDataTable.RemoveAt(tableIndex);
                        mDataGrid.DataContext = mDataTable;
                    }
                }
                catch (Exception err)
                {
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
                        mDataTable.Insert(tableIndex + 1, new MainTable() { mTable_IsEnable = true, mTable_Mode = "", mTable_Action = "", mTable_Event = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                    else
                    {
                        mDataGrid.DataContext = null;
                        mDataTable.Add(new MainTable() { mTable_IsEnable = true, mTable_Mode = "", mTable_Action = "", mTable_Event = "" });
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

        private void Hyperlink_RequestNavigate(object sender, RequestNavigateEventArgs e)
        {
            Process.Start("https://github.com/gemilepus/Scriptboxie");
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

        private void TextBox_OnOff_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            data["Def"]["OnOff_Hotkey"] = TextBox_OnOff_Hotkey.Text.ToString();
            parser.WriteFile("user.ini", data);
        }

        private void TextBox_Run_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            data["Def"]["Run_Hotkey"] = TextBox_Run_Hotkey.Text.ToString();
            parser.WriteFile("user.ini", data);
        }

        private void TextBox_Stop_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            data["Def"]["Stop_Hotkey"] = TextBox_Stop_Hotkey.Text.ToString();
            parser.WriteFile("user.ini", data);
        }

    }
}