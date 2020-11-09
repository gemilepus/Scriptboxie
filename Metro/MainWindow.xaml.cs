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
using OpenCvSharp.XFeatures2D;

using MahApps.Metro.Controls;
using MahApps.Metro.Controls.Dialogs;

using Ownskit.Utils;

using GameOverlay.Drawing;
using GameOverlay.Windows;

using Gma.System.MouseKeyHook;

using WindowsInputLib;
using WindowsInputLib.Native;

using IniParser;
using IniParser.Model;

namespace Metro
{
    public partial class MainWindow : MetroWindow
    {
        // **************************************** Motion ******************************************
        #region Motion
        // key actions
        [DllImport("user32.dll", SetLastError = true)]
        static extern void keybd_event(byte bVk, byte bScan, int dwFlags, int dwExtraInfo);
        public const int KEYEVENTF_EXTENDEDKEY = 0x0001; //Key down flag
        public const int KEYEVENTF_KEYUP = 0x0002; //Key up flag
        public const int VK_LCONTROL = 0xA2; //Left Control key code
        public const int A = 0x41; //A key code
        public const int C = 0x43; //C key code

        // Mouse move
        [DllImport("user32")]
        public static extern int SetCursorPos(int x, int y);

        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        public static extern void mouse_event(uint dwFlags, uint dx, uint dy, uint cButtons, uint dwExtraInfo);
        // Mouse actions
        private const int MOUSEEVENTF_LEFTDOWN = 0x02;
        private const int MOUSEEVENTF_LEFTUP = 0x04;
        private const int MOUSEEVENTF_RIGHTDOWN = 0x08;
        private const int MOUSEEVENTF_RIGHTUP = 0x10;

        // GetActiveWindowTitle
        [DllImport("user32.dll")]
        static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

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
        public static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        // Activate an application window.
        [DllImport("USER32.DLL")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        // VkKeyScan Char to 0x00
        [DllImport("user32.dll")]
        static extern byte VkKeyScan(char ch);
        #endregion
        // End *********************************************************************************************

        // ******************************** PostMessage & SendMessage & FindWindows ********************************
        #region PostMessage & SendMessage & FindWindows 
        // FindWindows EX
        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        // PostMessageA
        [DllImport("User32.Dll", EntryPoint = "PostMessageA")]
        private static extern bool PostMessage(IntPtr hWnd, uint msg, int wParam, int lParam);

        const int WM_KEYDOWN = 0x100;
        const int WM_KEYUP = 0x101;
        const int WM_CHAR = 0x0102;

        //const int WM_LBUTTONDOWN = 0x201;
        //const int WM_LBUTTONUP = 0x202;

        // SendMessage
        [DllImport("user32.dll")]
        public static extern int SendMessage(int hWnd, int msg, int wParam, [MarshalAs(UnmanagedType.LPStr)] string lParam);
        #endregion
        // End *********************************************************************************************

        // **************************************** OpenCV & Media ******************************************
        #region OpenCV & Media
        public static Bitmap makeScreenshot()
        {
            Bitmap screenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);

            System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);

            gfxScreenshot.Dispose();

            return screenshot;
        }

        public static Bitmap makeScreenshot_clip(int x, int y, int height, int width)
        {
            Bitmap screenshot = new Bitmap(width, height);

            System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);

            gfxScreenshot.CopyFromScreen(x, y, 0 , 0 , screenshot.Size);

            gfxScreenshot.Dispose();

            return screenshot;
        }


        //  Multiple results
        private String RunTemplateMatch_GRAY(Mat rec, Mat template)
        {
            string reText = "";

            using (Mat refMat = rec)
            using (Mat tplMat = template)
            using (Mat res = new Mat(refMat.Rows - tplMat.Rows + 1, refMat.Cols - tplMat.Cols + 1, MatType.CV_32FC1))
            {
                //Convert input images to gray
                Mat gref = refMat.CvtColor(ColorConversionCodes.BGR2GRAY);
                Mat gtpl = tplMat.CvtColor(ColorConversionCodes.BGR2GRAY);

                Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.CCoeffNormed);
                Cv2.Threshold(res, res, 0.8, 1.0, ThresholdTypes.Tozero);

                while (true)
                {
                    double minval, maxval, threshold = 0.8;
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

                        reText = reText + maxloc.X.ToString() + "," + maxloc.Y.ToString() + ",";
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

                return reText;
            }
        }

        private String RunTemplateMatch(Mat rec, Mat template)
        {
            string reText = "";

            using (Mat refMat = rec)
            using (Mat tplMat = template)
            using (Mat res = new Mat(refMat.Rows - tplMat.Rows + 1, refMat.Cols - tplMat.Cols + 1, MatType.CV_32FC1))
            {
                Mat gref = refMat.CvtColor(ColorConversionCodes.BGR2HLS);
                Mat gtpl = tplMat.CvtColor(ColorConversionCodes.BGR2HLS);
                Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.CCoeffNormed);
                //Cv2.MatchTemplate(gref, gtpl, res, TemplateMatchModes.SqDiffNormed);
                Cv2.Threshold(res, res, 0.8, 1.0, ThresholdTypes.Tozero);

                while (true)
                {
                    double minval, maxval, threshold = 0.8;
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

                        reText = reText + maxloc.X.ToString() + "," + maxloc.Y.ToString() + ",";
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

                return reText;
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
                OpenCvSharp.Rect[] faces = cascade.DetectMultiScale(gray, 1.08, 2, HaarDetectionType.ScaleImage, new OpenCvSharp.Size(30, 30));
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

        public DependencyProperty UnitIsCProperty = DependencyProperty.Register("IsActive", typeof(bool), typeof(MainWindow), new PropertyMetadata(false));
        public new bool IsActive
        {
            get { return (bool)this.GetValue(UnitIsCProperty); }
            set { this.SetValue(UnitIsCProperty, value); }
        }

        // WindowsInputLibrary
        InputSimulator mInputSimulator = new InputSimulator();

        // globalmousekeyhook
        private IKeyboardMouseEvents m_GlobalHook;
        private int now_x, now_y;
        private void Btn_Toggle_Click(object sender, RoutedEventArgs e)
        {
            if (Btn_Toggle.IsOn == true){
                Subscribe();
            }
            else{
                Unsubscribe();

                // remove event
                mDataGrid.DataContext = null;
                mDataTable.RemoveAt(mDataTable.Count - 1);
                mDataTable.RemoveAt(mDataTable.Count - 1);
                mDataGrid.DataContext = mDataTable;
            }
        }
        public void Subscribe(){
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress += GlobalHookKeyPress;
            m_GlobalHook.MouseMove += HookManager_MouseMove;
        }

        private void GlobalHookKeyPress(object sender, KeyPressEventArgs e){
            Console.WriteLine("KeyPress: \t{0}", e.KeyChar);
        }

        private void GlobalHookMouseDownExt(object sender, MouseEventExtArgs e){
            if (Btn_Toggle.IsOn == true){
                //if (e.Button.Equals("")) { }
                mDataGrid.DataContext = null;
                mDataTable.Add(new mTable() { mTable_IsEnable = true, mTable_Mode = "Move", mTable_Action = now_x.ToString() +","+ now_y.ToString(), mTable_Event = "" });
                mDataTable.Add(new mTable() { mTable_IsEnable = true, mTable_Mode = "Click", mTable_Action = e.Button.ToString(), mTable_Event = "" });
                mDataGrid.DataContext = mDataTable;
            }
            Console.WriteLine("MouseDown: \t{0}; \t System Timestamp: \t{1}", e.Button, e.Timestamp);

            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        private void HookManager_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e){
            now_x = e.X;
            now_y = e.Y;
            //Console.WriteLine("MouseMove: x={0:0000}; y={1:0000}", e.X, e.Y);
        }

        public void Unsubscribe(){
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.KeyPress -= GlobalHookKeyPress;

            //It is recommened to dispose it
            m_GlobalHook.Dispose();
        }
        
        KeyboardListener KListener = new KeyboardListener();

        // DataGrid
        List<mTable> mDataTable = new List<mTable>();
        List<eTable> eDataTable = new List<eTable>();

        private List<Thread> _workerThreads = new List<Thread>();

        public MainWindow()
        {
            InitializeComponent();

            KListener.KeyDown += new RawKeyEventHandler(KListener_KeyDown);
            // Data Binding
            //this.DataContext = this;

            // Combobox List
            List<string> mList = new List<string>() {
                "Move","Offset", "Loop", "Click", "Match", "Key","RemoveKey",
                "Delay", "Get Point","Run exe", "FindWindow","ScreenClip", 
                "Draw", "Sift Match",  "Clean Draw", "PostMessage", "PlaySound",
                "Color Test"
            };
            mComboBoxColumn.ItemsSource = mList;

            //mDataGrid.DataContext = mDataTable;
            //eDataTable.Add(new eTable() { eTable_Enable = true, eTable_Name = "Run", eTable_Key = "R", eTable_State = "", eTable_Note = "", eTable_Path = @"E:\Script_Lite\MoonyDesk\bin\Debug\Do.txt" });
            //eDataGrid.DataContext = eDataTable;

            // user.ini
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            try
            {
                data = parser.ReadFile("user.ini");
            }
            catch
            {
                parser.WriteFile("user.ini", new IniData());
            }

            // From location
            if (data["Def"]["x"] != null)
            {
                Left = double.Parse(data["Def"]["x"]);
                Top = double.Parse(data["Def"]["y"]);
            }

            // Load WindowTitle setting
            if (data["Def"]["WindowTitle"] != null)
            {
                TextBox_Title.Text = data["Def"]["WindowTitle"];
            }

            // Load Script
            if (data["Def"]["Script"] != null || data["Def"]["Script"] != "") {
                try {
                    Load_Script(data["Def"]["Script"]);
                } 
                catch {
                    data["Def"]["Script"] = "";
                    parser.WriteFile("user.ini", data);
                }
            }

            Load_Script_ini();

            for (int i = 0; i < eDataTable.Count; i++){
                if (eDataTable[i].eTable_Path.Length > 0){
                    Console.WriteLine(i + " " + eDataTable[i].eTable_Path);
                    string mScript_Local = eDataTable[i].eTable_Path;
                    Thread TempThread = new Thread(() =>{  
                        Script(Load_Script_to_DataTable(mScript_Local));
                    });
                    _workerThreads.Add(TempThread);
                }
            }
        }
        void AlertSound() {
            try
            {
                SoundPlayer mWaveFile = new SoundPlayer("UI211.wav");
                mWaveFile.PlaySync();
                mWaveFile.Dispose();
            }
            catch (InvalidCastException e)
            {
                Console.WriteLine(e);
            }

        }
        void KListener_KeyDown(object sender, RawKeyEventArgs args)
        {
            if (!TextBox_Title.Text.Equals("")) {
                if (GetActiveWindowTitle() == null){
                    return;
                }
                string ActiveTitle = GetActiveWindowTitle();
                if (ActiveTitle.Length == ActiveTitle.Replace(TextBox_Title.Text, "").Length)
                {
                    return;
                }
            }
           
            // ON / OFF
            if (args.ToString().Equals("'"))
            {
                if (Btn_ON.Content.Equals("ON"))
                {
                    Btn_ON.Content = "OFF";
                }
                else
                {
                    Btn_ON.Content = "ON";
                }
            }
            if (!Btn_ON.Content.Equals("ON")) {
                return;
            }

            KListener.Dispose();
           
            if (args.ToString().Equals("[")){
                AlertSound();
                Run_script();
            }
            if (args.ToString().Equals("]")){
                Stop_script();
            }

            // Select Script
            for (int i = 0; i < eDataTable.Count; i++){
                if (args.Key.ToString().Equals(eDataTable[i].eTable_Key) && eDataTable[i].eTable_Enable == true)
                {
                    Console.WriteLine("START " + _workerThreads[i].ThreadState.ToString());
                    //if (_workerThreads[i].ThreadState == System.Threading.ThreadState.WaitSleepJoin){
                    //    break;
                    //}
                    if (!_workerThreads[i].IsAlive){
                        if (_workerThreads[i].ThreadState != System.Threading.ThreadState.Stopped)
                        {
                            _workerThreads[i].Start();
                        }
                        else
                        {
                            Console.WriteLine(eDataTable[i].eTable_Path);
                            string mScript_Local = eDataTable[i].eTable_Path;
                            List<mTable> Script_DataTable = Load_Script_to_DataTable(mScript_Local);
                            Thread TempThread = new Thread(() =>{                                  
                                try
                                {
                                    Script(Script_DataTable);
                                }
                                catch (Exception e)
                                {
                                    Console.WriteLine("{0} Exception caught.", e);
                                }
                            });
                            _workerThreads[i] = TempThread;
                            _workerThreads[i].Start();
                        }
                        eDataTable[i].eTable_State = "Running";
                        Console.WriteLine(_workerThreads[i].ThreadState.ToString());
                    }
                    else
                    { 
                        _workerThreads[i].Abort();
                        Console.WriteLine(eDataTable[i].eTable_Path);
                        string mScript_Local = eDataTable[i].eTable_Path;
                        List<mTable> Script_DataTable = Load_Script_to_DataTable(mScript_Local);
                        Thread TempThread = new Thread(() =>{
                            try
                            {
                                Script(Script_DataTable);
                            }
                            catch (Exception e)
                            {
                                Console.WriteLine("{0} Exception caught.", e);
                            }
                        });
                        _workerThreads[i] = TempThread;
                        eDataTable[i].eTable_State = "Stop";
                        Console.WriteLine(_workerThreads[i].ThreadState.ToString());
                    }
                    eDataGrid.DataContext = null;
                    eDataGrid.DataContext = eDataTable;
                }
            }

            // Restart
            KListener = new KeyboardListener();
            KListener.KeyDown += new RawKeyEventHandler(KListener_KeyDown);

            Console.WriteLine(args.Key.ToString());
            // Prints the text of pressed button, takes in account big and small letters. E.g. "Shift+a" => "A"
            Console.WriteLine(args.ToString());
        }

        private void Script(List<mTable> minDataTable)
        {
            SortedList mDoSortedList = new SortedList();
            // key || value
            //mDoSortedList.Add("Point", "0,0");
            //mDoSortedList.Add("Point Array", "0,0,0,0");
            //mDoSortedList.Add("Draw", "");
            //mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey("Draw"));

            //  GameOverlay .Net
            OverlayWindow _window;
            GameOverlay.Drawing.Graphics _graphics;

            // Brush
            GameOverlay.Drawing.SolidBrush _red;
            GameOverlay.Drawing.Font _font;
            GameOverlay.Drawing.SolidBrush _black;

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
                                                                              // creates a simple font with no additional style
            _font = _graphics.CreateFont("Arial", 25);
            _black = _graphics.CreateSolidBrush(GameOverlay.Drawing.Color.Transparent);

            var gfx = _graphics; // little shortcut

            int n = 0;
            while (n < minDataTable.Count)
            {
                string Command = minDataTable[n].mTable_Mode;
                string CommandData = minDataTable[n].mTable_Action;
                bool CommandEnable = minDataTable[n].mTable_IsEnable;
                string[] Event = minDataTable[n].mTable_Event.Split(',');
                if (minDataTable[n].mTable_Event == "") { Event = new string[0];}

                if (CommandEnable) {

                    #region Switch Command
                    switch (Command)
                    {
                        case "Move":

                            if (Event.Length == 0)
                            {
                                string[] str_move = CommandData.Split(',');
                                SetCursorPos(int.Parse(str_move[0]), int.Parse(str_move[1]));
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
                                    else {
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
                                SetCursorPos(point.X + int.Parse(mOffset[0]), point.Y + int.Parse(mOffset[1]));
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
                                    System.Drawing.Point point = System.Windows.Forms.Control.MousePosition;
                                    SetCursorPos(point.X + int.Parse(Event_Data[0]), point.Y + int.Parse(Event_Data[1]));
                                }
                            }

                            break;

                        case "Delay":

                            if (Event.Length == 0)
                            {
                                Thread.Sleep(Int32.Parse(CommandData));
                            }
                            else
                            {
                                // Check Key
                                if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                {
                                    Thread.Sleep(Int32.Parse(CommandData));
                                }
                            }
                            break;

                        case "Click":

                            if (Event.Length == 0)
                            {
                                if (CommandData.Equals("Left"))
                                {
                                    //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                    //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                    //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.LeftButton);

                                    mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.LeftButton);
                                    Thread.Sleep(200);
                                    mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.LeftButton);
                                }
                                if (CommandData.Equals("Left_Down"))
                                {
                                    //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                    mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.LeftButton);
                                }
                                if (CommandData.Equals("Left_Up"))
                                {
                                    //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                    mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.LeftButton);
                                }
                                if (CommandData.Equals("Right"))
                                {
                                    //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                    //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                    //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.RightButton);

                                    mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.RightButton);
                                    Thread.Sleep(200);
                                    mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.RightButton);
                                }
                                if (CommandData.Equals("Right_Down"))
                                {
                                    //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                    mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.RightButton);
                                }
                                if (CommandData.Equals("Right_Up"))
                                {
                                    //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                    mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.RightButton);
                                }
                            }
                            else
                            {
                                // Check Key
                                if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                                {
                                    if (CommandData.Equals("Left"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                        //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                        //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.LeftButton);

                                        mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.LeftButton);
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.LeftButton);
                                    }
                                    if (CommandData.Equals("Left_Down"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.LeftButton);
                                    }
                                    if (CommandData.Equals("Left_Up"))
                                    {
                                        //mouse_event(MOUSEEVENTF_LEFTUP, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.LeftButton);
                                    }
                                    if (CommandData.Equals("Right"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                        //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                        //mInputSimulator.Mouse.MouseButtonClick(WindowsInputLib.MouseButton.RightButton);


                                        mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.RightButton);
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.RightButton);
                                    }
                                    if (CommandData.Equals("Right_Down"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTDOWN, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.MouseButtonDown(WindowsInputLib.MouseButton.RightButton);
                                    }
                                    if (CommandData.Equals("Right_Up"))
                                    {
                                        //mouse_event(MOUSEEVENTF_RIGHTUP, 0, 0, 0, 0);
                                        mInputSimulator.Mouse.MouseButtonUp(WindowsInputLib.MouseButton.RightButton);
                                    }
                                }
                            }

                            break;

                        case "Match":
                            do
                            {
                                string TempPath = CommandData;
                                Mat matTarget;
                                if (TempPath.Equals(""))
                                {
                                    TempPath = "s.png";
                                }

                                if (TempPath.IndexOf(',') != -1)
                                {
                                    string[] mSize = TempPath.Split(',');
                                    matTarget = BitmapConverter.ToMat(makeScreenshot_clip(int.Parse(mSize[1]), int.Parse(mSize[2]),
                                        int.Parse(mSize[3]), int.Parse(mSize[4])));
                                    TempPath = mSize[0];
                                }
                                else
                                {
                                    Bitmap screenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                                    System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);
                                    gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                                    matTarget = BitmapConverter.ToMat(screenshot);

                                    gfxScreenshot.Dispose();
                                    screenshot.Dispose();
                                }


                                Mat matTemplate = new Mat(TempPath, ImreadModes.Color);
                                int temp_w = matTemplate.Width / 2, temp_h = matTemplate.Height / 2; // center x y

                                //System.Windows.Forms.MessageBox.Show(RunTemplateMatch(matTarget, matTemplate));
                                string return_xy = RunTemplateMatch(matTarget, matTemplate);
                                //if (!return_xy.Equals(""))
                                //{
                                //    string[] xy = return_xy.Split(',');
                                //    SetCursorPos(int.Parse(xy[0]) + temp_w, int.Parse(xy[1]) + temp_h);
                                //}

                                if (!return_xy.Equals(""))
                                {
                                    // Add Key
                                    if (Event[0].Length > 0)
                                    {
                                        mDoSortedList.Add(Event[0], return_xy);
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

                            //mInputSimulator.Keyboard.KeyPress(VirtualKeyCode.Space);
                            //KeysConverter kc = new KeysConverter();
                            //string keyChar = kc.ConvertToString("A");
                            //***************** SendKeys *****************

                            //***************** InputSimulator *****************
                            string str = CommandData;
                            char[]  arr = str.ToCharArray();
                            foreach (char c in arr)
                            {
                                //mInputSimulator.Keyboard.KeyPress((VirtualKeyCode)ConvertCharToVirtualKey(c));
                                mInputSimulator.Keyboard.KeyDown((VirtualKeyCode)ConvertCharToVirtualKey(c));
                                Thread.Sleep(100);
                                mInputSimulator.Keyboard.KeyUp((VirtualKeyCode)ConvertCharToVirtualKey(c));
                            }
                            //VirtualKeyCode myEnum = (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), "Enter");

                            //Example: Modified keystrokes such as CTRL - C
                            // CTRL-C (effectively a copy command in many situations)
                            //mInputSimulator.Keyboard.KeyDown((VirtualKeyCode)ConvertCharToVirtualKey(c), (WindowsInputLib.Native.ModifierKeys)ConvertCharToVirtualKey(c));
                            //mInputSimulator.Keyboard.KeyDown((VirtualKeyCode)ConvertCharToVirtualKey(''), WindowsInputLib.Native.ModifierKeys.Control);
                            //***************** InputSimulator *****************

                            break;

                        case "ModifierKey":

                            //string str = CommandData;
                            //char[] arr = str.ToCharArray();
                            //foreach (char c in arr)
                            //{
                            //    mInputSimulator.Keyboard.KeyPress((VirtualKeyCode)ConvertCharToVirtualKey(c),
                            //        (WindowsInputLib.Native.ModifierKeys)ConvertCharToVirtualKey(c));
                            //}

                            break;

                        case "RemoveKey":

                            if (mDoSortedList.IndexOfKey(Event[0]) != -1)
                            {
                                mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey(Event[0]));
                            }

                            break;
                        case "Get Point":

                            // Add Key
                            if (Event[0].Length > 0)
                            {
                                System.Drawing.Point point = System.Windows.Forms.Control.MousePosition;
                                mDoSortedList.Add(Event[0], point.X.ToString() + "," + point.Y.ToString());
                            }

                            break;

                        case "Run exe":

                            try
                            {
                                Process.Start(CommandData);
                            }
                            catch { }

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
                            TempBitmap = makeScreenshot_clip(int.Parse(str_clip[0]), int.Parse(str_clip[1]),int.Parse(str_clip[2]), int.Parse(str_clip[3]));

                            break;

                        case "Draw":

                            string TempPathd = CommandData;
                            Mat matTargetd = null;
                            if (TempPathd.Equals(""))
                            {
                                TempPathd = "s.png";
                            }

                            if (TempPathd.IndexOf(',') != -1)
                            {
                                //s.png,0,0,300,300
                                string[] mSize = TempPathd.Split(',');
                                matTargetd = BitmapConverter.ToMat(makeScreenshot_clip(int.Parse(mSize[1]), int.Parse(mSize[2]), int.Parse(mSize[3]), int.Parse(mSize[4])));
                                TempPathd = mSize[0];
                            }
                            else
                            {

                                Bitmap screenshot = new Bitmap(Screen.PrimaryScreen.Bounds.Width, Screen.PrimaryScreen.Bounds.Height);
                                System.Drawing.Graphics gfxScreenshot = System.Drawing.Graphics.FromImage(screenshot);
                                gfxScreenshot.CopyFromScreen(Screen.PrimaryScreen.Bounds.X, Screen.PrimaryScreen.Bounds.Y, 0, 0, Screen.PrimaryScreen.Bounds.Size, CopyPixelOperation.SourceCopy);
                                matTargetd = BitmapConverter.ToMat(screenshot);

                                gfxScreenshot.Dispose();
                                screenshot.Dispose();
                            }


                            Mat matTemplated = new Mat(TempPathd, ImreadModes.Color);
                            int temp_wd = matTemplated.Width / 2, temp_hd = matTemplated.Height / 2; // center x y

                            //System.Windows.Forms.MessageBox.Show(RunTemplateMatch(matTarget, matTemplate));
                            string return_xyd = RunTemplateMatch(matTargetd, matTemplated);
                            if (!return_xyd.Equals(""))
                            {
                                string[] xy = return_xyd.Split(',');

                                gfx.BeginScene();
                                //gfx.DrawTextWithBackground(_font, _red, _black, 10, 10, return_xyd.ToString());
                                gfx.DrawRoundedRectangle(_red, RoundedRectangle.Create(int.Parse(xy[0]), int.Parse(xy[1]), temp_wd * 2, temp_hd * 2, 6), 2);
                                gfx.EndScene();

                                // Add Key
                                if (Event[0].Length > 0)
                                {
                                    mDoSortedList.Add(Event[0], return_xyd);
                                }
                                // Move
                                //SetCursorPos(int.Parse(xy[0]) + temp_wd, int.Parse(xy[1]) + temp_hd);
                            }
                            matTargetd.Dispose();
                            matTemplated.Dispose();

                            break;

                        case "Clean Draw":

                            gfx.BeginScene(); // call before you start any drawing
                            gfx.ClearScene();
                            gfx.EndScene();

                            break;

                        case "PlaySound":

                            string SoundPath = CommandData;
                            if (Tempflag == true)
                            {
                                // SoundPlayer
                                SoundPlayer mWaveFile = new SoundPlayer(SoundPath);
                                mWaveFile.PlaySync();
                                Tempflag = false;
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

                        case "Loop":
                            do
                            {
                                n = -1;
                            } while (false);

                            break;

                        default:

                            break;
                    }
                    #endregion

                }

                n++;
            }
        }

        #region Script Panel
        public class mTable
        {
            public bool mTable_IsEnable { get; set; }
            public string mTable_Mode { get; set; }
            public string mTable_Action { get; set; }
            public string mTable_Event { get; set; }
        }
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
                eDataTable.Add(new eTable()
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

        private List<mTable> Load_Script_to_DataTable(string mfilePath)
        {
            List<mTable> tempDataTable = new List<mTable>();
            string fileContent = string.Empty;
            StreamReader reader = new StreamReader(mfilePath);

            try
            {
                // read test
                fileContent = reader.ReadToEnd();
                fileContent.Replace(";", "%;");
                string[] SplitStr = fileContent.Split(';');

                for (int i = 0; i < SplitStr.Length - 4; i += 4)
                {
                    tempDataTable.Add(new mTable()
                    {
                        mTable_IsEnable = bool.Parse(SplitStr[i].Replace("%", "")),
                        mTable_Mode = SplitStr[i + 1].Replace("%", ""),
                        mTable_Action = SplitStr[i + 2].Replace("%", ""),
                        mTable_Event = SplitStr[i + 3].Replace("%", ""),
                    });
                }
            }
            catch
            {
                this.ShowMessageAsync("Load_Script_to_DataTable", "Error!");
            }
            return tempDataTable;
        }
        private void eDataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Save_Script();
        }

        private void eDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int columnIndex = eDataGrid.Columns.IndexOf(eDataGrid.CurrentCell.Column);
            if (columnIndex < 0)
            {
                return;
            }

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
                catch { }
            }

            if (eDataGrid.Columns[columnIndex].Header.ToString().Equals("+"))
            {
                // Get index
                int tableIndex = eDataGrid.Items.IndexOf(eDataGrid.CurrentItem);

                try
                {
                    if (tableIndex < eDataTable.Count() - 1)
                    {
                        // Insert Item
                        eDataGrid.DataContext = null;
                        eDataTable.Insert(tableIndex + 1, new eTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                    else
                    {
                        eDataGrid.DataContext = null;
                        eDataTable.Add(new eTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "", eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                }
                catch { }
            }
        }

        private void eDataGrid_CurrentCellChanged(object sender, EventArgs e)
        {
            Save_Script();
        }
        #endregion

        #region Edit Panel
        public class eTable
        {
            public bool eTable_Enable { get; set; }
            public string eTable_Name { get; set; }
            public string eTable_Key { get; set; }
            public string eTable_State { get; set; }
            public string eTable_Note { get; set; }
            public string eTable_Path { get; set; }
        }

        Thread mThread = null;
        // data
        bool Tempflag = false;
        Bitmap TempBitmap;
        private void Run_script()
        {
            if (Ring.IsActive == true)
            {

                mThread.Abort();
            }
            Ring.IsActive = true;

            mThread = new Thread(() =>
            {
                Script(mDataTable);
            });
            mThread.Start();
        }
        private void Stop_script()
        {
            mThread.Abort(); //main thread aborting newly created thread.  
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
        private async void Btn_Save_Click(object sender, RoutedEventArgs e) // async
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
        private void Btn_Run_Click(object sender, RoutedEventArgs ee)
        {
            Run_script();
        }
        private void Btn_Stop_Click(object sender, RoutedEventArgs ee)
        {
            Stop_script();
        }
        #endregion

        #region DataGrid Event
        private void mDataGrid_AddingNewItem(object sender, AddingNewItemEventArgs e)
        {
            e.NewItem = new mTable
            {
                mTable_IsEnable = true,
                mTable_Mode = "",
                mTable_Action = "",
                mTable_Event = ""
            };
        }
        private void mDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Header = (e.Row.GetIndex()).ToString();
        }
        private void mDataGrid_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            if (columnIndex < 0)
            {
                return;
            }

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
                catch { }
            }

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("+"))
            {
                // Get index
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);

                try
                {
                    if (tableIndex < mDataTable.Count() - 1)
                    {
                        // Insert Item
                        mDataGrid.DataContext = null;
                        mDataTable.Insert(tableIndex + 1, new mTable() { mTable_IsEnable = true, mTable_Mode = "", mTable_Action = "", mTable_Event = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                    else
                    {
                        mDataGrid.DataContext = null;
                        mDataTable.Add(new mTable() { mTable_IsEnable = true, mTable_Mode = "", mTable_Action = "", mTable_Event = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                }
                catch { }
            }
        }
        #endregion

        #region APP
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
            if (Btn_ON.Content.Equals("ON")){
                Btn_ON.Content = "OFF";
            }
            else{
                Btn_ON.Content = "ON";
            }
        }
        private void Btn_About_Click(object sender, RoutedEventArgs e)
        {
            this.ShowMessageAsync("...........", "About");
        }
        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            // .ini window (x,y)
            var parser = new FileIniDataParser();
            IniData data = new IniData();
            data = parser.ReadFile("user.ini");
            data["Def"]["x"] = this.Left.ToString();
            data["Def"]["y"] = this.Top.ToString();
            parser.WriteFile("user.ini", data);

            for (int i = 0; i < _workerThreads.Count; i++)
            {
                if (_workerThreads[i].IsAlive)
                {
                    _workerThreads[i].Abort();
                }
            }

            if (IsActive)
            {
                mThread.Abort();
            }
        }
        #endregion

        // **********************************  End  *****************************************

        public int MakeLParam(int LoWord, int HiWord)
        {
            return ((HiWord << 16) | (LoWord & 0xffff));
        }


        public static Keys ConvertCharToVirtualKey(char ch)
        {
            short vkey = VkKeyScan(ch);
            Keys retval = (Keys)(vkey & 0xff);
            int modifiers = vkey >> 8;
            if ((modifiers & 1) != 0) retval |= Keys.Shift;
            if ((modifiers & 2) != 0) retval |= Keys.Control;
            if ((modifiers & 4) != 0) retval |= Keys.Alt;
            return retval;
        }
        private int StringToVirtualKeyCode(String str)
        {
        int value = 0;
            Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
            foreach (int enumValue in enumValueArray)
            {
                if (Enum.GetName(typeof(VirtualKeyCode), enumValue).Equals(str))
                {
                    value = enumValue;
                }
            }
            // ArrayList
            //ArrayList myArrayList = new ArrayList();
            //myArrayList.AddRange(enumValueArray);
            return value;
        }
        private static void GetEnumVirtualKeyCodeValues()
        {
            Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
            foreach (int enumValue in enumValueArray)
            {
                Console.WriteLine("Name: " + Enum.GetName(typeof(VirtualKeyCode), enumValue) + " , Value: " + enumValue);
            }
        }

        private static void ParseEnum()
        {
            VirtualKeyCode ms = (VirtualKeyCode)Enum.Parse(typeof(VirtualKeyCode), "Enter");
            Console.WriteLine(ms.ToString());
            Array enumValueArray = Enum.GetValues(typeof(VirtualKeyCode));
        }

       
    }
}