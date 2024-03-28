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
using System.Text.RegularExpressions;
using DynamicExpresso;
using NUnit.Framework;
using System.Windows.Automation.Peers;
using System.Windows.Automation.Provider;
using System.Windows.Controls.Primitives;
using System.Reflection;
using System.Data;
using System.Web.UI.WebControls;

namespace Metro
{
    public partial class MainWindow : MetroWindow
    {
        public static bool IsAdmin => new WindowsPrincipal(WindowsIdentity.GetCurrent()).IsInRole(WindowsBuiltInRole.Administrator);
        private static readonly DependencyProperty IsRecordProperty = DependencyProperty.Register("IsRecord", typeof(bool), typeof(MainWindow), new PropertyMetadata(true));
        private bool IsRecord
        {
            get { return (bool)this.GetValue(IsRecordProperty); }
            set { this.SetValue(IsRecordProperty, value); }
        }

        // NotifyIcon
        private static NotifyIcon mNotifyIcon = new NotifyIcon();
        private static Icon OnIcon, OffIcon;
        // DataGrid
        private List<MainTable> mDataTable = new List<MainTable>();
        private List<EditTable> eDataTable = new List<EditTable>();

        private List<Thread> _workerThreads = new List<Thread>();
        private SoundPlayer mAlertSound = new SoundPlayer(Metro.Properties.Resources.sound);
        private SettingHelper mSettingHelper = new SettingHelper();
        private Edit mEdit = new Edit();
        private bool TestMode = false;
        private int TestMode_Delay = 0;

        #region Globalmousekeyhook
        private static IKeyboardMouseEvents m_GlobalHook, Main_GlobalHook, Main_GlobalKeyUpHook;

        private int now_x, now_y;

        private double getTimeToMilliseconds()
        {
            TimeSpan KeyTimeSpan = new TimeSpan(DateTime.Now.Ticks - StartTime.Ticks);
            return Math.Floor(KeyTimeSpan.TotalMilliseconds);
        }

        private void Subscribe()
        {
            // Note: for the application hook, use the Hook.AppEvents() instead
            m_GlobalHook = Hook.GlobalEvents();
            m_GlobalHook.MouseDownExt += GlobalHookMouseDownExt;
            m_GlobalHook.MouseUpExt += GlobalHookMouseUpExt;
            //m_GlobalHook.KeyPress += GlobalHookKeyPress;
            m_GlobalHook.KeyDown += GlobalHookKeyDown;
        }

       
        private void Unsubscribe()
        {
            m_GlobalHook.MouseDownExt -= GlobalHookMouseDownExt;
            m_GlobalHook.MouseUpExt -= GlobalHookMouseUpExt;
            //m_GlobalHook.KeyPress -= GlobalHookKeyPress;
            m_GlobalHook.KeyDown -= GlobalHookKeyDown;

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
            MouseRecord(e, "Down");
        }

        private void GlobalHookMouseUpExt(object sender, MouseEventExtArgs e)
        {
            MouseRecord(e,"Up");
        }

        private void MouseRecord(MouseEventExtArgs e,string type)
        {
            if (Btn_Toggle.IsOn == true && Btn_Toggle.IsMouseOver == false)
            {
                double KeyTimeValue = getTimeToMilliseconds();
                double DelayTime = getTimeToMilliseconds() - LKeyListTimeValue;
                LKeyListTimeValue = KeyTimeValue;

                if (MKeyList.IndexOfKey(e.Button.ToString() + "_" + "Down") != -1)
                    MKeyList.RemoveAt(MKeyList.IndexOfKey(e.Button.ToString() + "_" + "Down"));
                else
                    MKeyList.Add(e.Button.ToString() + "_" + "Down", "");

                if (mDataTable[mDataTable.Count() - 1].Mode.Equals("Delay")) {
                    int oValue = 0;
                    int.TryParse(mDataTable[mDataTable.Count() - 1].Action, out oValue);
                    mDataTable[mDataTable.Count() - 1].Action = (oValue + (int)DelayTime / 2).ToString();
                }
                else
                    mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = ((int)DelayTime / 2).ToString(), Event = "", Note = "" });

                mDataTable.Add(new MainTable() { Enable = true, Mode = "Move", Action = now_x.ToString() + "," + now_y.ToString(), Event = "", Note = "" });
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = ((int)DelayTime/2).ToString(), Event = "", Note = "" });
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Click", Action = e.Button.ToString() + "_" + type, Event = "", Note = "" });
              
            }
            // uncommenting the following line will suppress the middle mouse button click
            // if (e.Buttons == MouseButtons.Middle) { e.Handled = true; }
        }

        private void HookManager_MouseMove(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            now_x = e.X;
            now_y = e.Y;

            if (Btn_Toggle.IsOn == true && Btn_Toggle.IsMouseOver == false)
            {
                if ((MKeyList.IndexOfKey("Left_Down") != -1) || (MKeyList.IndexOfKey("Right_Down") != -1))
                {
                    double KeyTimeValue = getTimeToMilliseconds();
                    double DelayTime = getTimeToMilliseconds() - LKeyListTimeValue;
                    if (DelayTime > 100) {
                        LKeyListTimeValue = KeyTimeValue;

                        if (mDataTable[mDataTable.Count() - 1].Mode.Equals("Delay"))
                            mDataTable[mDataTable.Count() - 1].Action = (int.Parse(mDataTable[mDataTable.Count() - 1].Action) + (int)DelayTime / 2).ToString();
                        else
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = ((int)DelayTime / 2).ToString(), Event = "", Note = "" });

                        mDataTable.Add(new MainTable() { Enable = true, Mode = "Move", Action = now_x.ToString() + "," + now_y.ToString(), Event = "", Note = "" });
                        mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = ((int)DelayTime / 2).ToString(), Event = "", Note = "" });
                    }
                }
            }

            PopupText.Text = "X: " + e.X + " Y: " + e.Y;

            if (IsInfoToggleButtonChecked && Is_LostKeyboardFocus) {
                ToggleButtonAutomationPeer peer = new ToggleButtonAutomationPeer(InfoToggleButton);
                IToggleProvider toggleProvider = peer.GetPattern(PatternInterface.Toggle) as IToggleProvider;
                toggleProvider.Toggle();
                Is_LostKeyboardFocus = false;
            }
        }

        private void GlobalHookKeyDown(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (Btn_Toggle.IsOn == true)
            {
                KeyRecord(e, "KeyDown");
            }
        }

        private void Main_GlobalHookKeyUp(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            // for Hotkey setting
            if (TextBox_OnOff_Hotkey.IsFocused)
            {
                TextBox_OnOff_Hotkey.Text = ConvertHelper.ConvertKeyCode(e.KeyCode.ToString());
                OnOff_CrtlKey_Chk.Focus();
            }
            else if (TextBox_Run_Hotkey.IsFocused)
            {
                TextBox_Run_Hotkey.Text = ConvertHelper.ConvertKeyCode(e.KeyCode.ToString());
                Run_CrtlKey_Chk.Focus();
            }
            else if (TextBox_Stop_Hotkey.IsFocused)
            {
                TextBox_Stop_Hotkey.Text = ConvertHelper.ConvertKeyCode(e.KeyCode.ToString());
                Stop_CrtlKey_Chk.Focus();
            }

            if (Btn_Toggle.IsOn == true)
            {
                KeyRecord(e, "KeyUp");
            }
        }

        SortedList MKeyList = new SortedList();
        SortedList LKeyList = new SortedList();
        private double LKeyListTimeValue = 0;
        private void KeyRecord(System.Windows.Forms.KeyEventArgs e, string type)
        {
            if (!e.KeyCode.ToString().Equals(""))
            {
                string mKeyCode = e.KeyCode.ToString();
                mKeyCode = ConvertHelper.ConvertKeyCode(mKeyCode);

                if (LKeyList.IndexOfKey(e.KeyValue) != -1)
                {
                    if (!((string)LKeyList.GetByIndex(LKeyList.IndexOfKey(e.KeyValue))).Equals(type))
                    {
                        LKeyList.RemoveAt(LKeyList.IndexOfKey(e.KeyValue));
                    }
                    else
                        return;
                }
                else
                    LKeyList.Add(e.KeyValue, type);

                double KeyTimeValue = getTimeToMilliseconds();
                double DelayTime = getTimeToMilliseconds() - LKeyListTimeValue;
                LKeyListTimeValue = KeyTimeValue;

                if (mKeyCode.IndexOf("Oem") == -1)
                {
                   
                    if (mDataTable[mDataTable.Count() - 1].Mode.Equals("Delay"))
                        mDataTable[mDataTable.Count() - 1].Action = (int.Parse(mDataTable[mDataTable.Count() - 1].Action) + (int)DelayTime).ToString();
                    else
                        mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = ((int)DelayTime).ToString(), Event = "", Note = "" });

                    if (mSettingHelper.TypeOfKeyboardInput.Equals("Normal") || mKeyCode.Equals("WIN") || mKeyCode.Equals("Apps") || mKeyCode.Equals("SNAPSHOT") || mKeyCode.Equals("Scroll") || mKeyCode.Equals("Pause"))
                    {
                        if (type.Equals("KeyDown"))
                        {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Key", Action = mKeyCode +",Down", Event = "", Note = "" });
                        }
                        else 
                        {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "Key", Action = mKeyCode + ",Up", Event = "", Note = "" });
                        }
                    }
                    else
                    {
                        if (type.Equals("KeyDown"))
                        {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "SendKeyDown", Action = mKeyCode.ToUpper(), Event = "", Note = "" });
                        }
                        else
                        {
                            mDataTable.Add(new MainTable() { Enable = true, Mode = "SendKeyUp", Action = mKeyCode.ToUpper(), Event = "", Note = "" });
                        }
                    }
                }
            }
        }
            #endregion
            Overlay overlay = new Overlay();
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

            // Combobox List
            List<string> mList = new List<string>() {
                "Move","Offset","Click", 
                "Key","ModifierKey","SendKeyDown","SendKeyUp","WriteClipboard","Delay",
                "Calc","Calc-Check",
                "Match","Match RGB","Match&Draw","RandomTrigger",
                "RemoveEvent","Jump","Goto","Loop",
                "Run .exe","PlaySound","Clear Screen"
                //"PostMessage","FindWindow",
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
                mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "200", Event = "", Note = "" });
                mDataGrid.DataContext = mDataTable;
                mEdit.StartEdit("");
            }

            #endregion

            // Load Script setting
            LoadScriptSetting();

            // Load setting
            OnOff_CrtlKey_Chk.IsChecked = mSettingHelper.OnOff_CrtlKey;
            Run_CrtlKey_Chk.IsChecked = mSettingHelper.Run_CrtlKey;
            Stop_CrtlKey_Chk.IsChecked = mSettingHelper.Stop_CrtlKey;

            OnOff_AltKey_Chk.IsChecked = mSettingHelper.OnOff_AltKey;
            Run_AltKey_Chk.IsChecked = mSettingHelper.Run_AltKey;
            Stop_AltKey_Chk.IsChecked = mSettingHelper.Stop_AltKey;

            TextBox_OnOff_Hotkey.Text = mSettingHelper.OnOff_Hotkey;
            TextBox_Run_Hotkey.Text = mSettingHelper.Run_Hotkey;
            TextBox_Stop_Hotkey.Text = mSettingHelper.Stop_Hotkey;

            TypeOfKeyboardInput0.IsChecked = mSettingHelper.TypeOfKeyboardInput.Equals("Game") ? true : false;
            TypeOfKeyboardInput1.IsChecked = !mSettingHelper.TypeOfKeyboardInput.Equals("Game") ? true : false;

            // Load HideOnSatrt setting
            HideOnSatrt_Toggle.IsOn = mSettingHelper.HideOnSatrt;

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

            TestMode_Delay = mSettingHelper.TestMode_Delay;
            TestMode_Slider.Value = TestMode_Delay;

            // Load Topmost setting
            Top_Toggle.IsOn = mSettingHelper.Topmost;

            KListener();

            Main_GlobalKeyUpHook = Hook.GlobalEvents();
            Main_GlobalKeyUpHook.KeyUp += Main_GlobalHookKeyUp;

            // NotifyIcon
            System.Drawing.Bitmap bitmap = System.Drawing.Icon.ExtractAssociatedIcon(System.Reflection.Assembly.GetExecutingAssembly().Location).ToBitmap();
            for (var y = 0; y < bitmap.Height; y++)
            {
                for (var x = 0; x < bitmap.Width; x++)
                {
                    if (bitmap.GetPixel(x, y).R > 240)
                    {
                        bitmap.SetPixel(x, y, System.Drawing.Color.FromArgb(255, 141, 60)); 
                    }
                }
            }
            IntPtr Hicon =bitmap.GetHicon();
            OffIcon = System.Drawing.Icon.FromHandle(Hicon);
            OnIcon = System.Drawing.Icon.ExtractAssociatedIcon(System.Reflection.Assembly.GetExecutingAssembly().Location);

            mNotifyIcon.Icon = OnIcon;
            mNotifyIcon.ContextMenuStrip = new ContextMenuStrip();
            ToolStripItem mToolStripItem = mNotifyIcon.ContextMenuStrip.Items.Add("Close Menu", null, this.notifyIcon_Close_Click);
            mToolStripItem.Image = Properties.Resources.x;
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("Visit_Website").ToString(), null, this.notifyIcon_Visit_Click);
            mNotifyIcon.ContextMenuStrip.Items.Add("-");
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("HideShow").ToString(), null, this.notifyIcon_DoubleClick);
            mNotifyIcon.ContextMenuStrip.Items.Add(FindResource("Exit").ToString(), null, this.notifyIcon_Exit_Click);
            mNotifyIcon.MouseUp += MouseUp;
            mNotifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);
            mNotifyIcon.Visible = true;

            if (!IsAdmin)
            {
                TextBlock_Please.Visibility = Visibility.Visible;
            }

            this.WindowState = WindowState.Minimized;
            this.Show();
            HwndSource source = PresentationSource.FromVisual(this) as HwndSource;
            source.AddHook(WndProc);

            if (!mSettingHelper.HideOnSatrt)
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

            UnitTest("Delay","100");

            var assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream("Metro.Resources.Documentation.md"))
            using (StreamReader reader = new StreamReader(stream))
            {
                string result = reader.ReadToEnd();
                Viewer.Markdown = result;
            }
        }

        #region KListener
        private bool IsHookManager_MouseMove = false;
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

        SortedList KeyList = new SortedList();
        DateTime StartTime = new DateTime(2001, 1, 1);

        private string PassCtrlKey = "";

        private void Main_GlobalHookKeyPress(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            V V = new V();
            string KeyString = ConvertHelper.ConvertKeyCode(e.KeyCode.ToString());
            Console.WriteLine("KeyCode: {0} KeyString: {1}",e.KeyCode.ToString(),KeyString);

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

            // ON / OFF
            if (!Script_Toggle.IsOn && KeyString.Equals(mSettingHelper.OnOff_Hotkey)
                && !(mSettingHelper.OnOff_CrtlKey && e.Control == false) && !(mSettingHelper.OnOff_AltKey && e.Alt == false)){
                if (Btn_ON.Content.Equals("ON"))
                {
                    Btn_ON.Content = "OFF";
                    Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
                    mNotifyIcon.Icon = OffIcon;

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
                    mNotifyIcon.Icon = OnIcon;
                }
                eDataGrid.DataContext = null;
                eDataGrid.DataContext = eDataTable;
            }
            if (!Btn_ON.Content.Equals("ON"))
            {
                return;
            }

            if (KeyString.Equals(mSettingHelper.Run_Hotkey) 
                && !(mSettingHelper.Run_CrtlKey && e.Control == false) && !(mSettingHelper.Run_AltKey && e.Alt == false)){
                ClearScreen_Btn.Focus();
                AlertSound();
                ShowBalloon("Run", "...");
                Run_script();
            }
            if (KeyString.Equals(mSettingHelper.Stop_Hotkey) 
                && !(mSettingHelper.Stop_CrtlKey && e.Control == false) && !(mSettingHelper.Stop_AltKey && e.Alt == false)){
                ClearScreen_Btn.Focus();
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
            KeyString = PassCtrlKey + KeyString;

            // Select Script
            bool IsRun = false;
            for (int i = 0; i < eDataTable.Count; i++)
            {
                if (KeyString.Equals(eDataTable[i].eTable_Key) && eDataTable[i].eTable_Enable == true)
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
            V V = new V();

            SortedList mDoSortedList = new SortedList();
            SortedList KeyActionList = new SortedList();
            SortedList ShortKeyActionList = new SortedList();
            SortedList MouseActionList = new SortedList();

            Interpreter interpreter = new Interpreter();
            // StartPostion
            System.Drawing.Point mPoint = new System.Drawing.Point();
            GetCursorPos(ref mPoint);
            interpreter = interpreter.SetVariable("StartPostion_X", (int)mPoint.X);
            interpreter = interpreter.SetVariable("StartPostion_Y", (int)mPoint.Y);
            interpreter = interpreter.SetVariable("StartPosition_X", (int)mPoint.X);
            interpreter = interpreter.SetVariable("StartPosition_Y", (int)mPoint.Y);
            // random func
            Random RM = new Random();
            Func<int, int, int> random = (s, e) => RM.Next(s, e + 1);
            interpreter = interpreter.SetFunction("random", random);

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
                    Thread.Sleep(TestMode_Delay);
                }
                else if (Mode.Equals("Test"))
                {
                    Thread.Sleep(TestMode_Delay);
                }

                string Command = minDataTable[n].Mode;
                string CommandData = minDataTable[n].Action;
                bool CommandEnable = minDataTable[n].Enable;

                string[] Event = minDataTable[n].Event.Split(',');
                if (minDataTable[n].Event == "") { Event = new string[0]; }

                if (CommandEnable)
                {
                    // CurrentPosition
                    GetCursorPos(ref mPoint);
                    interpreter = interpreter.SetVariable("CurrentPosition_X", (int)mPoint.X);
                    interpreter = interpreter.SetVariable("CurrentPosition_Y", (int)mPoint.Y);

                    Mat matTemplate = null, matTarget = null;
                    Boolean err = false;
                    try
                    {
                        string Temp_CommandData = CommandData;
                        foreach (Match match in Regex.Matches(CommandData, @"\{.*?\}"))
                        {
                            string val = match.Value.Substring(1, match.Value.Length - 2);
                            if (interpreter.DetectIdentifiers(val).UnknownIdentifiers.ToArray().Length == 0)
                            {
                                Temp_CommandData = Temp_CommandData.Replace(match.Value, 
                                    interpreter.Eval(val).ToString());
                            }
                        }
                        CommandData = Temp_CommandData;

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
                                        if (MouseActionList.IndexOfKey("LEFT_DOWN") == -1) MouseActionList.Add("LEFT_DOWN", "");
                                        mInputSimulator.Mouse.LeftButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.LeftButtonUp();
                                        if (MouseActionList.IndexOfKey("LEFT_DOWN") != -1) MouseActionList.RemoveAt(MouseActionList.IndexOfKey("LEFT_DOWN"));
                                    }
                                    if (CommandData.Equals("LEFT_DOWN"))
                                    {
                                        if (MouseActionList.IndexOfKey("LEFT_DOWN") == -1) MouseActionList.Add("LEFT_DOWN", "");
                                        mInputSimulator.Mouse.LeftButtonDown();
                                    }
                                    if (CommandData.Equals("LEFT_UP"))
                                    {
                                        mInputSimulator.Mouse.LeftButtonUp();
                                        if (MouseActionList.IndexOfKey("LEFT_DOWN") != -1) MouseActionList.RemoveAt(MouseActionList.IndexOfKey("LEFT_DOWN"));
                                    }
                                    if (CommandData.Equals("RIGHT"))
                                    {
                                        if (MouseActionList.IndexOfKey("RIGHT_DOWN") == -1) MouseActionList.Add("RIGHT_DOWN", "");
                                        mInputSimulator.Mouse.RightButtonDown();
                                        Thread.Sleep(200);
                                        mInputSimulator.Mouse.RightButtonUp();
                                        if (MouseActionList.IndexOfKey("RIGHT_DOWN") != -1) MouseActionList.RemoveAt(MouseActionList.IndexOfKey("RIGHT_DOWN"));
                                    }
                                    if (CommandData.Equals("RIGHT_DOWN"))
                                    {
                                        if (MouseActionList.IndexOfKey("RIGHT_DOWN") == -1) MouseActionList.Add("RIGHT_DOWN", "");
                                        mInputSimulator.Mouse.RightButtonDown();
                                    }
                                    if (CommandData.Equals("RIGHT_UP"))
                                    {
                                        mInputSimulator.Mouse.RightButtonUp();
                                        if (MouseActionList.IndexOfKey("RIGHT_DOWN") != -1) MouseActionList.RemoveAt(MouseActionList.IndexOfKey("RIGHT_DOWN"));
                                    }
                                }

                                break;

                            case "Match&Draw":
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
                                            overlay.DrawRectangle(int.Parse(xy[0]), int.Parse(xy[1]), temp_w * 2, temp_h * 2, MatchArr[0]);
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
                                    Regex regex = new Regex(Regex.Escape(","));
                                    CommandData = regex.Replace(CommandData, "OEM_COMMA", 1);
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
                                            if (KeyActionList.IndexOfKey(mKeyCode) == -1) KeyActionList.Add(mKeyCode, "Key");
                                           
                                        }
                                        else if (mKey[1].Equals("UP"))
                                        {
                                            mInputSimulator.Keyboard.KeyUp(mKeyCode);

                                            if (KeyActionList.IndexOfKey(mKeyCode) != -1)KeyActionList.RemoveAt(KeyActionList.IndexOfKey(mKeyCode));
                                        }
                                        else {
                                            mInputSimulator.Keyboard.KeyDown(mKeyCode);
                                            KeyActionList.Add(mKeyCode, "Key");
                                            Thread.Sleep(int.Parse(mKey[1])); ;
                                            mInputSimulator.Keyboard.KeyUp(mKeyCode);

                                            if (KeyActionList.IndexOfKey(mKeyCode) != -1)KeyActionList.RemoveAt(KeyActionList.IndexOfKey(mKeyCode));
                                        }
                                    }
                                    else {
                                        mInputSimulator.Keyboard.KeyDown(mKeyCode);
                                        KeyActionList.Add(mKeyCode, "Key");
                                        Thread.Sleep(250);
                                        mInputSimulator.Keyboard.KeyUp(mKeyCode);

                                        if (KeyActionList.IndexOfKey(mKeyCode) != -1) KeyActionList.RemoveAt(KeyActionList.IndexOfKey(mKeyCode));
                                    }
                                }
                                else {
                                    string str = CommandData;
                                    char[] arr = str.ToCharArray();

                                    foreach (char c in arr)
                                    {
                                        mInputSimulator.Keyboard.KeyPress((VirtualKeyCode)ConvertHelper.ConvertCharToVirtualKey(c));
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
                                                if (ShortKeyActionList.IndexOfKey(value) == -1) ShortKeyActionList.Add(value, "SendKeyDown");

                                            }
                                            else {
                                                ky.SendKeyUp(value);
                                                if (ShortKeyActionList.IndexOfKey(value) != -1) ShortKeyActionList.RemoveAt(ShortKeyActionList.IndexOfKey(value));
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

                                String[] KeyStr = CommandData.ToUpper().Split('|');
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

                                if (Event.Length == 0)
                                {
                                    IList mlist = mDoSortedList.GetKeyList();
                                    Event = new string[mlist.Count];
                                    for (int i = 0; i < mlist.Count; i++)
                                    {
                                        Event[i] = mlist[i].ToString();
                                    }
                                }

                                for (int i = 0; i < Event.Length; i++)
                                {
                                    if (mDoSortedList.IndexOfKey(Event[i]) != -1)
                                    {
                                        string[] TempEvent_Data = mDoSortedList.GetByIndex(mDoSortedList.IndexOfKey(Event[i])).ToString().Split(',');
                                        mDoSortedList.RemoveAt(mDoSortedList.IndexOfKey(Event[i]));

                                        if (CommandData.ToUpper().Equals("PUSH"))
                                        {
                                            if (TempEvent_Data.Length >= 4)
                                            {
                                                mDoSortedList.Add(Event[i], string.Join(",", TempEvent_Data.Skip(2).ToArray()));
                                            }
                                        }
                                    }
                                }

                                break;

                            case "Run .exe":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    try
                                    {
                                        if (CommandData.ToUpper().EndsWith(".EXE"))
                                        {
                                            Process.Start(CommandData);
                                        }
                                        else {
                                            int SplitIndex = CommandData.ToUpper().IndexOf(".EXE") + 4;
                                            string mPrg  = CommandData.Substring(0, SplitIndex);
                                            string mArg = CommandData.Substring(SplitIndex+1, CommandData.Length - SplitIndex-1);
                                            Process.Start(mPrg, mArg);
                                        }
                                       
                                    }
                                    catch {

                                    }
                                }
                               
                                break;

                            case "FindWindow":

                                IntPtr hwnd = FindWindow(null, CommandData);
                                if (hwnd == IntPtr.Zero)
                                    hwnd = FindWindow(CommandData, null);

                                if (hwnd != IntPtr.Zero)
                                    SetForegroundWindow(hwnd);

                                break;

                            case "Clear Screen":

                                overlay.Clear();

                                break;

                            case "PlaySound":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    SoundPlayer mWaveFile = null;
                                    try
                                    {
                                        if (CommandData.Equals(""))
                                        {
                                            SystemSounds.Beep.Play();
                                        }
                                        else {
                                            mWaveFile = new SoundPlayer(CommandData);
                                            mWaveFile.PlaySync();
                                        }
                                    }
                                    catch (Exception e)
                                    {
                                        Console.WriteLine("{0} Exception caught.", e);
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

                            case "PostMessage":

                                #region PostMessage
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
                                        overlay.Clear();
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
                                #endregion

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

                            case "RandomTrigger":

                                // Add Key
                                if (Event.Length != 0)
                                {
                                    if (V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                    {
                                        Random mRandom = new Random();
                                        if (mRandom.Next(1, 100) <= int.Parse(CommandData)) {
                                            mDoSortedList.Add(Event[0], "");
                                        }
                                    }
                                }

                                break;

                            case "Calc":

                                if (Event.Length == 0 || V.Get_EventValue(mDoSortedList, Event[0]) != null)
                                {
                                    if (Regex.IsMatch(CommandData,@"(?<!=)=(?!=)"))
                                    {
                                        string Variable = CommandData.Substring(0, CommandData.IndexOf("=")).Trim();
                                        CommandData = CommandData.Replace(CommandData.Substring(0, CommandData.IndexOf("=") + 1), "");

                                        if (interpreter.DetectIdentifiers(Variable).UnknownIdentifiers.ToArray().Length > 0) 
                                            interpreter = interpreter.SetVariable(Variable, 0);

                                        interpreter = interpreter.SetVariable(Variable, interpreter.Eval(CommandData));
                                    }
                                }

                                break;


                            case "Calc-Check":

                                // Add Key
                                if (Event.Length != 0)
                                {
                                    if (V.Get_EventValue(mDoSortedList, Event[0]) == null)
                                    {
                                        if ((bool)interpreter.Eval(CommandData))
                                        {
                                            mDoSortedList.Add(Event[0], "");
                                        }
                                    }
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

                        if (e.Source.Equals("System.Drawing"))
                        {
                            Thread.Sleep(1000);
                            n--;
                        }
                        else
                        {
                            foreach (DictionaryEntry list in KeyActionList)
                            {
                                if (list.Value.Equals("Key")) mInputSimulator.Keyboard.KeyUp((VirtualKeyCode)list.Key);
                            }
                            foreach (DictionaryEntry list in ShortKeyActionList)
                            {
                                if (list.Value.Equals("SendKeyDown")) ky.SendKeyUp((short)list.Key);
                            }
                            foreach (DictionaryEntry list in MouseActionList)
                            {
                                if (list.Key.Equals("LEFT_DOWN")) mInputSimulator.Mouse.LeftButtonUp();
                                if (list.Key.Equals("RIGHT_DOWN")) mInputSimulator.Mouse.RightButtonUp();
                            }
                            

                            mInputSimulator = null;
                            ky = null;
                            V = null;
                            ConvertHelper = null;
                            mDoSortedList.Clear();

                            err = true;

                            if (e.Message.ToString().IndexOf("Thread") == -1)
                            {
                                if (Mode.Equals("Debug"))
                                {
                                    // debug stop msg
                                    SystemSounds.Hand.Play();
                                    CreateMessage("9487");

                                    System.Windows.Forms.MessageBox.Show("[Error] Line " + (n + 1).ToString() + "\nMessage: " + e.Message, "Scriptboxie", 
                                        System.Windows.Forms.MessageBoxButtons.OK,
                                        MessageBoxIcon.Error, MessageBoxDefaultButton.Button1, 
                                        System.Windows.Forms.MessageBoxOptions.DefaultDesktopOnly);

                                    break;
                                }
                                else if (Mode.Equals("Test"))
                                {
                                    // test stop msg
                                    SystemSounds.Hand.Play();

                                    System.Windows.Forms.MessageBox.Show("[Error] Message: " + e.Message, "Scriptboxie",
                                        System.Windows.Forms.MessageBoxButtons.OK,
                                        MessageBoxIcon.Error, MessageBoxDefaultButton.Button1,
                                        System.Windows.Forms.MessageBoxOptions.DefaultDesktopOnly);

                                    break;
                                }
                            }
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
            // script end msg
            CreateMessage("1000");

        }
        #region APP
        [DllImport("user32.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern uint RegisterWindowMessage(string lpString);
        private uint MSG_SHOW;

        private void MetroWindow_Deactivated(object sender, EventArgs e)
        {
            System.Windows.Window window = (System.Windows.Window)sender;
            window.Topmost = mSettingHelper.Topmost;
        }

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

                // TestMode - mark step
                if (wParam.ToString().Substring(0, 1).Equals("8"))
                {
                    DataGridRow row = (DataGridRow)mDataGrid.ItemContainerGenerator.ContainerFromIndex(LastNumber);
                    if (row != null)
                    {
                        mDataTable[LastNumber].flag = false;
                    }

                    int number = int.Parse(wParam.ToString().Substring(1, 4));
                    row = (DataGridRow)mDataGrid.ItemContainerGenerator.ContainerFromIndex(number);
                    if (row != null)
                    {
                        LastNumber = number;
                        mDataTable[number].flag = true;
                    }

                    mDataGrid.Items.Refresh();
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
            try
            {
                System.Windows.Forms.MouseEventArgs mMouseEventArgs = (System.Windows.Forms.MouseEventArgs)e;
                if (mMouseEventArgs.Button == MouseButtons.Right)
                {
                    return;
                }
            }
            catch (Exception err)
            {
                Console.WriteLine(err);
            }

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

        private void MouseUp(object sender, System.Windows.Forms.MouseEventArgs e)
        {
            if (e.Button == MouseButtons.Right)
            {
                int BarHeight = Screen.PrimaryScreen.Bounds.Height - Screen.PrimaryScreen.WorkingArea.Height;
                if (BarHeight < 0) BarHeight = 0;
                NotifyIcon mNotifyIcon = (NotifyIcon)sender;
                int Bottom = mNotifyIcon.ContextMenuStrip.Top;
                mNotifyIcon.ContextMenuStrip.Top = (int)(Bottom - BarHeight / 2);
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
        private void notifyIcon_Close_Click(object sender, EventArgs e)
        {
            mNotifyIcon.ContextMenuStrip.Close();
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

        private void MetroWindow_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (e.NewSize.Height > 300) {
                ScriptGridScroll.MinHeight = e.NewSize.Height - 100;
                mDataGrid.MinHeight = e.NewSize.Height - 100;
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
                mNotifyIcon.Icon = OffIcon;
            }
            else
            {
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                mNotifyIcon.Icon = OnIcon;
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
                    if (int.Parse(tag_name.ToString().Replace("v", "").Replace(".", ""))
                        < int.Parse(System.Diagnostics.FileVersionInfo.GetVersionInfo(assembly.Location).FileVersion.Replace("v", "").Replace(".", "")))
                    {
                        NewVersion.Text = "Beta Version";
                    }
                    else {
                        NewVersion.Text = NewVersion.Text + " " + tag_name;
                    }
                    NewVersion.Visibility = Visibility.Visible;
                }
               
            }
        }

        private void MetroWindow_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            UnKListener();
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
        private void Save_Script()
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
            mEdit.StartEdit(filePath);
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

            mEdit.ModifiedTime = mEdit.GetModifiedTime(filePath);
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
                    eDataTable.Add(new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "※ " + string.Format(FindResource("Double_click_to_select_file").ToString()), eTable_State = "" });
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
                mNotifyIcon.Icon = OffIcon;
            }
            else {
                Save_Script();
                LoadScriptSetting();

                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                mNotifyIcon.Icon = OnIcon;
            }
        }

        private void ScriptGridScroll_PreviewMouseWheel(object sender, MouseWheelEventArgs e)
        {
            ScrollViewer mScrollViewer = sender as ScrollViewer;
            if (mScrollViewer == null) return;
            mScrollViewer.ScrollToVerticalOffset(mScrollViewer.VerticalOffset - e.Delta);
            e.Handled = true;
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
                        eDataTable.Insert(tableIndex + 1, new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "※ " + string.Format(FindResource("Double_click_to_select_file").ToString()), eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                    else
                    {
                        eDataGrid.DataContext = null;
                        eDataTable.Add(new EditTable() { eTable_Enable = true, eTable_Key = "", eTable_Name = "", eTable_Note = "", eTable_Path = "※ " + string.Format(FindResource("Double_click_to_select_file").ToString()), eTable_State = "" });
                        eDataGrid.DataContext = eDataTable;
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
            else if (eDataGrid.Columns[columnIndex].Header.ToString().Equals("Path"))
            {
                int tableIndex = eDataGrid.Items.IndexOf(mDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < eDataTable.Count())
                    {
                        string filePath = string.Empty;

                        System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                        openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
                        openFileDialog.Filter = "script files (*.txt)|*.txt"; // "txt files (*.txt)|*.txt|All files (*.*)|*.*"
                        openFileDialog.FilterIndex = 2;
                        openFileDialog.RestoreDirectory = true;
                        openFileDialog.ShowDialog();

                        //Get the path of specified file
                        filePath = openFileDialog.FileName;
                        if (filePath.Equals("")) { return; }

                        DataGridCellInfo cellInfo = eDataGrid.CurrentCell;
                        FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                        System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;
                        System.Windows.Controls.TextBox mTextBlock = (System.Windows.Controls.TextBox)cell.Content;
                        
                        mTextBlock.Text = filePath.Replace(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location)+"\\", "");

                        cell.IsEditing = false;
                        eDataGrid.CommitEdit();
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

        private Thread mThread = null, uThread = null;
        private bool Is_LostKeyboardFocus = false, IsInfoToggleButtonChecked = false;
        private void MetroWindow_LostKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            Console.WriteLine("MetroWindow_LostKeyboardFocus");
            Console.WriteLine(this.WindowState);

            if (IsInfoToggleButtonChecked && !Is_LostKeyboardFocus)
            {
                ToggleButtonAutomationPeer peer = new ToggleButtonAutomationPeer(InfoToggleButton);
                IToggleProvider toggleProvider = peer.GetPattern(PatternInterface.Toggle) as IToggleProvider;
                Is_LostKeyboardFocus = true;

                toggleProvider.Toggle();
            }
        }

        private bool isShowDialog =false;
        private async void _Window_GotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            if (mEdit.CheckIsModifie() && !isShowDialog) {
                isShowDialog = true;

                var mMetroDialogSettings = new MetroDialogSettings()
                {
                    AffirmativeButtonText = "Yes",
                    NegativeButtonText = "No",
                    AnimateShow = true,
                    AnimateHide = false
                };
                MessageDialogResult result = await this.ShowMessageAsync("",
                    "The file has been updated.Do you want to load the file again?",
                    MessageDialogStyle.AffirmativeAndNegative, mMetroDialogSettings);

                if (result == MessageDialogResult.Affirmative)
                {
                    Load_Script(mEdit.FilePath);
                }
                else {

                    mEdit.ModifiedTime = mEdit.GetModifiedTime(mEdit.FilePath);
                }
                isShowDialog = false;
            }
          
        }

        private void InfoToggleButton_Checked(object sender, RoutedEventArgs e)
        {
            if (!IsHookManager_MouseMove)
            {
                IsHookManager_MouseMove = true;
                Main_GlobalHook.MouseMove += HookManager_MouseMove;
            }

            ToggleButton mToggleButton = (ToggleButton)sender;
            IsInfoToggleButtonChecked = mToggleButton.IsChecked.Value;
        }

        private void TabControl_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            int TabIndex = ((System.Windows.Controls.TabControl)sender).SelectedIndex;
            switch (TabIndex)
            {
                case 0:
                case 2:
                    IsInfoToggleButtonChecked = false;
                    break;
            }
        }

        private void Btn_Toggle_Click(object sender, RoutedEventArgs e)
        {
            DateTime KeyTime = DateTime.Now;
            TimeSpan KeyTimeSpan = new TimeSpan(KeyTime.Ticks - StartTime.Ticks);
            double KeyTimeValue = Math.Floor(KeyTimeSpan.TotalMilliseconds);
            LKeyListTimeValue = KeyTimeValue;

            ClearScreen_Btn.Focus();

            if (!IsHookManager_MouseMove)
            {
                IsHookManager_MouseMove = true;
                Main_GlobalHook.MouseMove += HookManager_MouseMove;
            }

            if (Btn_Toggle.IsOn == true)
            {
                Stop_script();

                this.IsRecord = false;

                Btn_ON.Content = "OFF";
                Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
                mNotifyIcon.Icon = OffIcon;
                Subscribe();
            }
            else
            {
                this.IsRecord = true;

                mDataGrid.DataContext = null;
                mDataGrid.DataContext = mDataTable;
                mDataGrid_ScrollToBottom();

                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                mNotifyIcon.Icon = OnIcon;
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
                try
                {
                    Script(mDataTable, "Debug");
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

        private void UnitTest(string Mode,string Action)
        {
            if (uThread != null)
            {
                uThread.Abort();
            }

            List<MainTable> uMainTable = new List<MainTable>
            {
                new MainTable() { Enable = true, Mode = Mode, Action = Action, Event = "", Note = "" }
            };

            uThread = new Thread(() =>
                {
                    try
                    {
                        Script(uMainTable, "Test");
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
            uThread.Start();
        }

        private void ClearScreen_Click(object sender, RoutedEventArgs e)
        {
            overlay.Clear();
        }
        private void ClearScreen_Btn_PreviewKeyDown(object sender, System.Windows.Input.KeyEventArgs e)
        {
             e.Handled = true;
        }

        private void ClearScreen_Btn_PreviewMouseMove(object sender, System.Windows.Input.MouseEventArgs e)
        {
            ClearScreen_Btn.ToolTip = FindResource("Clear_screen");
        }
        private void ClearScreen_Btn_PreviewGotKeyboardFocus(object sender, KeyboardFocusChangedEventArgs e)
        {
            ClearScreen_Btn.ToolTip = null;
        }

        private void Btn_New_Click(object sender, RoutedEventArgs e)
        {
            // Table Clear
            mDataGrid.DataContext = null;
            mDataTable.Clear();
            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "200", Event = "", Note = "" });
            mDataGrid.DataContext = mDataTable;

            ScriptName.Text = "";
            mEdit.StartEdit("");

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
            openFileDialog.Filter = "script files (*.txt)|*.txt"; // "txt files (*.txt)|*.txt|All files (*.*)|*.*"
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
        private async void Btn_Save_as_Click(object sender, RoutedEventArgs e)
        {
            Btn_ON.Content = "OFF";
            Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
            mNotifyIcon.Icon = OffIcon;

            var result = await this.ShowInputAsync(FindResource("Save").ToString(), FindResource("Input_filename").ToString());


            Btn_ON.Content = "ON";
            Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            mNotifyIcon.Icon = OnIcon;

            if (result == null) { return; }
            if (File.Exists(System.Windows.Forms.Application.StartupPath + "/" + result + ".txt")) {
                await this.ShowMessageAsync("", FindResource("Save_could_not_be_completed").ToString() 
                    + " " + FindResource("File_exists").ToString());
                return;
            }

            try
            {
                string JSON_String = JsonSerializer.Serialize(mDataTable.Select(i => new { i.Enable, i.Mode, i.Action, i.Event, i.Note }).ToList());
                JSON_String = JSON_String.Insert(1, "\n");
                JSON_String = JSON_String.Insert(JSON_String.Length - 1, "\n");
                JSON_String = JSON_String.Replace("\"},", "\"},\n");

                mEdit.ModifiedTime = "";
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
                string JSON_String = JsonSerializer.Serialize(mDataTable.Select(i => new { i.Enable, i.Mode, i.Action, i.Event, i.Note }).ToList());
                JSON_String = JSON_String.Insert(1, "\n");
                JSON_String = JSON_String.Insert(JSON_String.Length-1, "\n");
                JSON_String = JSON_String.Replace("\"},", "\"},\n");

                System.IO.File.WriteAllText(result, JSON_String);
                mEdit.ModifiedTime = mEdit.GetModifiedTime(result);

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
            mNotifyIcon.Icon = OnIcon;
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


                // ToolBar
                int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
                if (columnIndex < 0) { return; }

                string head = mDataGrid.Columns[columnIndex].Header.ToString();
                if (head.Equals("#"))
                {
                    ToolBar.Items.Clear();

                    MainTable row = (MainTable)mDataGrid.CurrentItem;
                    string[] btnlist = new string[] { };
                    switch (row.Mode)
                    {
                        case "Calc":
                        case "Calc-Check":
                        case "RemoveEvent":
                        case "Delay":
                        case "Jump":
                        case "Goto":
                        case "Loop":
                            
                            break;
                        default:
                            btnlist = new string[] { "Play" };
                            break;
                    }

                    if (btnlist.Length <= 0)
                    {
                        ToolBar.Visibility = Visibility.Collapsed;
                        return;
                    }

                    for (int i = 0; i < btnlist.Length; i++)
                    {
                        System.Windows.Controls.Button btn = new System.Windows.Controls.Button();
                        btn.Height = 30;
                        btn.Content = btnlist[i];
                        btn.Background = new SolidColorBrush(Colors.Orange);
                        btn.Foreground = new SolidColorBrush(Colors.DarkRed);
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
                            ToolBar.Margin = new Thickness(0, screenCoordinates.Y - 5,
                                (_Window.ActualWidth - mDataGrid.ActualWidth) / 2 + mDataGrid.ActualWidth - 20, 0);
                        }
                    }
                    ToolBar.Visibility = Visibility.Visible;
                }
            }
        }
        private void mDataGrid_KeyUp(object sender, System.Windows.Input.KeyEventArgs e)
        {
            if (e.Key.ToString().Equals("Return"))
            {
                CellEditComplete();
            }
        }
        private void mDataGrid_HeaderClick(object sender, RoutedEventArgs e)
        {       
            mDataGrid.DataContext = null;
            mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "", Event = "", Note = "" });
            mDataGrid.DataContext = mDataTable;
        }

        private void mDataGrid_BeginningEdit(object sender, DataGridBeginningEditEventArgs e)
        {
            int columnIndex = mDataGrid.Columns.IndexOf(mDataGrid.CurrentCell.Column);
            string head = mDataGrid.Columns[columnIndex].Header.ToString();

            if (!head.Equals("Enable") && !head.Equals(" ") && !head.Equals("+") ){
                Btn_ON.Content = "OFF";
                Btn_ON.Foreground = System.Windows.Media.Brushes.Red;
                mNotifyIcon.Icon = OffIcon;
            }

            // ToolBar
            ToolBar.Items.Clear();
            if (head.Equals("Action"))
            {
                MainTable row = (MainTable)mDataGrid.CurrentItem;
                string[] btnlist = new string[] { };
                string[] txtlist = new string[] { };

                switch (row.Mode)
                {
                    case "Click":
                        btnlist = new string[] { "Left", "Right", "Left_Down", "Left_Up", "Right_Down", "Right_Up" };
                        break;
                    case "RemoveEvent":
                        btnlist = new string[] { "PUSH" };
                        break;
                    case "Match":
                    case "Match RGB":
                    case "Match&Draw":
                    case "Run .exe":
                    case "PlaySound":
                        txtlist = new string[] { "※ " + string.Format(FindResource("Double_click_to_select_file").ToString()) };
                        break;
                }

                if ((btnlist.Length <= 0 && txtlist.Length <= 0) || row.Action.Length > 0)
                {
                    ToolBar.Visibility = Visibility.Collapsed;
                    return;
                }

                for (int i = 0; i < txtlist.Length; i++)
                {
                    System.Windows.Controls.Label label = new System.Windows.Controls.Label();
                    label.Height = 30;
                    label.Content = txtlist[i];
                    label.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(240, 240, 240, 240));
                    label.Foreground = new SolidColorBrush(Colors.Red);
                    ToolBar.Items.Add(label);
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

            if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Action") || mDataGrid.Columns[columnIndex].Header.ToString().Equals("#"))
            {
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);
                try
                {
                    if (tableIndex < mDataTable.Count())
                    {
                        if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Action")){
                            DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
                            FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                            System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;
                            System.Windows.Controls.Button mButton = (System.Windows.Controls.Button)sender;
                            System.Windows.Controls.TextBox mTextBlock = (System.Windows.Controls.TextBox)cell.Content;

                            mTextBlock.Text = mTextBlock.Text + mButton.Content;
                        }
                        else {
                            // Unit test
                            MainTable row = (MainTable)mDataGrid.CurrentItem;

                            UnitTest(row.Mode, row.Action);
                        }
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
                finally
                {
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
            CellEditComplete();
        }

        private void mDataGrid_LoadingRow(object sender, DataGridRowEventArgs e)
        {
            e.Row.Tag = (e.Row.GetIndex() + 1).ToString();
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
                        mNotifyIcon.Icon = OnIcon;
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
                        mDataTable.Insert(tableIndex + 1, new MainTable() { Enable = true, Mode = "Delay", Action = "", Event = "", Note = "" });
                        mDataGrid.DataContext = mDataTable;
                    }
                    else
                    {
                        mDataGrid.DataContext = null;
                        mDataTable.Add(new MainTable() { Enable = true, Mode = "Delay", Action = "", Event = "", Note = "" });
                        mDataGrid.DataContext = mDataTable;

                        mDataGrid_ScrollToBottom();
                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
            else if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Action"))
            {
                int tableIndex = mDataGrid.Items.IndexOf(mDataGrid.CurrentItem);

                try
                {
                    if (tableIndex < mDataTable.Count())
                    {
                        MainTable row = (MainTable)mDataGrid.CurrentItem;

                        if (row.Mode.Equals("Match") || row.Mode.Equals("Match RGB") || row.Mode.Equals("Match&Draw") || row.Mode.Equals("Run .exe") || row.Mode.Equals("PlaySound"))
                        {
                            string filePath = string.Empty;

                            System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                            openFileDialog.InitialDirectory = System.Windows.Forms.Application.StartupPath;
                            switch (row.Mode)
                            {
                                case "Match":
                                case "Match RGB":
                                case "Match&Draw":
                                    openFileDialog.Filter = "image files (*.png)|*.png";
                                    break;
                                case "Run .exe":
                                    openFileDialog.Filter = "exe files (*.exe)|*.exe";
                                    break;
                                case "PlaySound":
                                    openFileDialog.Filter = "sound files (*.wav)|*.wav";
                                    break;
                            }
                            openFileDialog.FilterIndex = 2;
                            openFileDialog.RestoreDirectory = true;
                            openFileDialog.ShowDialog();

                            //Get the path of specified file
                            filePath = openFileDialog.FileName;
                            if (filePath.Equals("")) { return; }

                            DataGridCellInfo cellInfo = mDataGrid.CurrentCell;
                            FrameworkElement cellContent = cellInfo.Column.GetCellContent(cellInfo.Item);
                            System.Windows.Controls.DataGridCell cell = cellContent.Parent as System.Windows.Controls.DataGridCell;
                            System.Windows.Controls.TextBox mTextBlock = (System.Windows.Controls.TextBox)cell.Content;

                            mTextBlock.Text = filePath.Replace(Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location) + "\\", "");

                            cell.IsEditing = false;
                            CellEditComplete();
                        }

                    }
                }
                catch (Exception err)
                {
                    Console.WriteLine("{0} Exception caught.", err);
                }
            }
            else if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("Enable")) {
                // Automatically turned on
                Btn_ON.Content = "ON";
                Btn_ON.Foreground = System.Windows.Media.Brushes.White;
                mNotifyIcon.Icon = OnIcon;
            }
            else if (mDataGrid.Columns[columnIndex].Header.ToString().Equals("#"))
            {
                // disable
                if (mDataGrid.EnableRowVirtualization == false && false) {
                    DataGridRow row = (DataGridRow)mDataGrid.ItemContainerGenerator.ContainerFromIndex(mDataGrid.Items.IndexOf(mDataGrid.CurrentItem));

                    System.Windows.Controls.DataGridCell cell = mDataGrid.Columns[0].GetCellContent(row).Parent as System.Windows.Controls.DataGridCell;
                    if (cell.Background.ToString().Equals("#64FF0000") && mDataGrid.EnableRowVirtualization == false)
                    {
                        cell.Background = System.Windows.Media.Brushes.Transparent;
                        cell.Foreground = System.Windows.Media.Brushes.Black;
                    }
                    else
                    {
                        cell.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(100, 255, 0, 0));
                    }
                }
            }
        }

        private void CellEditComplete() {
            ToolBar.Visibility = Visibility.Collapsed;
            // CommitEdit & Change Focus
            mDataGrid.CommitEdit();
            //ClearScreen_Btn.Focus();

            Btn_ON.Content = "ON";
            Btn_ON.Foreground = System.Windows.Media.Brushes.White;
            mNotifyIcon.Icon = OnIcon;
        }

        private void mDataGrid_ScrollToBottom()
        {
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

        private void CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.CheckBox mCheckBox = (System.Windows.Controls.CheckBox)sender;
            bool IsChecked = (bool)mCheckBox.IsChecked;

            switch (mCheckBox.Name)
            {
                case "OnOff_CrtlKey_Chk":
                    mSettingHelper.OnOff_CrtlKey = IsChecked;
                    break;
                case "OnOff_AltKey_Chk":
                    mSettingHelper.OnOff_AltKey = IsChecked;
                    break;
                case "Run_CrtlKey_Chk":
                    mSettingHelper.Run_CrtlKey = IsChecked;
                    break;
                case "Run_AltKey_Chk":
                    mSettingHelper.Run_AltKey = IsChecked;
                    break;
                case "Stop_CrtlKey_Chk":
                    mSettingHelper.Stop_CrtlKey = IsChecked;
                    break;
                case "Stop_AltKey_Chk":
                    mSettingHelper.Stop_AltKey = IsChecked;
                    break;
            }
        }

        private void TextBox_OnOff_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
          
            mSettingHelper.OnOff_Hotkey = ConvertHelper.ConvertKeyCode(TextBox_OnOff_Hotkey.Text.ToString());
        }

        private void TextBox_Run_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            mSettingHelper.Run_Hotkey = ConvertHelper.ConvertKeyCode(TextBox_Run_Hotkey.Text.ToString());
        }

        private void TextBox_Stop_Hotkey_TextChanged(object sender, TextChangedEventArgs e)
        {
            mSettingHelper.Stop_Hotkey = ConvertHelper.ConvertKeyCode(TextBox_Stop_Hotkey.Text.ToString());
        }

        private void HideOnSatrt_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            mSettingHelper.HideOnSatrt = HideOnSatrt_Toggle.IsOn;
        }

        private void TestMode_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (TestMode_Toggle.IsOn == true)
            {
                TestMode = true;
                mSettingHelper.TestMode = "1";
                //mDataGrid.EnableRowVirtualization = false;
            }
            else
            {
                TestMode = false;
                mSettingHelper.TestMode = "0";
                //mDataGrid.EnableRowVirtualization = true;
            }
        }

        private void TestMode_Slider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            TestMode_Delay = (int)TestMode_Slider.Value;
            mSettingHelper.TestMode_Delay = TestMode_Delay;
        }

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            System.Windows.Controls.RadioButton mRadioButton = (System.Windows.Controls.RadioButton)sender;
            mSettingHelper.TypeOfKeyboardInput = mRadioButton.Tag.ToString();
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

        private void Top_Toggle_Toggled(object sender, RoutedEventArgs e)
        {
            if (Top_Toggle.IsOn == true)
            {
                mSettingHelper.Topmost = true;
            }
            else
            {
                mSettingHelper.Topmost = false;
                System.Windows.Application.Current.MainWindow.Topmost = false;
            }
        }

        #endregion

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

        private String RunTemplateMatch(Mat rec, Mat template, string Mode, double mThreshold)
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

        #endregion
    }
}