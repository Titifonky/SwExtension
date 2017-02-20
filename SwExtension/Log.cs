using log4net;
using log4net.Appender;
using log4net.Config;
using log4net.Repository;
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Runtime.InteropServices;
using Outils;
using System.Windows.Forms;
using System.Collections.Generic;

namespace LogDebugging
{
    [Flags]
    internal enum LogLevelL4N
    {
        DEBUG = 1,
        ERROR = 2,
        FATAL = 4,
        INFO = 8,
        WARN = 16
    }

    internal static class Log
    {
        private static readonly ILog _Logger = LogManager.GetLogger("DLL");

        private static Boolean _EstInitialise = false;

        private static Boolean _Actif = true;

        static Log()
        {
            String Dossier = Path.GetDirectoryName(Assembly.GetAssembly(typeof(Log)).Location);
            String Chemin = Dossier + @"\" + "log4net.config";
            XmlConfigurator.Configure(_Logger.Logger.Repository, new FileInfo(Chemin));

            IAppender[] appenders = _Logger.Logger.Repository.GetAppenders();
            foreach (IAppender appender in appenders)
            {
                FileAppender fileAppender = appender as FileAppender;

                String CheminFichier = Path.Combine(Dossier, Path.GetFileName(fileAppender.File));
                if (File.Exists(CheminFichier))
                    File.Delete(CheminFichier);

                fileAppender.File = Path.Combine(Dossier, Path.GetFileName(fileAppender.File));
                fileAppender.ActivateOptions();
            }
        }

        internal static void Demarrer()
        {
            Activer = true;
            Entete();
        }

        internal static void Stopper()
        {
            IAppender[] appenders = _Logger.Logger.Repository.GetAppenders();
            foreach (IAppender appender in appenders)
            {
                appender.Close();
            }
            _Logger.Logger.Repository.Shutdown();
        }

        internal static void Entete()
        {
            if (_EstInitialise)
                return;

            Write("\n ");
            Write("====================================================================================================");
            Write("|                                                                                                  |");
            Write("|                                          DEBUG                                                   |");
            Write("|                                                                                                  |");
            Write("====================================================================================================");
            Write("\n ");

            _EstInitialise = true;
        }

        internal static Boolean Activer
        {
            get
            {
                return _Actif;
            }
            set
            {
                _Actif = value;

                log4net.Core.Level pLevel = log4net.Core.Level.Debug;
                if (value)
                    pLevel = log4net.Core.Level.All;

                ILoggerRepository repository = _Logger.Logger.Repository;
                repository.Threshold = pLevel;

                ((log4net.Repository.Hierarchy.Logger)_Logger.Logger).Level = pLevel;

                log4net.Repository.Hierarchy.Hierarchy h = (log4net.Repository.Hierarchy.Hierarchy)repository;
                log4net.Repository.Hierarchy.Logger rootLogger = h.Root;
                rootLogger.Level = pLevel;

            }
        }

        internal static void Write(Object Message, LogLevelL4N Level = LogLevelL4N.DEBUG)
        {
            try
            {
                if (Level.Equals(LogLevelL4N.DEBUG))
                    _Logger.Debug(Message.ToString());
                else if (Level.Equals(LogLevelL4N.ERROR))
                    _Logger.Error(Message.ToString());
                else if (Level.Equals(LogLevelL4N.FATAL))
                    _Logger.Fatal(Message.ToString());
                else if (Level.Equals(LogLevelL4N.INFO))
                    _Logger.Info(Message.ToString());
                else if (Level.Equals(LogLevelL4N.WARN))
                    _Logger.Warn(Message.ToString());
            }
            catch { }
        }

        internal static void Message(Object message)
        {
            if (!_Actif)
                return;

            Write("\t\t\t\t-> " + message.ToString());
        }

        internal static void LogMethode(this Object O, Object[] Message, [CallerMemberName] String methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t" + O.GetType().Name + "." + methode + "  ->  " + String.Join(" ", Message));
        }

        internal static void LogMethode(this Object O, [CallerMemberName] String methode = "")
        {
            Methode(O.GetType().Name, methode);
        }

        internal static void LogResultat(this Object O, String Text, [CallerMemberName] String methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t Resultat dans " + methode + "  " + Text + " : " + O.ToString());
        }

        internal static void Methode(String nomClasse, [CallerMemberName] String methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t" + nomClasse + "." + methode);
        }

        internal static void Methode(String nomClasse, Object message, [CallerMemberName] String methode = "")
        {
            if (!_Actif)
                return;

            Write("\t\t\t" + nomClasse + "." + methode);
            if (message != null)
                Write("\t\t\t\t-> " + message.ToString());
        }

        internal static void Methode<T>([CallerMemberName] String methode = "")
        {
            Methode(typeof(T).ToString(), methode);
        }

        internal static void Methode<T>(Object message, [CallerMemberName] String methode = "")
        {
            Methode(typeof(T).ToString(), message, methode);
        }
    }

    internal static class WindowLog
    {
        private static SwExtension.NotePad _NotePad = null;

        public static void AfficherFenetre(Boolean a = true)
        {
            if (a)
            {
                _NotePad.Show();
                System.Drawing.Point Pt = new System.Drawing.Point(Cursor.Position.X - _NotePad.Width, Cursor.Position.Y);
                _NotePad.Location = Pt;
            }
            else
                _NotePad.Close();
        }

        static WindowLog()
        {
            if (_NotePad.IsNull())
                _NotePad = new SwExtension.NotePad();
        }

        private static DateTime pStart;
        private static DateTime pDiff;

        internal static void Ecrire(Object o)
        {
            Ecrire(o.ToString());
        }

        internal static void Ecrire(string Message)
        {
            Afficher();
            Text(Message, true);
            
            String T = String.Format("{0:hh\\:mm\\:ss}", DateTime.Now) + " " +
                                    Remplir(Math.Round((DateTime.Now - pStart).TotalSeconds).ToString()) + " " +
                                    Remplir(Math.Round((DateTime.Now - pDiff).TotalSeconds).ToString());
            pDiff = DateTime.Now;

            String ligne = String.Format("{0}  {1}", T, Message);

            _NotePad.AppendText(ligne + Environment.NewLine);

            Log.Message("## " + ligne);
        }

        internal static void Ecrire(List<String> listeMessage)
        {
            foreach (String s in listeMessage)
                Ecrire(s);
        }

        internal static void EcrireF(string Message, params Object[] objs)
        {
            Ecrire(String.Format(Message, objs));
        }

        internal static void SautDeLigne(int nb = 1)
        {
            for (int i = 0; i < nb; i++)
            {
                Ecrire("");
            }
        }

        internal static void Effacer()
        {
            Text("", false);
            _NotePad.Effacer();

            pStart = DateTime.Now;
            pDiff = DateTime.Now;
        }

        private static String Remplir(String s)
        {
            return " ".eRepeter(6 - s.Length) + s;
        }

        internal delegate void TextEvent(String s, Boolean Ajouter);

        internal static TextEvent Text;

        internal delegate void AfficherEvent();

        internal static AfficherEvent Afficher;
    }

    //private static class NotePad
    //{
    //    private enum SystemMetric
    //    {
    //        SM_CXSCREEN = 0,  // 0x00
    //        SM_CYSCREEN = 1,  // 0x01
    //        SM_CXVSCROLL = 2,  // 0x02
    //        SM_CYHSCROLL = 3,  // 0x03
    //        SM_CYCAPTION = 4,  // 0x04
    //        SM_CXBORDER = 5,  // 0x05
    //        SM_CYBORDER = 6,  // 0x06
    //        SM_CXDLGFRAME = 7,  // 0x07
    //        SM_CXFIXEDFRAME = 7,  // 0x07
    //        SM_CYDLGFRAME = 8,  // 0x08
    //        SM_CYFIXEDFRAME = 8,  // 0x08
    //        SM_CYVTHUMB = 9,  // 0x09
    //        SM_CXHTHUMB = 10, // 0x0A
    //        SM_CXICON = 11, // 0x0B
    //        SM_CYICON = 12, // 0x0C
    //        SM_CXCURSOR = 13, // 0x0D
    //        SM_CYCURSOR = 14, // 0x0E
    //        SM_CYMENU = 15, // 0x0F
    //        SM_CXFULLSCREEN = 16, // 0x10
    //        SM_CYFULLSCREEN = 17, // 0x11
    //        SM_CYKANJIWINDOW = 18, // 0x12
    //        SM_MOUSEPRESENT = 19, // 0x13
    //        SM_CYVSCROLL = 20, // 0x14
    //        SM_CXHSCROLL = 21, // 0x15
    //        SM_DEBUG = 22, // 0x16
    //        SM_SWAPBUTTON = 23, // 0x17
    //        SM_CXMIN = 28, // 0x1C
    //        SM_CYMIN = 29, // 0x1D
    //        SM_CXSIZE = 30, // 0x1E
    //        SM_CYSIZE = 31, // 0x1F
    //        SM_CXSIZEFRAME = 32, // 0x20
    //        SM_CXFRAME = 32, // 0x20
    //        SM_CYSIZEFRAME = 33, // 0x21
    //        SM_CYFRAME = 33, // 0x21
    //        SM_CXMINTRACK = 34, // 0x22
    //        SM_CYMINTRACK = 35, // 0x23
    //        SM_CXDOUBLECLK = 36, // 0x24
    //        SM_CYDOUBLECLK = 37, // 0x25
    //        SM_CXICONSPACING = 38, // 0x26
    //        SM_CYICONSPACING = 39, // 0x27
    //        SM_MENUDROPALIGNMENT = 40, // 0x28
    //        SM_PENWINDOWS = 41, // 0x29
    //        SM_DBCSENABLED = 42, // 0x2A
    //        SM_CMOUSEBUTTONS = 43, // 0x2B
    //        SM_SECURE = 44, // 0x2C
    //        SM_CXEDGE = 45, // 0x2D
    //        SM_CYEDGE = 46, // 0x2E
    //        SM_CXMINSPACING = 47, // 0x2F
    //        SM_CYMINSPACING = 48, // 0x30
    //        SM_CXSMICON = 49, // 0x31
    //        SM_CYSMICON = 50, // 0x32
    //        SM_CYSMCAPTION = 51, // 0x33
    //        SM_CXSMSIZE = 52, // 0x34
    //        SM_CYSMSIZE = 53, // 0x35
    //        SM_CXMENUSIZE = 54, // 0x36
    //        SM_CYMENUSIZE = 55, // 0x37
    //        SM_ARRANGE = 56, // 0x38
    //        SM_CXMINIMIZED = 57, // 0x39
    //        SM_CYMINIMIZED = 58, // 0x3A
    //        SM_CXMAXTRACK = 59, // 0x3B
    //        SM_CYMAXTRACK = 60, // 0x3C
    //        SM_CXMAXIMIZED = 61, // 0x3D
    //        SM_CYMAXIMIZED = 62, // 0x3E
    //        SM_NETWORK = 63, // 0x3F
    //        SM_CLEANBOOT = 67, // 0x43
    //        SM_CXDRAG = 68, // 0x44
    //        SM_CYDRAG = 69, // 0x45
    //        SM_SHOWSOUNDS = 70, // 0x46
    //        SM_CXMENUCHECK = 71, // 0x47
    //        SM_CYMENUCHECK = 72, // 0x48
    //        SM_SLOWMACHINE = 73, // 0x49
    //        SM_MIDEASTENABLED = 74, // 0x4A
    //        SM_MOUSEWHEELPRESENT = 75, // 0x4B
    //        SM_XVIRTUALSCREEN = 76, // 0x4C
    //        SM_YVIRTUALSCREEN = 77, // 0x4D
    //        SM_CXVIRTUALSCREEN = 78, // 0x4E
    //        SM_CYVIRTUALSCREEN = 79, // 0x4F
    //        SM_CMONITORS = 80, // 0x50
    //        SM_SAMEDISPLAYFORMAT = 81, // 0x51
    //        SM_IMMENABLED = 82, // 0x52
    //        SM_CXFOCUSBORDER = 83, // 0x53
    //        SM_CYFOCUSBORDER = 84, // 0x54
    //        SM_TABLETPC = 86, // 0x56
    //        SM_MEDIACENTER = 87, // 0x57
    //        SM_STARTER = 88, // 0x58
    //        SM_SERVERR2 = 89, // 0x59
    //        SM_MOUSEHORIZONTALWHEELPRESENT = 91, // 0x5B
    //        SM_CXPADDEDBORDER = 92, // 0x5C
    //        SM_DIGITIZER = 94, // 0x5E
    //        SM_MAXIMUMTOUCHES = 95, // 0x5F

    //        SM_REMOTESESSION = 0x1000, // 0x1000
    //        SM_SHUTTINGDOWN = 0x2000, // 0x2000
    //        SM_REMOTECONTROL = 0x2001, // 0x2001


    //        SM_CONVERTABLESLATEMODE = 0x2003,
    //        SM_SYSTEMDOCKED = 0x2004,
    //    }

    //    [Flags()]
    //    private enum SetWindowPosFlags : uint
    //    {
    //        /// <summary>If the calling thread and the thread that owns the window are attached to different input queues,
    //        /// the system posts the request to the thread that owns the window. This prevents the calling thread from
    //        /// blocking its execution while other threads process the request.</summary>
    //        /// <remarks>SWP_ASYNCWINDOWPOS</remarks>
    //        AsynchronousWindowPosition = 0x4000,
    //        /// <summary>Prevents generation of the WM_SYNCPAINT message.</summary>
    //        /// <remarks>SWP_DEFERERASE</remarks>
    //        DeferErase = 0x2000,
    //        /// <summary>Draws a frame (defined in the window's class description) around the window.</summary>
    //        /// <remarks>SWP_DRAWFRAME</remarks>
    //        DrawFrame = 0x0020,
    //        /// <summary>Applies new frame styles set using the SetWindowLong function. Sends a WM_NCCALCSIZE message to
    //        /// the window, even if the window's size is not being changed. If this flag is not specified, WM_NCCALCSIZE
    //        /// is sent only when the window's size is being changed.</summary>
    //        /// <remarks>SWP_FRAMECHANGED</remarks>
    //        FrameChanged = 0x0020,
    //        /// <summary>Hides the window.</summary>
    //        /// <remarks>SWP_HIDEWINDOW</remarks>
    //        HideWindow = 0x0080,
    //        /// <summary>Does not activate the window. If this flag is not set, the window is activated and moved to the
    //        /// top of either the topmost or non-topmost group (depending on the setting of the hWndInsertAfter
    //        /// parameter).</summary>
    //        /// <remarks>SWP_NOACTIVATE</remarks>
    //        DoNotActivate = 0x0010,
    //        /// <summary>Discards the entire contents of the client area. If this flag is not specified, the valid
    //        /// contents of the client area are saved and copied back into the client area after the window is sized or
    //        /// repositioned.</summary>
    //        /// <remarks>SWP_NOCOPYBITS</remarks>
    //        DoNotCopyBits = 0x0100,
    //        /// <summary>Retains the current position (ignores X and Y parameters).</summary>
    //        /// <remarks>SWP_NOMOVE</remarks>
    //        IgnoreMove = 0x0002,
    //        /// <summary>Does not change the owner window's position in the Z order.</summary>
    //        /// <remarks>SWP_NOOWNERZORDER</remarks>
    //        DoNotChangeOwnerZOrder = 0x0200,
    //        /// <summary>Does not redraw changes. If this flag is set, no repainting of any kind occurs. This applies to
    //        /// the client area, the nonclient area (including the title bar and scroll bars), and any part of the parent
    //        /// window uncovered as a result of the window being moved. When this flag is set, the application must
    //        /// explicitly invalidate or redraw any parts of the window and parent window that need redrawing.</summary>
    //        /// <remarks>SWP_NOREDRAW</remarks>
    //        DoNotRedraw = 0x0008,
    //        /// <summary>Same as the SWP_NOOWNERZORDER flag.</summary>
    //        /// <remarks>SWP_NOREPOSITION</remarks>
    //        DoNotReposition = 0x0200,
    //        /// <summary>Prevents the window from receiving the WM_WINDOWPOSCHANGING message.</summary>
    //        /// <remarks>SWP_NOSENDCHANGING</remarks>
    //        DoNotSendChangingEvent = 0x0400,
    //        /// <summary>Retains the current size (ignores the cx and cy parameters).</summary>
    //        /// <remarks>SWP_NOSIZE</remarks>
    //        IgnoreResize = 0x0001,
    //        /// <summary>Retains the current Z order (ignores the hWndInsertAfter parameter).</summary>
    //        /// <remarks>SWP_NOZORDER</remarks>
    //        IgnoreZOrder = 0x0004,
    //        /// <summary>Displays the window.</summary>
    //        /// <remarks>SWP_SHOWWINDOW</remarks>
    //        ShowWindow = 0x0040,
    //    }

    //    private enum GetWindowType : uint
    //    {
    //        /// <summary>
    //        /// The retrieved handle identifies the window of the same type that is highest in the Z order.
    //        /// <para/>
    //        /// If the specified window is a topmost window, the handle identifies a topmost window.
    //        /// If the specified window is a top-level window, the handle identifies a top-level window.
    //        /// If the specified window is a child window, the handle identifies a sibling window.
    //        /// </summary>
    //        GW_HWNDFIRST = 0,
    //        /// <summary>
    //        /// The retrieved handle identifies the window of the same type that is lowest in the Z order.
    //        /// <para />
    //        /// If the specified window is a topmost window, the handle identifies a topmost window.
    //        /// If the specified window is a top-level window, the handle identifies a top-level window.
    //        /// If the specified window is a child window, the handle identifies a sibling window.
    //        /// </summary>
    //        GW_HWNDLAST = 1,
    //        /// <summary>
    //        /// The retrieved handle identifies the window below the specified window in the Z order.
    //        /// <para />
    //        /// If the specified window is a topmost window, the handle identifies a topmost window.
    //        /// If the specified window is a top-level window, the handle identifies a top-level window.
    //        /// If the specified window is a child window, the handle identifies a sibling window.
    //        /// </summary>
    //        GW_HWNDNEXT = 2,
    //        /// <summary>
    //        /// The retrieved handle identifies the window above the specified window in the Z order.
    //        /// <para />
    //        /// If the specified window is a topmost window, the handle identifies a topmost window.
    //        /// If the specified window is a top-level window, the handle identifies a top-level window.
    //        /// If the specified window is a child window, the handle identifies a sibling window.
    //        /// </summary>
    //        GW_HWNDPREV = 3,
    //        /// <summary>
    //        /// The retrieved handle identifies the specified window's owner window, if any.
    //        /// </summary>
    //        GW_OWNER = 4,
    //        /// <summary>
    //        /// The retrieved handle identifies the child window at the top of the Z order,
    //        /// if the specified window is a parent window; otherwise, the retrieved handle is NULL.
    //        /// The function examines only child windows of the specified window. It does not examine descendant windows.
    //        /// </summary>
    //        GW_CHILD = 5,
    //        /// <summary>
    //        /// The retrieved handle identifies the enabled popup window owned by the specified window (the
    //        /// search uses the first such window found using GW_HWNDNEXT); otherwise, if there are no enabled
    //        /// popup windows, the retrieved handle is that of the specified window.
    //        /// </summary>
    //        GW_ENABLEDPOPUP = 6
    //    }

    //    [Flags]
    //    private enum ProcessAccessFlags : uint
    //    {
    //        All = 0x001F0FFF,
    //        Terminate = 0x00000001,
    //        CreateThread = 0x00000002,
    //        VirtualMemoryOperation = 0x00000008,
    //        VirtualMemoryRead = 0x00000010,
    //        VirtualMemoryWrite = 0x00000020,
    //        DuplicateHandle = 0x00000040,
    //        CreateProcess = 0x000000080,
    //        SetQuota = 0x00000100,
    //        SetInformation = 0x00000200,
    //        QueryInformation = 0x00000400,
    //        QueryLimitedInformation = 0x00001000,
    //        Synchronize = 0x00100000
    //    }

    //    /// <summary>
    //    ///     Special window handles
    //    /// </summary>
    //    public enum SpecialWindowHandles
    //    {
    //        // ReSharper disable InconsistentNaming
    //        /// <summary>
    //        ///     Places the window at the top of the Z order.
    //        /// </summary>
    //        HWND_TOP = 0,
    //        /// <summary>
    //        ///     Places the window at the bottom of the Z order. If the hWnd parameter identifies a topmost window, the window loses its topmost status and is placed at the bottom of all other windows.
    //        /// </summary>
    //        HWND_BOTTOM = 1,
    //        /// <summary>
    //        ///     Places the window above all non-topmost windows. The window maintains its topmost position even when it is deactivated.
    //        /// </summary>
    //        HWND_TOPMOST = -1,
    //        /// <summary>
    //        ///     Places the window above all non-topmost windows (that is, behind all topmost windows). This flag has no effect if the window is already a non-topmost window.
    //        /// </summary>
    //        HWND_NOTOPMOST = -2
    //        // ReSharper restore InconsistentNaming
    //    }

    //    [DllImport("user32.dll", EntryPoint = "FindWindowA")]
    //    private static extern IntPtr FindWindow(String lpClassName, String lpWindowName);

    //    [DllImport("user32.dll", SetLastError = true)]
    //    private static extern IntPtr FindWindowEx(IntPtr parentHandle, IntPtr childAfter, string className, string windowTitle);

    //    [DllImport("User32.dll")]
    //    private static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, String lParam);

    //    [DllImport("user32.dll")]
    //    private static extern int GetSystemMetrics(SystemMetric smIndex);

    //    [DllImport("user32.dll", SetLastError = true)]
    //    private static extern bool SetWindowPos(IntPtr hWnd, SpecialWindowHandles hWndInsertAfter, int X, int Y, int cx, int cy, SetWindowPosFlags uFlags);

    //    [DllImport("user32.dll", ExactSpelling = true, CharSet = CharSet.Auto)]
    //    private static extern IntPtr GetParent(IntPtr hWnd);

    //    [DllImport("user32.dll", SetLastError = true)]
    //    private static extern IntPtr GetWindow(IntPtr hWnd, GetWindowType uCmd);

    //    // When you don't want the ProcessId, use this overload and pass IntPtr.Zero for the second parameter
    //    [DllImport("user32.dll", SetLastError = true)]
    //    private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    //    [return: MarshalAs(UnmanagedType.Bool)]
    //    [DllImport("user32.dll")]
    //    private static extern bool IsWindow(IntPtr hWnd);

    //    private const int WM_SETTEXT = 0x000c;

    //    private static int IdNotepad;
    //    private static IntPtr hwnd;
    //    private static IntPtr chWnd;
    //    private static Process NotePadproc;

    //    private static String pTexte;

    //    private static DateTime pStart;
    //    private static DateTime pDiff;

    //    private static Boolean _Afficher = false;

    //    public static void Afficher(Boolean a = true)
    //    {
    //        _Afficher = a;
    //        if (a)
    //            Init();
    //        else
    //            Fermer();
    //    }

    //    private static void Init()
    //    {
    //        if (!IsWindow(hwnd))
    //        {
    //            NotePadproc = new Process();
    //            NotePadproc.StartInfo.FileName = "NOTEPAD.EXE";
    //            NotePadproc.Start();
    //            NotePadproc.WaitForInputIdle();
    //            IdNotepad = NotePadproc.Id;
    //            hwnd = NotePadproc.MainWindowHandle;

    //            //hwnd = GetWinHandle(IdNotepad);
    //            chWnd = FindWindowEx(hwnd, new IntPtr(0), "", "");
    //            int Ht = GetSystemMetrics(SystemMetric.SM_CYSCREEN);
    //            int Lg = GetSystemMetrics(SystemMetric.SM_CXSCREEN);
    //            int LgF = 500;
    //            int HtF = Ht;
    //            SetWindowPos(hwnd, SpecialWindowHandles.HWND_NOTOPMOST, Lg - LgF, 0, LgF, HtF - 45, 0);
    //        }
    //    }

    //    private static void Fermer()
    //    {
    //        if (NotePadproc.IsRef())
    //            NotePadproc.Kill();
    //    }

    //    public static void Start()
    //    {
    //        pStart = DateTime.Now;
    //        pDiff = DateTime.Now;
    //        pTexte = Environment.NewLine + "Start";
    //    }

    //    public static void Ecrire(String Texte)
    //    {
    //        String T = String.Format("{0:hh\\:mm\\:ss}", DateTime.Now) + " " +
    //                        Remplir((DateTime.Now - pStart).Seconds.ToString()) + " " +
    //                        Remplir((DateTime.Now - pDiff).Seconds.ToString()) + " ";
    //        pDiff = DateTime.Now;

    //        pTexte += T + Texte + Environment.NewLine;

    //        if (_Afficher)
    //        {
    //            Init();
    //            SendMessage(chWnd, WM_SETTEXT, 0, pTexte);
    //        }
    //    }

    //    public static void Effacer()
    //    {
    //        pTexte = "";

    //        if (!IsWindow(hwnd)) return;
    //        SendMessage(chWnd, WM_SETTEXT, 0, pTexte);
    //    }

    //    private static String Remplir(String s)
    //    {
    //        return " ".eRepeter(6 - s.Length) + s;
    //    }

    //    private static uint ProcIDFromWnd(IntPtr hwnd)
    //    {
    //        uint idProc;


    //        // Get PID for this HWnd
    //        GetWindowThreadProcessId(hwnd, out idProc);


    //        // Return PID
    //        return idProc;
    //    }

    //    private static IntPtr GetWinHandle(int hInstance)
    //    {
    //        IntPtr tempHwnd = new IntPtr(0);

    //        // Grab the first window handle that Windows finds:
    //        tempHwnd = FindWindow("", "");

    //        do
    //        {
    //            if (GetParent(tempHwnd) == new IntPtr(0))
    //            {
    //                // Check for PID match
    //                if (hInstance == ProcIDFromWnd(tempHwnd))
    //                    break;
    //            }
    //            // Get the next window handle
    //            tempHwnd = GetWindow(tempHwnd, GetWindowType.GW_HWNDNEXT);
    //        } while (tempHwnd != new IntPtr(0));

    //        return tempHwnd;
    //    }
    //}
}

