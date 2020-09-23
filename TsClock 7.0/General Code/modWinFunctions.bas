Attribute VB_Name = "modWinFunctions"
'=================================================
'AUTHOR   : Eric O'Sullivan
' -----------------------------------------------
'DATE     : 2 September 1999
' -----------------------------------------------
'CONTACT  : DiskJunky@hotmail.com
' -----------------------------------------------
'TITLE    : General Window Related Functions
' -----------------------------------------------
'COMMENTS :
'These code functions are generally associated
'with changing windows settings, or for example
'getting the computer to shutdown.
'=================================================

'erquire variable declaration
Option Explicit

'------------------------------------------------
'               API DECLARATIONS
'------------------------------------------------

'adjust the privilages for the current user to allow
'them access to certain features
Private Declare Function AdjustTokenPrivileges _
        Lib "advapi32" _
            (ByVal TokenHandle As Long, _
             ByVal DisableAllPrivileges As Long, _
             NewState As TOKEN_PRIVILEGES, _
             ByVal BufferLength As Long, _
             PreviousState As TOKEN_PRIVILEGES, _
             ReturnLength As Long) _
             As Long

'create a sound using the pc speaker - this does NOT
'work on win 9x systems
Private Declare Function Beep _
        Lib "kernel32" _
            (ByVal dwFreq As Long, _
             ByVal dwDuration As Long) _
             As Long

'This function closes an open object handle
Private Declare Function CloseHandle _
        Lib "kernel32.dll" _
            (ByVal Handle As Long) _
             As Long

'Takes a snapshot of the processes and the heaps, modules, and threads
'used by the processes.
Private Declare Function CreateToolhelp32Snapshot _
        Lib "kernel32" _
            (ByVal dwFlags As Long, _
             ByVal th32ProcessID As Long) _
             As Long

'this will remove the specified menu item from the menu
Private Declare Function DeleteMenu _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long) _
             As Long

'redraws the menus title bar and menu
Private Declare Function DrawMenuBar _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'The EnumProcesses function retrieves the process identifier for each
'process object in the system.
Private Declare Function EnumProcesses _
        Lib "PSAPI.DLL" _
            (ByRef lpidProcess As Long, _
             ByVal cb As Long, _
             ByRef cbNeeded As Long) _
             As Long

'The EnumProcessModules function retrieves a handle for each module in
'the specified process.
Private Declare Function EnumProcessModules _
        Lib "PSAPI.DLL" _
            (ByVal hProcess As Long, _
             ByRef lphModule As Long, _
             ByVal cb As Long, _
             ByRef cbNeeded As Long) _
             As Long

'shut down windows using the specified method
Private Declare Function ExitWindowsEx _
        Lib "user32" _
            (ByVal uFlags As Long, _
             ByVal dwReserved As Long) _
             As Long

'finds the first window in the queue with the caption matching the specified
'null termimated string.
Private Declare Function FindWindow _
        Lib "user32" _
        Alias "FindWindowA" _
            (ByVal lpClassName As String, _
             ByVal lpWindowName As String) _
             As Long

'finds the first CHILD window in the queue with the caption matching the specified
'null termimated string. Only works for MDI forms and children
Private Declare Function FindWindowEx _
        Lib "user32" _
        Alias "FindWindowExA" _
            (ByVal hWnd1 As Long, _
             ByVal hWnd2 As Long, _
             ByVal lpsz1 As String, _
             ByVal lpsz2 As String) _
             As Long

'get the class name of the specified window
Private Declare Function GetClassName _
        Lib "user32" _
        Alias "GetClassNameA" _
            (ByVal hWnd As Long, _
             ByVal lpClassName As String, _
             ByVal nMaxCount As Long) _
             As Long

'get a process pointer to the current process
Private Declare Function GetCurrentProcess _
        Lib "kernel32" _
            () _
             As Long

'get the id of the current process
Private Declare Function GetCurrentProcessId _
        Lib "kernel32" _
            () _
             As Long

'gets the current thread
Private Declare Function GetCurrentThread _
        Lib "kernel32" () As Long

'retrieves a handle to the desktop window
Private Declare Function GetDesktopWindow _
        Lib "user32" () _
                      As Long

'The GetModuleFileName function retrieves the full path and filename for
'the executable file containing the specified module.
Private Declare Function GetModuleFileNameExA _
        Lib "PSAPI.DLL" _
            (ByVal hProcess As Long, _
             ByVal hModule As Long, _
             ByVal ModuleName As String, _
             ByVal nSize As Long) _
             As Long

'get the class priority information for the specified process
Private Declare Function GetPriorityClass _
        Lib "kernel32" _
            (ByVal hProcess As Long) _
             As Long

'The GetProcessMemoryInfo function retrieves information about the memory
'usage of the specified process in the PROCESS_MEMORY_COUNTERS structure.
Private Declare Function GetProcessMemoryInfo _
        Lib "PSAPI.DLL" _
            (ByVal hProcess As Long, _
             ppsmemCounters As PROCESS_MEMORY_COUNTERS, _
             ByVal cb As Long) _
             As Long

'this will return a handle to the System Menu of the specified window
Private Declare Function GetSystemMenu _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal bRevert As Long) _
             As Long

'Get the priority level of the specified thread
Private Declare Function GetThreadPriority _
        Lib "kernel32" _
            (ByVal hThread As Long) _
             As Long

'get the current amount of milliseconds that windows
'has been acitve
Private Declare Function GetTickCount _
        Lib "kernel32" _
            () _
             As Long

'get the current top-most window (not necessarily
'the one with the current focus
Private Declare Function GetTopWindow _
        Lib "user32" _
            (ByVal hWnd As Long) _
             As Long

'get information about the current operating system
Private Declare Function GetVersionEx _
        Lib "kernel32" _
        Alias "GetVersionExA" _
            (ByRef lpVersionInformation As OSVERSIONINFO) _
             As Long

'gets the specified window in relation to the window specified
Private Declare Function GetWindow _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal wCmd As Long) _
             As Long

'get information about the specified window
Private Declare Function GetWindowLong _
        Lib "user32" _
        Alias "GetWindowLongA" _
            (ByVal hWnd As Long, _
             ByVal nIndex As Long) _
             As Long

'get the complete path to the windows directory
Private Declare Function GetWindowsDirectory _
        Lib "kernel32" _
        Alias "GetWindowsDirectoryA" _
            (ByVal lpBuffer As String, _
             ByVal nSize As Long) _
             As Long

'get the text caption of the specified window
Private Declare Function GetWindowText _
        Lib "user32" _
        Alias "GetWindowTextA" _
            (ByVal hWnd As Long, _
             ByVal lpString As String, _
             ByVal cch As Long) _
             As Long

'The GlobalMemoryStatus function retrieves information about current available
'memory. The function returns information about both physical and virtual
'memory. This function supersedes the GetFreeSpace function.
Private Declare Sub GlobalMemoryStatus _
        Lib "kernel32" _
            (lpBuffer As MEMORYSTATUS)

'destroyes the specified timer
Private Declare Function KillTimer _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal nIDEvent As Long) _
             As Long

'This will lock the computer for the current user. Only works on machines that
'are win2000+
Private Declare Function LockWorkStation _
        Lib "user32" _
            () As Long

'check what the current user is allowed to do
Private Declare Function LookupPrivilegeValue _
        Lib "advapi32" _
        Alias "LookupPrivilegeValueA" _
            (ByVal lpSystemName As String, _
             ByVal lpName As String, _
             lpLuid As LUID) _
             As Long

'opens an existing process object
Private Declare Function OpenProcess _
        Lib "kernel32.dll" _
            (ByVal dwDesiredAccessas As Long, _
             ByVal bInheritHandle As Long, _
             ByVal dwProcId As Long) _
             As Long

'request access to a specified process
Private Declare Function OpenProcessToken _
        Lib "advapi32" _
            (ByVal ProcessHandle As Long, _
             ByVal DesiredAccess As Long, _
             TokenHandle As Long) _
             As Long

'gets the first process in the list (95/98 only)
Private Declare Function Process32First _
        Lib "kernel32" _
            (ByVal hSnapshot As Long, _
             lppe As PROCESSENTRY32) _
             As Long

'gets the next process in the list (95/98 only)
Private Declare Function Process32Next _
        Lib "kernel32" _
            (ByVal hSnapshot As Long, _
             lppe As PROCESSENTRY32) _
             As Long

'used to set a hotkey
Private Declare Function RegisterHotKey _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal id As Long, _
             ByVal fsModifiers As Long, _
             ByVal vk As Long) _
             As Long

'set the specified service as the specified type
Private Declare Function RegisterServiceProcess _
        Lib "kernel32" _
            (ByVal dwProcessId As Long, _
             ByVal dwType As Long) _
             As Long

'this will remove a part of the system menu from a window
Private Declare Function RemoveMenu _
        Lib "user32" _
            (ByVal hMenu As Long, _
             ByVal nPosition As Long, _
             ByVal wFlags As Long) _
             As Long

'creates a timer with the specified time out value
Private Declare Function SetTimer _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal nIDEvent As Long, _
             ByVal uElapse As Long, _
             ByVal lpTimerFunc As Long) _
             As Long

'send events to the specified window
Private Declare Function SendMessage _
        Lib "user32" _
        Alias "SendMessageA" _
            (ByVal hWnd As Long, _
             ByVal wMsg As Long, _
             ByVal wParam As Integer, _
             ByVal lParam As Long) _
             As Long

'set the windows focus to the specified window
Private Declare Function SetFocusAPI _
        Lib "user32" _
        Alias "SetFocus" _
            (ByVal hWnd As Long) _
             As Long

'this will set the windows layer attributes. This needs to
'be updated with the api UpdateLayeredWindow. Works only
'on W2000 and WXP >+
Private Declare Function SetLayeredWindowAttributes _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal crKey As Long, _
             ByVal bAlpha As Byte, _
             ByVal dwFlags As Long) _
             As Long

'This will set the class priority for the specified process
Private Declare Function SetPriorityClass _
        Lib "kernel32" _
            (ByVal hProcess As Long, _
             ByVal dwPriorityClass As Long) _
             As Long

'set the priority level of the specified thread
Private Declare Function SetThreadPriority _
        Lib "kernel32" _
            (ByVal hThread As Long, _
             ByVal nPriority As Long) _
             As Long

'set information about the specified window
Private Declare Function SetWindowLong _
        Lib "user32" _
            Alias "SetWindowLongA" _
                (ByVal hWnd As Long, _
                 ByVal nIndex As Long, _
                 ByVal dwNewLong As Long) _
                 As Long

'set the position of the specified window
Private Declare Function SetWindowPos _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal hWndInsertAfter As Long, _
             ByVal x As Long, _
             ByVal Y As Long, _
             ByVal cx As Long, _
             ByVal cy As Long, _
             ByVal wFlags As Long) _
             As Long

'display the specified window
Private Declare Function ShowWindow _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal nCmdShow As Long) _
             As Long

'return or set system information
Private Declare Function SystemParametersInfo _
        Lib "user32" _
        Alias "SystemParametersInfoA" _
            (ByVal uAction As Long, _
             ByVal uParam As Long, _
             ByRef lpvParam As Any, _
             ByVal fuWinIni As Long) _
             As Long

'destroyes the specified process
Private Declare Function TerminateProcess _
        Lib "kernel32" _
            (ByVal hProcess As Long, _
             ByVal uExitCode As Long) _
             As Long

'remove the specified registered hotkey for the
'specified window
Private Declare Function UnregisterHotKey _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal id As Long) _
             As Long

'This will update the window with the new values specified
'by SetLayedWindowAttributes. Works on W2000, Wxp >+
Private Declare Function UpdateLayeredWindow _
        Lib "user32" _
            (ByVal hWnd As Long, _
             ByVal hdcDst As Long, _
             pptDst As Any, _
             psize As Any, _
             ByVal hdcSrc As Long, _
             pptSrc As Any, _
             crKey As Long, _
             ByVal pblend As Long, _
             ByVal dwFlags As Long) _
             As Long

'------------------------------------------------
'               USER-DEFINED TYPES
'------------------------------------------------
Private Type OSVERSIONINFO
    dwOSVersionInfoSize         As Long
    dwMajorVersion              As Long
    dwMinorVersion              As Long
    dwBuildNumber               As Long
    dwPlatformId                As Long
    szCSDVersion                As String * 128
End Type

Private Type LUID
    LowPart                     As Long
    HighPart                    As Long
End Type

Private Type LUID_AND_ATTRIBUTES
    pLuid                       As LUID
    Attributes                  As Long
End Type

Private Type TOKEN_PRIVILEGES
    PrivilegeCount              As Long
    Privileges(1)               As LUID_AND_ATTRIBUTES    'we only want one privilage
End Type

Private Type MEMORYSTATUS
    dwLength                    As Long
    dwMemoryLoad                As Long
    dwTotalPhys                 As Long
    dwAvailPhys                 As Long
    dwTotalPageFile             As Long
    dwAvailPageFile             As Long
    dwTotalVirtual              As Long
    dwAvailVirtual              As Long
End Type

Private Type PROCESS_MEMORY_COUNTERS
    cb                          As Long
    PageFaultCount              As Long
    PeakWorkingSetSize          As Long
    WorkingSetSize              As Long
    QuotaPeakPagedPoolUsage     As Long
    QuotaPagedPoolUsage         As Long
    QuotaPeakNonPagedPoolUsage  As Long
    QuotaNonPagedPoolUsage      As Long
    PagefileUsage               As Long
    PeakPagefileUsage           As Long
End Type

Private Type PROCESSENTRY32
    dwSize                      As Long
    cntUsage                    As Long
    th32ProcessID               As Long         ' This process
    th32DefaultHeapID           As Long
    th32ModuleID                As Long         ' Associated exe
    cntThreads                  As Long
    th32ParentProcessID         As Long         ' This process's parent process
    pcPriClassBase              As Long         ' Base priority of process threads
    dwFlags                     As Long
    szExeFile                   As String * 260 ' MAX_PATH
End Type

Public Type ProcessInfoType
    lngPID                      As Long         'holds the process id
    strProcName                 As String       'holds the name of the process
    pmcProcMemInfo              As PROCESS_MEMORY_COUNTERS
End Type

Private Type LOGFONT
    lfHeight                    As Long
    lfEscapement                As Long
    lfUnderline                 As Byte
    lfStrikeOut                 As Byte
    lfWidth                     As Long
    lfWeight                    As Long
    lfItalic                    As Byte
    lfCharSet                   As Byte
    lfClipPrecision             As Byte
    lfOutPrecision              As Byte
    lfQuality                   As Byte
    lfPitchAndFamily            As Byte
    lfOrientation               As Byte
    lfFaceName(32)              As Byte
End Type

Private Type NONCLIENTMETRICS
    cbSize                      As Long
    iBorderWidth                As Long
    iScrollWidth                As Long
    iScrollHeight               As Long
    iCaptionWidth               As Long
    iCaptionHeight              As Long
    lfCaptionFont               As LOGFONT
    iSMCaptionWidth             As Long
    iSMCaptionHeight            As Long
    lfSMCaptionFont             As LOGFONT
    iMenuWidth                  As Long
    iMenuHeight                 As Long
    lfMenuFont                  As LOGFONT
    lfStatusFont                As LOGFONT
    lfMessageFont               As LOGFONT
End Type


'------------------------------------------------
'                   ENUMERATORS
'------------------------------------------------
'base thread priorities
Public Enum EnumBaseThread
    THREAD_BASE_PRIORITY_IDLE = -15
    THREAD_BASE_PRIORITY_LOWRT = 15
    THREAD_BASE_PRIORITY_MIN = -2
    THREAD_BASE_PRIORITY_MAX = 2
End Enum

'thread priorities
Public Enum EnumThreadPriority
    THREAD_PRIORITY_LOWEST = THREAD_BASE_PRIORITY_MIN
    THREAD_PRIORITY_HIGHEST = THREAD_BASE_PRIORITY_MAX
    THREAD_PRIORITY_BELOW_NORMAL = (THREAD_PRIORITY_LOWEST + 1)
    THREAD_PRIORITY_ABOVE_NORMAL = (THREAD_PRIORITY_HIGHEST - 1)
    THREAD_PRIORITY_IDLE = THREAD_BASE_PRIORITY_IDLE
    THREAD_PRIORITY_NORMAL = 0
    THREAD_PRIORITY_TIME_CRITICAL = THREAD_BASE_PRIORITY_LOWRT
End Enum

'thread class priorities
Public Enum EnumClassPriority
    HIGH_PRIORITY_CLASS = &H80
    IDLE_PRIORITY_CLASS = &H40
    NORMAL_PRIORITY_CLASS = &H20
    REALTIME_PRIORITY_CLASS = &H100
End Enum

'font details
Public Enum EnumSysFonts
    FNT_CAPTION = 0
    FNT_MENU = 1
    FNT_MESSAGE = 2
    FNT_STATUS = 3
End Enum


'------------------------------------------------
'               MODULE-LEVEL CONSTANTS
'------------------------------------------------
'system menu constants
Private Const MF_BYPOSITION             As Long = &H400&
Private Const MF_APPEND                 As Long = &H100&
Private Const MF_BITMAP                 As Long = &H4&
Private Const MF_BYCOMMAND              As Long = &H0&
Private Const MF_CALLBACKS              As Long = &H8000000
Private Const MF_CHANGE                 As Long = &H80&
Private Const MF_CHECKED                As Long = &H8&
Private Const MF_CONV                   As Long = &H40000000
Private Const MF_DELETE                 As Long = &H200&
Private Const MF_DISABLED               As Long = &H2&
Private Const MF_ENABLED                As Long = &H0&
Private Const MF_END                    As Long = &H80
Private Const MF_ERRORS                 As Long = &H10000000
Private Const MF_GRAYED                 As Long = &H1&
Private Const MF_HELP                   As Long = &H4000&
Private Const MF_HILITE                 As Long = &H80&
Private Const MF_HSZ_INFO               As Long = &H1000000
Private Const MF_INSERT                 As Long = &H0&
Private Const MF_LINKS                  As Long = &H20000000
Private Const MF_MASK                   As Long = &HFF000000
Private Const MF_MENUBARBREAK           As Long = &H20&
Private Const MF_MENUBREAK              As Long = &H40&
Private Const MF_MOUSESELECT            As Long = &H8000&
Private Const MF_OWNERDRAW              As Long = &H100&
Private Const MF_POPUP                  As Long = &H10&
Private Const MF_POSTMSGS               As Long = &H4000000
Private Const MF_REMOVE                 As Long = &H1000&
Private Const MF_SENDMSGS               As Long = &H2000000
Private Const MF_SEPARATOR              As Long = &H800&
Private Const MF_STRING                 As Long = &H0&
Private Const MF_SYSMENU                As Long = &H2000&
Private Const MF_UNCHECKED              As Long = &H0&
Private Const MF_UNHILITE               As Long = &H0&
Private Const MF_USECHECKBITMAPS        As Long = &H200&
Private Const MFCOMMENT                 As Long = 15
Private Const MH_CLEANUP                As Long = 4
Private Const MH_CREATE                 As Long = 1

'used by GetWindow to return the specified window
Private Const GW_HWNDFIRST              As Long = 0
Private Const GW_HWNDLAST               As Long = 1
Private Const GW_HWNDNEXT               As Long = 2
Private Const GW_HWNDPREV               As Long = 3
Private Const GW_OWNER                  As Long = 4
Private Const GW_CHILD                  As Long = 5
Private Const GW_MAX                    As Long = 5

'window layer constants
Private Const GWL_EXSTYLE               As Long = (-20)
Private Const LWA_COLORKEY              As Long = &H1
Private Const LWA_ALPHA                 As Long = &H2
Private Const ULW_COLORKEY              As Long = &H1
Private Const ULW_ALPHA                 As Long = &H2
Private Const ULW_OPAQUE                As Long = &H4
Private Const WS_EX_LAYERED             As Long = &H80000

'process constants
Private Const Default_Log_Size          As Long = 10000000
Private Const Default_Log_Days          As Integer = 0
Private Const hNull                     As Integer = 0
Private Const MAX_PATH                  As Integer = 260
Private Const PROCESS_ALL_ACCESS        As Long = &H1F0FFF
Private Const PROCESS_QUERY_INFORMATION As Integer = 1024
Private Const PROCESS_VM_READ           As Integer = 16
Private Const SPECIFIC_RIGHTS_ALL       As Long = &HFFFF
Private Const STANDARD_RIGHTS_ALL       As Long = &H1F0000
Private Const SYNCHRONIZE               As Long = &H100000
Private Const TH32CS_SNAPPROCESS        As Long = &H2&
Private Const WIN95_System_Found        As Integer = 1
Private Const WINNT_System_Found        As Integer = 2

'these are used to specify the hotkeys
Private Const MOD_ALT                   As Long = &H1
Private Const MOD_CONTROL               As Long = &H2
Private Const MOD_SHIFT                 As Long = &H4

Private Const WM_LBUTTONDBLCLICK        As Long = &H203
Private Const WM_MOUSEMOVE              As Long = &H200
Private Const WM_RBUTTONUP              As Long = &H205
Private Const WM_LBUTTONDOWN            As Long = &H201
Private Const WM_LBUTTONUP              As Long = &H202
Private Const WM_LBUTTONDBLCLK          As Long = &H203
Private Const WM_RBUTTONDOWN            As Long = &H204
Private Const WM_RBUTTONDBLCLK          As Long = &H206
Private Const WM_CHAR                   As Long = &H102
Private Const WM_CLOSE                  As Long = &H10
Private Const WM_USER                   As Long = &H400
Private Const WM_COMMAND                As Long = &H111
Private Const WM_GETTEXT                As Long = &HD
Private Const WM_GETTEXTLENGTH          As Long = &HE
Private Const WM_KEYDOWN                As Long = &H100
Private Const WM_KEYUP                  As Long = &H101
Private Const WM_MOVE                   As Long = &HF012
Private Const WM_SETTEXT                As Long = &HC
Private Const WM_CLEAR                  As Long = &H303
Private Const WM_DESTROY                As Long = &H2
Private Const WM_SYSCOMMAND             As Long = &H112

Private Const SWP_NOSIZE                As Long = &H1
Private Const SWP_NOMOVE                As Long = &H2
Private Const SW_MINIMIZE               As Long = 6
Private Const SW_HIDE                   As Long = 0
Private Const SW_MAXIMIZE               As Long = 3
Private Const SW_SHOW                   As Long = 5
Private Const SW_RESTORE                As Long = 9
Private Const SW_SHOWDEFAULT            As Long = 10
Private Const SW_SHOWMAXIMIZED          As Long = 3
Private Const SW_SHOWMINIMIZED          As Long = 2
Private Const SW_SHOWMINNOACTIVE        As Long = 7
Private Const SW_SHOWNOACTIVATE         As Long = 4
Private Const SW_SHOWNORMAL             As Long = 1

Private Const HWND_TOP                  As Long = 0
Private Const HWND_TOPMOST              As Long = -1
Private Const HWND_NOTOPMOST            As Long = -2

Private Const EWX_LOGOFF                As Long = 0
Private Const EWX_SHUTDOWN              As Long = 1
Private Const EWX_REBOOT                As Long = 2
Private Const EWX_FORCE                 As Long = 4
Private Const EWX_POWEROFF              As Long = 8

Private Const RSP_SIMPLE_SERVICE        As Long = 1
Private Const RSP_UNREGISTER_SERVICE    As Long = 0

Private Const SPI_SCREENSAVERRUNNING    As Long = 97
Private Const STANDARD_RIGHTS_REQUIRED  As Long = &HF0000

Private Const FLAGS                     As Long = SWP_NOSIZE + SWP_NOMOVE

Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const TOKEN_QUERY               As Long = &H8
Private Const SE_PRIVILEGE_ENABLED      As Long = &H2
Private Const ANYSIZE_ARRAY             As Long = 1
Private Const VER_PLATFORM_WIN32_NT     As Long = 2
Private Const SPI_GETNONCLIENTMETRICS   As Long = 41
Private Const LF_FACESIZE               As Long = 32


'------------------------------------------------
'                   PROCEDURES
'------------------------------------------------
Public Sub StayOnTop(ByVal frm As Form)
    'This will put the specified window on top
    'on the z order pile
    
    Dim lngSetWinOnTop As Long
    
    lngSetWinOnTop = SetWindowPos(frm.hWnd, _
                                  HWND_TOPMOST, _
                                  0, _
                                  0, _
                                  0, _
                                  0, _
                                  FLAGS)
End Sub

Public Sub NotOnTop(ByVal frm As Form)
    'This will put the specified window to be
    'handeled in the z order as normal
    
    Dim lngSetWinOnTop As Long
    
    lngSetWinOnTop = SetWindowPos(frm.hWnd, _
                                  HWND_NOTOPMOST, _
                                  0, _
                                  0, _
                                  0, _
                                  0, _
                                  FLAGS)
End Sub

Public Sub HideTaskBar()
    'This will make the task bar invisible to the
    'user. The Start menu is still accessable through
    'the Start key or by pressing Ctrl+Esc
    
    Dim lngHandle As Long
    
    lngHandle = FindWindow("Shell_TrayWnd", _
                           vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowTaskBar()
    'See the procedure HideTaskBar
    
    Dim lngHandle As Long
    
    lngHandle = FindWindow("Shell_TrayWnd", _
                           vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub HideStartButton()
    'This will hide the start button from the user
    
    Dim lngHandle       As Long
    Dim lngFindClass    As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "Button", _
                             vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowStartButton()
    'This will make sure that the start button
    'is visible to the user
    
    Dim lngHandle       As Long
    Dim lngFindClass    As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "Button", _
                             vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub DestroyStartButton()
    'This will remove the start bottom from the current
    'windows session. The user will not be able to
    'access the start menu.
    
    Dim lngHandle       As Long
    Dim lngFindClass    As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "Button", _
                             vbNullString)
    SendMessage lngHandle, WM_DESTROY, 0, 0
End Sub

Public Sub HideTaskBarClock()
    'This will hide the clock from the user
    
    Dim lngFindClass    As Long
    Dim lngFindParent   As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", _
                              vbNullString)
    lngFindParent = FindWindowEx(lngFindClass, _
                                 0, _
                                 "TrayNotifyWnd", _
                                 vbNullString)
    lngHandle = FindWindowEx(lngFindParent, _
                             0, _
                             "TrayClockWClass", _
                             vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowTaskBarClock()
    'This will show the task bar clock to the user
    
    Dim lngFindClass    As Long
    Dim lngFindParent   As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", _
                              vbNullString)
    lngFindParent = FindWindowEx(lngFindClass, _
                                 0, _
                                 "TrayNotifyWnd", _
                                 vbNullString)
    lngHandle = FindWindowEx(lngFindParent, _
                             0, _
                             "TrayClockWClass", _
                             vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub DestroyTaskBarClock()
    'This will remove the clock from the current
    'windows session
    
    Dim lngFindClass    As Long
    Dim lngFindParent   As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", _
                              vbNullString)
    lngFindParent = FindWindowEx(lngFindClass, _
                                 0, _
                                 "TrayNotifyWnd", _
                                 vbNullString)
    lngHandle = FindWindowEx(lngFindParent, _
                             0, _
                             "TrayClockWClass", _
                             vbNullString)
    SendMessage lngHandle, WM_DESTROY, 0, 0
End Sub

Public Sub HideTaskBarIcons()
    'This will hide any tool bars on the task bar.
    'This only applies to versions of window > 98
    
    Dim lngFindClass    As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "TrayNotifyWnd", _
                             vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowTaskBarIcons()
    'This will show any hidden active toolbars
    'on the task bar. See HideTaskbarIcons for
    'more details
    
    Dim lngFindClass    As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "TrayNotifyWnd", _
                             vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub DestroyTaskBarIcons()
    'This will remove any tool bars currently active
    'on the task bar.
    
    Dim lngFindClass    As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngHandle = FindWindowEx(lngFindClass, _
                             0, _
                             "TrayNotifyWnd", _
                             vbNullString)
    SendMessage lngHandle, WM_DESTROY, 0, 0
End Sub

Public Sub HideProgramsShowingInTaskBar()
    'This will hide any program icon showing in the
    'task bar.
    
    Dim lngFindClass    As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngFindClass2 = FindWindowEx(lngFindClass, _
                                 0, _
                                 "ReBarWindow32", _
                                 vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                             0, _
                             "MSTaskSwWClass", _
                             vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                             0, _
                             "SysTabControl32", _
                             vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowProgramsShowingInTaskBar()
    'This will show any program icons in the taskbar
    'that were hidden. The programs must be active
    
    Dim lngFindClass    As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngFindClass2 = FindWindowEx(lngFindClass, _
                                 0, _
                                 "ReBarWindow32", _
                                 vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                             0, _
                             "MSTaskSwWClass", _
                             vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                             0, _
                             "SysTabControl32", _
                             vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub DestroyProgramsShowingInTaskBar()
    'This will permenently remove any currently showing
    'program icons in the task bar
    
    Dim lngFindClass    As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass = FindWindow("Shell_TrayWnd", "")
    lngFindClass2 = FindWindowEx(lngFindClass, _
                                 0, _
                                 "ReBarWindow32", _
                                 vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                             0, _
                             "MSTaskSwWClass", _
                             vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                             0, _
                             "SysTabControl32", _
                             vbNullString)
    SendMessage lngHandle, WM_DESTROY, 0, 0
End Sub

Public Sub HideWindowsToolBar()
    'This will hide the windows tool bar from the user
    
    Dim lngFindClass1   As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass1 = FindWindow("BaseBar", vbNullString)
    lngFindClass2 = FindWindowEx(lngFindClass1, _
                                 0, _
                                 "ReBarWindow32", _
                                 vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                             0, _
                             "SysPager", _
                             vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                             0, _
                             "ToolbarWindow32", _
                             vbNullString)
    ShowWindow lngHandle, 0
End Sub

Public Sub ShowWindowsToolBar()
    'This will show the currently running windows
    'tool bar
    
    Dim lngFindClass1   As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass1 = FindWindow("BaseBar", vbNullString)
    lngFindClass2 = FindWindowEx(lngFindClass1, _
                                 0, _
                                 "ReBarWindow32", _
                                 vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                             0, _
                             "SysPager", _
                             vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                             0, _
                             "ToolbarWindow32", _
                             vbNullString)
    ShowWindow lngHandle, 1
End Sub

Public Sub DestroyWindowsToolBar()
    'This will remove the windows tool bar from
    'the current windows session
    
    Dim lngFindClass1   As Long
    Dim lngFindClass2   As Long
    Dim lngParent       As Long
    Dim lngHandle       As Long
    
    lngFindClass1 = FindWindow("BaseBar", vbNullString)
    lngFindClass2 = FindWindowEx(lngFindClass1, _
                                  0, _
                                  "ReBarWindow32", _
                                  vbNullString)
    lngParent = FindWindowEx(lngFindClass2, _
                              0, _
                              "SysPager", _
                              vbNullString)
    lngHandle = FindWindowEx(lngParent, _
                              0, _
                              "ToolbarWindow32", _
                              vbNullString)
    SendMessage lngHandle, WM_DESTROY, 0, 0
End Sub

Public Sub ScreenBlackOut(ByVal TheForm As Form)
    'This will black out the screen by forcing the
    'current window to the top of the zorder and
    'disabling the appropiate counter measures
    
    'disbale the counter measures and set the
    'form details to cover the entire screen
    'preventing the user from accessing the desktop
    'or any other programs
    Call StayOnTop(TheForm)
    Call HideTaskBar
    Call HideWindowsToolBar
    Call PreventFromClosing
    
    'set the forms appearance to conver the entire
    'screen
    With TheForm
        .BorderStyle = 0
        .Caption = ""
        .BackColor = &H0&
        .Height = Screen.Height
        .Width = Screen.Width
        .Left = Screen.Width - Screen.Width
        .Top = Screen.Height - Screen.Height
    End With
    
    Screen.MousePointer = vbHourglass
End Sub

Public Sub ScreenUnBlackOut(ByVal TheForm As Form)
    'This will undo what was done by calling the
    'ScreenBlackOut procedure
    
    'return normal window management
    Call NotOnTop(TheForm)
    Call ShowTaskBar
    Call ShowWindowsToolBar
    Call UnPreventFromClosing
    
    'set the forms details back to what they were
    With TheForm
        TheForm.BorderStyle = 3
        TheForm.Caption = "Form"
        TheForm.BackColor = &H8000000A
        TheForm.Width = Screen.Width / 2
        TheForm.Height = Screen.Height / 2
        TheForm.Left = Screen.Width / 2 - TheForm.Width / 2
        TheForm.Top = Screen.Height / 2 - TheForm.Height / 2
    End With
    
    Screen.MousePointer = vbArrow
End Sub

Public Sub PreventFromClosing()
    'This will stop the program from being closed.
    'Also see the VB help on the QueryUnLoad event
    
    Dim lngProcessId    As Long
    Dim lngResult       As Long
    
    lngProcessId = GetCurrentProcessId()
    
    lngResult = RegisterServiceProcess(lngProcessId, _
                                       RSP_SIMPLE_SERVICE)
End Sub

Public Sub UnPreventFromClosing()
    'See the PreventFromClosing procedure
    
    Dim lngProcessId    As Long
    Dim lngResult       As Long
    
    lngProcessId = GetCurrentProcessId()
    lngResult = RegisterServiceProcess(lngProcessId, _
                                       RSP_UNREGISTER_SERVICE)
End Sub

Public Sub WINLogUserOff()
    'Log off the user from the current windows session
    
    Dim lngResult As Long
    
    Call EnableShutDown
    lngResult = ExitWindowsEx(EWX_LOGOFF, 1)
End Sub

Public Sub WINForceClose()
    'Immediatly close all programs without sending
    'a message to all applicable windows that windows
    'is logging off. Please note that any unsaved
    'imformation will be lost
    
    Dim lngResult As Long
    
    Call EnableShutDown
    lngResult = ExitWindowsEx(EWX_FORCE Or _
                              EWX_SHUTDOWN, _
                              1)
End Sub

Public Sub WINShutdown()
    'Shut down the computer. Does not work on NT
    'based machines
    
    Dim lngResult As Long
    
    Call EnableShutDown
    lngResult = ExitWindowsEx(EWX_SHUTDOWN, 1)
End Sub

Public Sub WINReboot()
    'Reboot the machine. Does not work on NT based
    'machines
    
    Dim lngResult As Long
    
    Call EnableShutDown
    lngResult = ExitWindowsEx(EWX_REBOOT, 1)
End Sub

Public Sub WINPowerDown()
    'This will power down the computer.
    
    Dim lngResult As Long
    
    Call EnableShutDown
    lngResult = ExitWindowsEx(EWX_POWEROFF Or _
                              EWX_SHUTDOWN Or _
                              EWX_FORCE, _
                              1)
End Sub

Public Sub WINLock()
    'This will lock the workstation if the os is win 2000 or greater
    
    Dim lngResult   As Long     'holds any returned value from an api call
    
    If IsW2000 Then
        lngResult = LockWorkStation
    End If
End Sub

Private Sub EnableShutDown()
    'set the shut down privilege for the current
    'application
    
    Dim hProc       As Long
    Dim hToken      As Long
    Dim mLUID       As LUID
    Dim mPriv       As TOKEN_PRIVILEGES
    Dim mNewPriv    As TOKEN_PRIVILEGES
    Dim lngResult   As Long   'holds any returned error value from the api call
    
    'only enable shutdown if this is an NT based system
    If Not IsWinNT Then
        Exit Sub
    End If
    
    'This is a winNT based system - adjust the
    'privilages so that we are allowed to shut down the
    'current windows session
    
    'get the privilages for this app
    hProc = GetCurrentProcess()
    lngResult = OpenProcessToken(hProc, _
                                 TOKEN_ADJUST_PRIVILEGES + TOKEN_QUERY, _
                                 hToken)
    lngResult = LookupPrivilegeValue("", _
                                     "SeShutdownPrivilege", _
                                     mLUID)
    
    'adjust this applications privilages for shut down
    With mPriv
        .PrivilegeCount = 1
        .Privileges(0).Attributes = SE_PRIVILEGE_ENABLED
        .Privileges(0).pLuid = mLUID
    End With
    
    'enable shutdown privilege for the current
    'application
    lngResult = AdjustTokenPrivileges(hToken, _
                                      False, _
                                      mPriv, _
                                      4 + (12 * mPriv.PrivilegeCount), _
                                      mNewPriv, _
                                      4 + (12 * mNewPriv.PrivilegeCount))
End Sub

Public Sub DestroyShell()
    'This will stop the explorer.exe process from running and if this is a
    'win NT based system, it will also temperorily disable the auto restart
    'function
    '
    'NOTE: This procedure requires the Registry module for "AutoRestartShell"
    
    Call AutoRestartShell(False)
    Call KillProcess("explorer.exe")
End Sub

Public Sub DestroyTaskBar()
    'This will close the processing thread that
    'manages the taskbar. This is basically closing
    'the program that is the task bar.
    
    Dim lngHandle   As Long     'holds a handle to the taskbar "window"
    Dim lngResult   As Long     'holds the returned error value from the api call
    
    lngHandle = FindWindow("Shell_TrayWnd", _
                        vbNullString)
    If lngHandle <> 0 Then
        lngResult = SendMessage(lngHandle, WM_DESTROY, 0, 0)
    End If
End Sub

Public Sub DestroyDesktop()
    'This will close the processing thread that
    'manages the taskbar. This is basically closing
    'the program that is the task bar.
    
    Dim lngHandle   As Long         'holds a handle to the taskbar "window"
    Dim lngResult   As Long         'holds the returned error value from the api call
    
    lngHandle = GetWindowHandle("Program Manager")
    
    If lngHandle <> 0 Then
        lngResult = SendMessage(lngHandle, WM_DESTROY, 0, 0)
    End If
End Sub

Public Function GetWindowHandle(ByVal strCaptionText As String) _
                                As Long
    'This will return a handle to the window found with matching text (if any
    'valid window is found)
    
    Dim lngResult       As Long         'holds any returned error value from an api call
    Dim hWnd            As Long         'holds a handle to the current window
    Dim strWindowText   As String * 128 'holds the window caption
    Dim hWndDesktop     As Long         'holds a handle to the desktop window
    
    'get a handle to the first window
    hWndDesktop = GetDesktopWindow
    hWnd = GetWindow(hWndDesktop, GW_CHILD)
    'hWnd = GetWindow(hWndDesktop, GW_HWNDFIRST)
    Do While hWnd <> 0
        'check the text
        lngResult = GetWindowText(hWnd, strWindowText, 127)
        
        'Debug.Print strWindowText
        'does the caption contain the string we are looking for
        If (InStr(1, strWindowText, strCaptionText, vbTextCompare) > 0) Then
            'window found, exit loop still holding a handle to the window
            Exit Do
        End If
        
        'get a handle to the next window to check
        strWindowText = String(Len(strWindowText), vbNullChar)
        hWnd = GetWindow(hWnd, GW_HWNDNEXT)
    Loop
    
    'return the handle of the window
    GetWindowHandle = hWnd
End Function

Public Sub CreateShell()
    'This will create the task bar if it has been destroyed
    '
    'NOTE: This procedure requires the registry module for
    '      "AutoRestartShell" and "GetWinDirectories"
    
    Dim strWinPath      As String       'holds the path to the windows directory
    Dim strExplorerPath As String       'holds the complete file path to Exploere.exe
    Dim lngInstanceID   As Long         'holds the instance id of explorer when started
    
    'if this is windows NT, make sure that the shell will start automatically
    'should it unexpectadly shutdown
    Call AutoRestartShell(True)
    
    'get the direcotry path for the windows directory
    strWinPath = GetWinDirectories(WindowsDir)
    
    'get the path to "explorer.exe" and run it to create the task bar
    strExplorerPath = AddToPath(strWinPath, "Explorer.exe")
    lngInstanceID = Shell(strExplorerPath)
End Sub

Public Function AddToPath(ByVal strDirectory As String, _
                          ByVal strFileName As String) _
                          As String
    'This will add a file name or a directory to an
    'existing directory path to create a full filepath.
    'This is particuarly usefull when the application
    'returns just a drive, eg, "C:\"
    
    If Right(strDirectory, 1) <> "\" Then
        'insert a backslash
        strDirectory = strDirectory + "\"
    End If
    
    'append the file name to the directory path
    AddToPath = strDirectory + strFileName
End Function

Public Sub CtrlAltDel(ByVal blnEnable As Boolean)
    'This will prevent the task manager from appearing
    'when the user presses Ctrl+Alt+Del. This will also disable the
    'Alt+Tab and the Ctrl+Esc key combinations. Only works on 95/98
    
    Dim lngResult       As Integer      'holds any returned error value from an api call
    Dim blnOld          As Boolean      'holds whether or not the screensaver was already running
    
    blnOld = blnEnable
    lngResult = SystemParametersInfo(SPI_SCREENSAVERRUNNING, _
                                     blnEnable, _
                                     blnOld, _
                                     0)
End Sub

Public Sub AltTab(ByVal blnEnable As Boolean, _
                  Optional ByVal hWndCapture As Long)
    'This will enable or disable the Alt+Tab functionality for windows. The
    'hWnd parameter is needed, because Alt+Tab must be re-directed to a window
    'instead of the operating system. The parameter is also needed to remove the
    'functionality.
    
    Const HOT_KEY       As Long = 9 'holds the numeric/ascii value for the hotkey
    
    'If the code is compiled to an APP, this needs to be between 0 and 49151
    'If the code is compiled to a DLL, this nees to be between 49152 and 65535
    Static intHotkeyId  As Integer  'holds the windows id for the hotkey
    Static hWnd         As Long     'holds a handle to the desktop window
    
    Dim lngResult       As Long     'holds any returned error value from an api call
    Dim blnOld          As Boolean  'holds whether or not the screensaver was already running
    
    'turn on/off Alt+Tab
    If blnEnable Then
        'turn on the Alt+Tab functionality for windows
        If IsWinNT Then
            If intHotkeyId <> 0 Then
                lngResult = UnregisterHotKey(hWnd, intHotkeyId)
                If lngResult <> False Then
                    hWnd = 0
                    intHotkeyId = 0
                End If
            End If
        
        Else
            'on non-NT based systems, we can fool the computer into disabling
            'Alt+Tab by telling it a screen saver is running
            lngResult = SystemParametersInfo(SPI_SCREENSAVERRUNNING, True, blnOld, 0)
        End If
    Else
        'disable the Alt+Tab functionality for windows
        
        If IsWinNT Then
            'we cannot register a hotkey unless a hand to a window was specified
            If hWndCapture = 0 Then
                Exit Sub
            End If
            
            'remove any active hotkey
            If intHotkeyId <> 0 Then
                'hotkey already active - try to disable
                lngResult = UnregisterHotKey(hWnd, intHotkeyId)
                If lngResult <> False Then
                    hWnd = 0
                    intHotkeyId = 0
                Else
                    'unable to remove hotkey or invalid hWnd
                    Exit Sub
                End If
            End If
            
            'get a handle to the desktop window
            'hWnd = GetDesktopWindow
            hWnd = hWndCapture
            
            'register the hotkey to disable Alt+Tab
            lngResult = RegisterHotKey(hWnd, intHotkeyId, MOD_ALT, HOT_KEY)
        
        Else
            'on non-NT based systems, we can fool the computer into disabling
            'Alt+Tab by telling it a screen saver is running
            lngResult = SystemParametersInfo(SPI_SCREENSAVERRUNNING, False, blnOld, 0)
        End If  'is the version of windows NT based
    End If
End Sub

Public Function IsWinNT() As Boolean
    'Detect if the program is running under Windows NT
    
    Const VER_PLATFORM_WIN32_NT     As Long = 2
    
    Dim udtOsInfo   As OSVERSIONINFO    'holds the operating system information
    Dim lngResult   As Long             'returned error value from the api call
    
    'get version information
    udtOsInfo.dwOSVersionInfoSize = Len(udtOsInfo)
    lngResult = GetVersionEx(udtOsInfo)
    
    'return True if the test of windows NT is positive
    IsWinNT = (udtOsInfo.dwPlatformId = VER_PLATFORM_WIN32_NT)
End Function

Public Function IsWndTransparent(ByVal hWnd As Long) As Boolean
    'This will return True if the window is already transparent
    
    Dim lngResult       As Long         'holds any returned value from an api call
    
    On Error GoTo Err_Trap
    
    'make sure that a window handle was passed
    If (hWnd = 0) Then
        IsWndTransparent = False
        Exit Function
    End If
    
    'get the state of the layered attribute
    lngResult = GetWindowLong(hWnd, WS_EX_LAYERED)
    
    'is the window layered
    IsWndTransparent = ((lngResult And WS_EX_LAYERED) = WS_EX_LAYERED)  'returns True or False
    
    Exit Function
Err_Trap:
    'the api call does not exist on this os
    IsWndTransparent = False
End Function

Public Sub MakeWndTransparent(ByVal hWnd As Long, _
                              Optional ByVal sngPercTrans As Single = 0)
    'This will make the specified window transparent to the specified value (between
    '0 and 1, where 0 is fully opaque and 100 is fully transparent)
    
    Dim intPercByte     As Integer      'holds a value between 0 and 255 that relates to the opacity of the window
    Dim lngMsg          As Long         'holds the window attributes to set to the specified window
    Dim lngResult       As Long         'holds any returned value from an api call
    
    On Error GoTo Err_Trap
    
    'validate the parameters
    If (hWnd = 0) Then
        Exit Sub
    End If
    
    'make sure that the transparent value is not outside valid ranges
    Select Case sngPercTrans
    Case Is > 100
        sngPercTrans = 100
    Case Is < 0
        sngPercTrans = 0
    End Select
    
    'convert the percentage to a value between 0 and 255 where 0 is opaque
    'and 255 is fully transparent
    intPercByte = (100 - sngPercTrans) * 2.55
    
    'get the current window transparent status
    lngMsg = GetWindowLong(hWnd, GWL_EXSTYLE)
    
    'make the window transparent to the specified amount
    lngMsg = lngMsg Or WS_EX_LAYERED
    lngResult = SetWindowLong(hWnd, GWL_EXSTYLE, lngMsg)
    lngResult = SetLayeredWindowAttributes(hWnd, 0, intPercByte, LWA_ALPHA)
    
    Exit Sub
Err_Trap:
    'ignore any errors caused by "Invalid dll entry point" which would be generated
    'on any operating system not greater than Win2000 or WinXP
End Sub

Public Sub LockDown(ByVal blnLockComputer As Boolean, _
                    ByVal hWndMainForm As Long)
    'This will lock the computer if flagged True and unlock the computer
    'if flagged False
    '
    'NOTE: This procedure requires the registry module for "NTMenus"
    '      and "AutoRestartShell"
    
    Dim blnEnable   As Boolean      'enable or disable specific features
    
    'if we WANT to lock down the computer (True), then we must DISable
    'certain features (False), and vice versa
    blnEnable = Not blnLockComputer
    
    'enable/disable windows features
    Call AutoRestartShell(blnEnable)
    Call AltTab(blnEnable, hWndMainForm)
    Call CtrlAltDel(blnEnable)          'only works on 95/98
    
    'set windows NT/2000/XP settings
    Call NTMenus(CHANGE_PASSWORD, blnEnable)
    Call NTMenus(LOCK_WORKSTATION, blnEnable)
    Call NTMenus(TASK_MGR, blnEnable)
    
    'hide the program from the end task list
    App.TaskVisible = blnEnable
    
    'create/destroy the windows shell
    If blnEnable Then
        'restart the explorer shell
        Call CreateShell
    Else
        'stop all threads and instances of the shell from running
        Call DestroyShell
    End If
End Sub

Public Sub GetProcesses(ByVal ExeName As String)

    Dim booResult               As Boolean
    Dim lngLength               As Long
    Dim strProcessName          As String
    Dim lngCbSize               As Long 'Specifies the size, In bytes, of the lpidProcess array
    Dim lngCbSizeReturned       As Long 'Receives the number of bytes returned
    Dim lngNumElements          As Long
    Dim lngProcessIds()         As Long
    Dim lngCbSize2              As Long
    Dim lngModules(1 To 200)    As Long
    Dim lngReturn               As Long
    Dim strModuleName           As String
    Dim lngSize                 As Long
    Dim lngHwndProcess          As Long
    Dim lngLoop                 As Long
    Dim Pmc                     As PROCESS_MEMORY_COUNTERS
    Dim lRet                    As Long
    Dim strProcName            As String

    'Turn on Error handler
    On Error GoTo Error_Handler

    booResult = False

    ExeName = UCase$(Trim$(ExeName))
    lngLength = Len(ExeName)

    'ProcessInfo.bolRunning = False

    Select Case GetVersion()
        'I'm not bothered about windows 95/98 becasue this class probably wont be used on it anyway.
        Case WIN95_System_Found 'Windows 95/98

        Case WINNT_System_Found 'Windows NT

            lngCbSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
            lngCbSizeReturned = 96

            Do While lngCbSize <= lngCbSizeReturned
                DoEvents
                'Increment Size
                lngCbSize = lngCbSize * 2
                'Allocate Memory for Array
                ReDim lngProcessIds(lngCbSize / 4)
                'Get Process ID's
                lngReturn = EnumProcesses(lngProcessIds(1), lngCbSize, lngCbSizeReturned)
            Loop

            'Count number of processes returned
            lngNumElements = lngCbSizeReturned / 4
            'Loop thru each process

            For lngLoop = 1 To lngNumElements
                DoEvents
    
                'Get a handle to the Process and Open
                lngHwndProcess = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, _
                                             0, _
                                             lngProcessIds(lngLoop))
    
                If lngHwndProcess <> 0 Then
                    'Get an array of the module handles for the specified process
                    lngReturn = EnumProcessModules(lngHwndProcess, lngModules(1), 200, lngCbSize2)
    
                    'If the Module Array is retrieved, Get the ModuleFileName
                    If lngReturn <> 0 Then
    
                        'Buffer with spaces first to allocate memory for byte array
                        strModuleName = Space(MAX_PATH)
    
                        'Must be set prior to calling API
                        lngSize = 500
    
                        'Get Process Name
                        lngReturn = GetModuleFileNameExA(lngHwndProcess, _
                                                         lngModules(1), _
                                                         strModuleName, _
                                                         lngSize)
    
                        'Remove trailing spaces
                        strProcessName = Left(strModuleName, lngReturn)
    
                        'gstrCheck for Matching Upper case result
                        strProcessName = UCase$(Trim$(strProcessName))
    
                        strProcName = GetElement(Trim(Replace(strProcessName, _
                                                              vbNullChar, _
                                                              "")), _
                                                 "\", _
                                                 0, _
                                                 0, _
                                                 GetNumElements(Trim(Replace(strProcessName, _
                                                                             vbNullChar, _
                                                                             "")), _
                                                                "\") - 1)
    
                        If strProcName = ExeName Then
    
                            'Get the Site of the Memory Structure
                            Pmc.cb = LenB(Pmc)
    
                            lRet = GetProcessMemoryInfo(lngHwndProcess, Pmc, Pmc.cb)
    
                             Debug.Print ExeName & "::" & CStr(Pmc.WorkingSetSize / 1024)
                        End If
                    End If
                End If
                'Close the handle to this process
                lngReturn = CloseHandle(lngHwndProcess)
                DoEvents
            Next
    End Select

IsProcessRunning_Exit:

'Exit early to avoid error handler
Exit Sub
Error_Handler:
    Err.Raise Err, Err.Source, "ProcessInfo", Error
    Resume Next
End Sub

Public Function GetProcessList(ByRef intNumFound As Integer, _
                               Optional ByVal strPattern As String = "*") _
                               As ProcessInfoType()
    'This function will return a list of active processes. If a patter string is specified in the parameters
    'then only programs matching this string will be returned. The string values returned are the complete
    'file and path names to the processes. The pattern matching is NOT case sensitive.
    
    Dim prcList()               As ProcessInfoType          'holds the list of active processes currently found
    Dim intNumProc              As Integer                  'holds the number of active processes in the prcList() array
    Dim lngResult               As Long                     'holds any returned value from an api call
    Dim lngPIDs()               As Long                     'holds a list of all the process IDs
    Dim udtPMC                  As PROCESS_MEMORY_COUNTERS  'holds process information returned from an api call
    Dim intNumDel               As Integer                  'holds the number of processes that have been removed from the array when checking against the specified parameter
    Dim lngCbSize               As Long
    Dim lngCbSize2              As Long
    Dim lngCbSizeRet            As Long
    Dim lngSize                 As Long                     'holds the size of the buffer to contain the process name
    Dim strProcName             As String                   'holds the name of the process
    Dim intIndex                As Integer                  'temperorily holds the current index of the next element to be copied while removing elements from the process list
    Dim lngHwndProc             As Long                     'holds the handle of the main window for the process
    Dim lngNumElem              As Long                     'holds the number of elements in the module array
    Dim lngProcLoop             As Long                     'used to cycle through the list of processes
    Dim lngModules(1 To 200)    As Long                     'holds the list of module handles
    Dim strModuleName           As String                   'holds the name of the module
    Dim intCounter              As Integer                  'used to cycle through the list of processes as we check the names against the specified pattern
    
    'initialise the parameters and return values
    intNumFound = 0
    ReDim prcList(intNumFound)
    
    'if this isn't a WinNT machine, then exit
    If Not IsWinNT Then
        GetProcessList = prcList()
        Exit Function
    End If
    
    'if the user didn't specify any patter default to "all" processes
    If (Trim(strPattern) = "") Then
        strPattern = "*"
    End If
    
    lngCbSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
    lngCbSizeRet = 96
 
    Do While (lngCbSize <= lngCbSizeRet)
        'Increment Size
        lngCbSize = lngCbSize * 2
        
        'Allocate Memory for Array
        ReDim lngPIDs(lngCbSize / 4)
        
        'Get Process ID's
        lngResult = EnumProcesses(lngPIDs(1), lngCbSize, lngCbSizeRet)
    Loop
    
    'Count number of processes returned
    lngNumElem = lngCbSizeRet / 4
    
    'Loop thru each process
    For lngProcLoop = 1 To lngNumElem
        
        'Get a handle to the Process and Open
        lngHwndProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, _
                                  0, _
                                  lngPIDs(lngProcLoop))
        
        If lngHwndProc <> 0 Then
            'Get an array of the module handles for the specified process
            lngResult = EnumProcessModules(lngHwndProc, lngModules(1), 200, lngCbSize2)
            
            'If the Module Array is retrieved, Get the ModuleFileName
            If lngResult <> 0 Then
                
                'Buffer with spaces first to allocate memory for byte array
                strModuleName = Space(MAX_PATH)
                
                'Must be set prior to calling API
                lngSize = 500
                
                'Get Process Name
                lngResult = GetModuleFileNameExA(lngHwndProc, _
                                                 lngModules(1), _
                                                 strModuleName, _
                                                 lngSize)
                
                'Remove trailing spaces
                strProcName = Left(strModuleName, lngResult)
                
                'gstrCheck for Matching Upper case result
                strProcName = Trim$(strProcName)
                strProcName = Replace(strProcName, "\??\", "")
                
                'Get the Site of the Memory Structure
                udtPMC.cb = LenB(udtPMC)
                lngResult = GetProcessMemoryInfo(lngHwndProc, udtPMC, udtPMC.cb)
                
                'add the information to the array
                ReDim Preserve prcList(intNumFound)
                With prcList(intNumFound)
                    .lngPID = lngPIDs(lngProcLoop)
                    .strProcName = strProcName
                    
                    'could we get information about the process memory
                    If (lngResult <> 0) Then
                        .pmcProcMemInfo = udtPMC
                    End If  'could we get process memory information
                End With    'prcList(intNumFound)
                intNumFound = intNumFound + 1
            End If  'could we get the module array
        End If  'could we get a process handle
        
        'Close the handle to this process
        lngResult = CloseHandle(lngHwndProc)
    Next lngProcLoop
    
    If (intNumFound > 0) Then
        'remove any processes that don't match the pattern
        intNumDel = 0
        intCounter = 0
        intIndex = 0
        Do While (intCounter <= (UBound(prcList) - intNumDel))
            
            'copy the next element down
            If (intNumDel > 0) Then
                prcList(intCounter) = prcList(intIndex)
            End If
            
            'check this element to see if it matches the pattern
            If Not (UCase(prcList(intCounter).strProcName) Like UCase(strPattern)) Then
                'rescan this element so that it will be overwritten by copying the next element from further on
                'in the array (see above code)
                intNumDel = intNumDel + 1
                intCounter = intCounter - 1
            End If
            
            'chcek the next element
            intCounter = intCounter + 1
            intIndex = (intCounter + intNumDel)
        Loop
        
        'did we delete any elements
        If (intNumDel > 0) Then
            
            intNumFound = UBound(prcList) - intNumDel
            If (intNumFound > 0) Then
                'shrink the array
                ReDim Preserve prcList(intNumFound - 1)
                
            Else
                'wipe the array
                ReDim prcList(0)
            End If
        End If  'did we delete any elements
    End If  'did we find any processes
    
    'return the process list
    GetProcessList = prcList()
End Function

Private Function GetVersion() As Long
    'get the platform number of the current operating system
    'WinNT = 2
    
    Dim osinfo      As OSVERSIONINFO
    Dim RetValue    As Integer

    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    RetValue = GetVersionEx(osinfo)
    GetVersion = osinfo.dwPlatformId
End Function

Private Function StrZToStr(s As String) _
                           As String
    'removes the last character in a string
    
    StrZToStr = Left$(s, Len(s) - 1)
End Function

Public Function GetElement(ByVal strList As String, _
                           ByVal strDelimiter As String, _
                           ByVal lngNumColumns As Long, _
                           ByVal lngRow As Long, _
                           ByVal lngColumn As Long) _
                           As String

    Dim lngCounter As Long

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Calculate the offset for the item required based on the number of columns the list
    ' 'strList' has i.e. 'lngNumColumns' and from which row the element is to be
    ' selected i.e. 'lngRow'.
    lngColumn = IIf(lngRow = 0, lngColumn, (lngRow * lngNumColumns) + lngColumn)

    ' Search for the 'lngColumn' item from the list 'strList'.
    For lngCounter = 0 To lngColumn - 1

        ' Remove each item from the list.
        strList = Mid(strList, _
                      InStr(strList, strDelimiter) + Len(strDelimiter), _
                      Len(strList))

        ' If list becomes empty before 'lngColumn' is found then just
        ' return an empty string.
        If Len(strList) = 0 Then
            GetElement = ""
            Exit Function
        End If

    Next lngCounter

    ' Return the sought list element.
    GetElement = Left$(strList, InStr(strList, strDelimiter) - 1)

End Function


''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Function GetNumElements (ByVal strList As String,
'                         ByVal strDelimiter As String)
'                         As Integer
'
'  strList      = The element list.
'  strDelimiter = The delimiter by which the elements in
'                 'strList' are seperated.
'
'  The function returns an integer which is the count of the
'  number of elements in 'strList'.
'
'  Author: Roger Taylor
'
'  Date:26/12/1998
'
'  Additional Information:
'
'  Revision History:
'
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function GetNumElements(ByVal strList As String, _
                               ByVal strDelimiter As String) _
                               As Integer

    Dim intElementCount As Integer

    ' If no elements in the list 'strList' then just return 0.
    If Len(strList) = 0 Then
        GetNumElements = 0
        Exit Function
    End If

    ' Append delimiter text to the end of the list as a terminator.
    strList = strList & strDelimiter

    ' Count the number of elements in 'strlist'
    Do While InStr(strList, strDelimiter) > 0
        intElementCount = intElementCount + 1
        strList = Mid$(strList, InStr(strList, strDelimiter) + 1, Len(strList))
    Loop

    ' Return the number of elements in 'strList'.
    GetNumElements = intElementCount
End Function

Public Sub KillProcess(ByVal strExeName As String)
    'This will stop the specified executable file from running if it
    'is active.
    
    Dim lngResult               As Long     'holds any returned error value from an api call
    Dim strProcessName          As String
    Dim lngCbSize               As Long     'Specifies the size, In bytes, of the lpidProcess array
    Dim lngCbSizeReturned       As Long     'Receives the number of bytes returned
    Dim lngNumElements          As Long
    Dim lngProcessIds()         As Long
    Dim lngCbSize2              As Long
    Dim lngModules(1 To 200)    As Long
    Dim strModuleName           As String
    Dim lngSize                 As Long
    Dim hWndProcess             As Long
    Dim lngCounter              As Long
    Dim strProcName             As String
    
    'make sure something was passed
    If Trim(strExeName) = "" Then
        Exit Sub
    End If
    
    lngCbSize = 8 ' Really needs To be 16, but Loop will increment prior to calling API
    lngCbSizeReturned = 96
    
    Do While lngCbSize <= lngCbSizeReturned
        'Increment lngSize
        lngCbSize = lngCbSize * 2
        
        'Allocate Memory for Array
        ReDim lngProcessIds(lngCbSize / 4)
        
        'Get Process ID's
        lngResult = EnumProcesses(lngProcessIds(1), lngCbSize, lngCbSizeReturned)
    Loop
    
    'Count number of processes returned
    lngNumElements = lngCbSizeReturned / 4
    
    'Loop thru each process
    For lngCounter = 1 To lngNumElements
        'Get a handle to the Process and Open
        hWndProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, lngProcessIds(lngCounter))
    
        If hWndProcess <> 0 Then
            'Get an array of the module handles for the specified process
            lngResult = EnumProcessModules(hWndProcess, lngModules(1), 200, lngCbSize2)
    
            'If the Module Array is retrieved, Get the ModuleFileName
            If lngResult <> 0 Then
    
                'Buffer with spaces first to allocate memory for byte array
                strModuleName = Space(MAX_PATH)
    
                'Must be set prior to calling API
                lngSize = 500
    
                'Get Process Name
                lngResult = GetModuleFileNameExA(hWndProcess, _
                                                 lngModules(1), _
                                                 strModuleName, _
                                                 lngSize)
    
                'Remove trailing spaces
                strProcessName = Left(strModuleName, lngResult)
    
                'gstrCheck for Matching Upper case result
                strProcessName = UCase$(Trim$(strProcessName))
    
                strProcName = GetElement(Trim(Replace(strProcessName, _
                                                      vbNullChar, _
                                                      "")), _
                                         "\", _
                                         0, _
                                         0, _
                                         GetNumElements(Trim(Replace(strProcessName, _
                                                                     vbNullChar, _
                                                                     "")), _
                                                        "\") - 1)
    
                If (UCase(Trim(strProcName)) = Trim(UCase(strExeName))) Then
                    'terminate the process
                    lngResult = TerminateProcess(hWndProcess, 0)
                End If
            End If
        End If
        
        'Close the handle to this process
        lngResult = CloseHandle(hWndProcess)
    Next
End Sub

Public Sub PcSpeaker(Optional ByVal lngFrequency As Long = 1000, _
                     Optional ByVal lngDuration As Long = 150)
    'This will create a sound using the pc speaker.
    '"Call PcSpeaker" works pretty well as a general warning. lngDuration
    'is in milliseconds, and lngFrequency must be in the range 37 to 32767.
    'In win 9x, the parameters are ignored and a default beep is used.
    
    Dim lngResult       As Long         'holds any returned value from an api call
    
    lngResult = Beep(lngFrequency, lngDuration)
End Sub

Public Sub SetAppThreadPriority(ByVal enmThreadPriority As EnumThreadPriority, _
                                Optional ByVal blnActivate As Boolean = True)
    'This will set the applications thread priority to the specified level.
    'this should be run from the Form_Load if the blnActivate is False
    
    Dim lngResult       As Long                 'holds any returned error value from an api call
    Dim enmClass        As EnumClassPriority    'holds the priority level for the class
    Dim lngAppThread    As Long                 'holds the applications main thread id
    Dim lngAppProcess   As Long                 'holds the application main class id
    
    'match the class priority to the thread priority
    Select Case enmThreadPriority
    'high priority
    Case THREAD_PRIORITY_ABOVE_NORMAL, THREAD_PRIORITY_HIGHEST
        enmClass = HIGH_PRIORITY_CLASS
    
    'normal priority
    Case THREAD_PRIORITY_NORMAL
        enmClass = NORMAL_PRIORITY_CLASS
    
    'low thread priority
    Case THREAD_PRIORITY_BELOW_NORMAL, THREAD_PRIORITY_LOWEST, THREAD_PRIORITY_IDLE
        enmClass = IDLE_PRIORITY_CLASS
    
    'real time priority
    Case THREAD_PRIORITY_TIME_CRITICAL
        enmClass = REALTIME_PRIORITY_CLASS
    End Select
    
    'make sure the application is active
    If blnActivate Then
        Call AppActivate(App.Title, True)
    End If
    
    'get the applications handle
    lngAppThread = GetCurrentThread
    lngAppProcess = GetCurrentProcess
    
    'set the priorities
    lngResult = SetThreadPriority(lngAppThread, enmThreadPriority)
    lngResult = SetPriorityClass(lngAppProcess, enmClass)
End Sub

Public Sub RemoveX(ByVal hWnd As Long)
    'this will remove the Close menu item from the specified window
    
    Dim hSysMenu    As Long         'holds a handle to the windows system menu
    Dim lngResult   As Long         'holds any returned value from an api call
    
    'get the system menu for this form
    hSysMenu = GetSystemMenu(hWnd, 0)
    
    'remove the close item
    Call RemoveMenu(hSysMenu, 6, MF_BYPOSITION)
    
    'remove the separator that was over the close item
    Call RemoveMenu(hSysMenu, 5, MF_BYPOSITION)
End Sub


Public Sub RemoveControlMenu(ByVal hWnd As Long, _
                             ByVal strMenuName As String)
    'this will remove the Close menu item from the specified window
    
    Dim hSysMenu    As Long         'holds a handle to the windows system menu
    Dim lngResult   As Long         'holds any returned value from an api call
    
    'we need to remove the Break if the menu item is Close
    If (UCase(Trim(strMenuName)) = "CLOSE") Then
        Call RemoveX(hWnd)
    End If
    
    'get the system menu for this form
    hSysMenu = GetSystemMenu(hWnd, 0)
    
    'remove the close item
    lngResult = DeleteMenu(hSysMenu, 3, MF_BYPOSITION)
    
    'redraw the menu
    lngResult = DrawMenuBar(hWnd)
    
    'remove the separator that was over the close item
    'Call RemoveMenu(hSysMenu, ByVal strMenuName, MF_BYCOMMAND)
End Sub

Public Sub ResetSystemMenu(ByVal hWnd As Long)
    'This will reset the windows system menu to its default state. Any changed made to it are deleted
    
    
    Dim hSysMenu    As Long         'holds a handle to the windows system menu
    
    'get the system menu for this form
    hSysMenu = GetSystemMenu(hWnd, True)
End Sub

Public Function GetFont(ByVal enmFont As EnumSysFonts) As StdFont
    'This will return the specified system font details
    
    Dim ncm             As NONCLIENTMETRICS     'holds the system font metrics
    Dim lngResult       As Long                 'holds any returned value from an api call
    Dim strBuffer       As String               'holds the name of the font
    Dim fntReturn       As StdFont              'holds the font details to return back to the user
    Dim lfnFont         As LOGFONT              'holds the specific font details that we need to translate
    
    'initialise the font variable
    Set fntReturn = Nothing
    
    'get the font metrics
    ncm.cbSize = 340
    lngResult = SystemParametersInfo(SPI_GETNONCLIENTMETRICS, ncm.cbSize, ncm, 0)
    
    'were we able to get the font details
    If lngResult = 0 Then
        Set GetFont = fntReturn
        Exit Function
    End If  'were we able to get the font details
    
    'what font details are we meant to get
    Select Case enmFont
    Case FNT_CAPTION
        lfnFont = ncm.lfCaptionFont
        
    Case FNT_MENU
        lfnFont = ncm.lfMenuFont
        
    Case FNT_MESSAGE
        lfnFont = ncm.lfMessageFont
        
    Case FNT_STATUS
        lfnFont = ncm.lfStatusFont
        
    Case Else
        'invalid parameter value
        Set GetFont = fntReturn
        Exit Function
    End Select
    
    'create the font object
    Set fntReturn = New StdFont
    
    With fntReturn
        
        'convert the pixel size to a font size
        .Size = CInt(-0.75 * lfnFont.lfHeight)
        
        'is the font bold
        If (lfnFont.lfWeight < 700) Then
            .Bold = False
        Else
            .Bold = True
        End If
        
        .Italic = lfnFont.lfItalic
        .Name = StrConv(lfnFont.lfFaceName, vbUnicode)
        .Strikethrough = lfnFont.lfStrikeOut
        .Underline = lfnFont.lfUnderline
        .Weight = lfnFont.lfWeight
    End With    'fntReturn
    
    'return the font details
    Set GetFont = fntReturn
End Function

Public Sub SetFormFontsToSystem(ByRef frmSetSysFonts As Form, _
                                Optional ByVal enmSysFont As EnumSysFonts = FNT_MESSAGE)
    'This will cycle through all the controls on a form and if they're font has been unchange (ie, the vb
    'default font), then it will be set to the details of the specified system font.
    
    Dim ctlCounter      As Control      'used for cycling through all the controls in the form
    Dim fntSystem       As StdFont      'holds a reference to the details of the system font
    
    If (frmSetSysFonts Is Nothing) Then
        'we need a form to be passed
        Exit Sub
    End If
    
    'get the font details
    Set fntSystem = GetFont(enmSysFont)
    
    'were we able to get the specified system font
    If (fntSystem Is Nothing) Then
        Exit Sub
    End If
    
    On Error Resume Next
    
    'set the font for all controls with the default vb font
    For Each ctlCounter In frmSetSysFonts.Controls
        
        'reset the error object so we can test if this control has a font
        Call Err.Clear
        
        With ctlCounter.Font
            
            'if the With statement caused an error, then the control does NOT have a Font object to set
            If (Err.Number = 0) Then
                
                'is the font of this control the vb default
                If (.Name = "MS Sans Serif") And _
                   (.Size = 8.25) And _
                   (.Bold = False) And _
                   (.Italic = False) And _
                   (.Underline = False) And _
                   (.Strikethrough = False) Then
                   
                   'set the new font - a new instance of a Font object for each control
                   Set fntSystem = GetFont(enmSysFont)
                   Set ctlCounter.Font = fntSystem
                End If  'is this the default vb font
                
            End If  'does the control have a Font object
        End With    'ctlcounter.Font
    Next ctlCounter
End Sub
