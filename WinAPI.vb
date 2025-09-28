
Imports System.Drawing
Imports System.Runtime.InteropServices

''' <summary>
''' Declarations for Windows API functions, structures, and constants, along with some helper functions.
''' </summary>
Public Class WinAPI

    'Declarations
    Public Const HWND_BROADCAST As Integer = 65535
    Public Const WM_SYSCOMMAND As Integer = 274 '&H112
    Public Const SC_MAXIMIZE As UShort = 61488 '&HF030
    Public Const SC_MAXIMIZE_TBAR As UShort = 61490 '&HF032 'NOT A WINDOWS CONSTANT 'This value is passed to WndProc when DoubleClicking on the Title Bar to Maximize a window.
    Public Const SC_RESTORE As UShort = 61728 '&HF120
    Public Const SC_RESTORE_TBAR As UShort = 61730 '&HF122 'NOT A WINDOWS CONSTANT 'This value is passed to WndProc when DoubleClicking on the Title Bar to Restore a window.
    Public Const SC_CLOSE As UShort = 61536 '&HF060
    Public Const WM_ACTIVATE As UShort = &H6
    Public Const WM_PAINT As Integer = &HF
    Public Const WM_RIGHTCLICKTASKBAR As Integer = 787 '&H313 'UnDocumented Message passed when a user RightClicks on the app taskbar entry before the SystemMenu is displayed. Unless this message is fowarded, the SystemMenu won't display. Useful for replacing with custom menus or for when FormBorderStyle is set to None and the SystemMenu never displays. Just intercept and show your own custom menu. XP & earlier; after XP, ShiftRightClick is required.
    Public Const WM_LBUTTONDOWN As Integer = &H201
    Public Const WM_LBUTTONDBLCLK As Integer = &H203
    Public Const WM_LBUTTONUP As Integer = &H202
    Public Const WM_RBUTTONUP As Integer = &H205
    Public Const WM_RBUTTONDOWN As Integer = &H204
    Public Const WM_MOUSEACTIVATE As Integer = &H21
    Public Const MA_NOACTIVATE As Integer = 3
    Public Const WM_NCACTIVATE As Integer = &H86
    Public Const WM_SIZE As Integer = &H5
    Public Const WM_GET_CUSTOM_DATA As UInteger = &H8001
    Public Declare Auto Function GetClassName Lib "user32.dll" (ByVal hwnd As IntPtr, ByVal lpClassName As String, ByVal nMaxCount As Integer) As Integer
    Public Declare Auto Function SendMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As IntPtr
    Public Declare Auto Function PostMessage Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal Msg As UInteger, ByVal wParam As IntPtr, ByVal lParam As IntPtr) As Boolean
    Public Declare Auto Function LockWorkStation Lib "user32.dll" () As Boolean
    Public Declare Auto Function IsWow64Process Lib "kernel32.dll" (ByVal hProcess As IntPtr, ByRef Wow64Process As Boolean) As Boolean

    'ClipBoard
    Public Const WM_CHANGECBCHAIN As Integer = 781 '&H30D
    Public Const WM_DRAWCLIPBOARD As Integer = 776 '&H308
    Public Declare Auto Function SetClipboardViewer Lib "user32.dll" (ByVal hWndNewViewer As IntPtr) As IntPtr 'adds the specified window to the chain of clipboard viewers
    Public Declare Auto Function ChangeClipboardChain Lib "user32.dll" (ByVal hWndRemove As IntPtr, ByVal hWndNewNext As IntPtr) As Boolean 'removes a specified window from the chain of clipboard viewers
    Public Declare Auto Function GetClipboardSequenceNumber Lib "user32.dll" () As UInteger

    'Window Functions 'Get Information About & Change Attributes Of A Window
    'HideFormInTaskSwitcher, Customize Minimize/Maximize/Restore Functions
    'Usage In Forms: When Custom Maximizing(to properly set the maximize icon & window menu), simply set Form.WindowState to Normal when restoring: SetWindowLong(Me.Handle, GWL_STYLE, GetWindowLong(Me.Handle, GWL_STYLE) Or WS_MAXIMIZE)
    'Detecting if an App is in FullScreen Mode
    Public Const GWL_STYLE As Integer = -16
    Public Const GWL_EXSTYLE As Integer = -20
    Public Const WS_EX_TOOLWINDOW As Integer = 128 '&H80
    Public Const WS_EX_TOPMOST As Integer = &H8
    Public Const WS_POPUP As Integer = &H80000000
    Public Const WS_EX_NOACTIVATE As Integer = &H8000000
    Public Const WS_EX_TRANSPARENT As Integer = &H20
    Public Const WS_MINIMIZEBOX As Integer = 131072 '&H20000 'Turn on the WS_MINIMIZEBOX style flag for borderless windows so you can minimize/restore from the taskbar.
    Public Const WS_MAXIMIZE As Integer = 16777216 '&H1000000
    Public Const SW_SHOWNOACTIVATE As Integer = 4
    Public Const SWP_NOACTIVATE As UInteger = &H10
    Public Const SWP_SHOWWINDOW As UInteger = &H40
    Public Const SWP_NOZORDER As UInteger = &H4
    Public Const SWP_NOMOVE As UInteger = &H2
    Public Const SWP_NOSIZE As UInteger = &H1
    Public Const SWP_FRAMECHANGED As UInteger = &H20
    Public Shared ReadOnly HWND_TOPMOST As IntPtr = New IntPtr(-1)
    Public Structure RECT
        Dim Left As Integer
        Dim Top As Integer
        Dim Right As Integer
        Dim Bottom As Integer
    End Structure
    Public Declare Auto Function GetForegroundWindow Lib "user32.dll" () As IntPtr
    Public Declare Auto Function SetForegroundWindow Lib "user32.dll" (ByVal hWnd As IntPtr) As Boolean
    Public Declare Auto Function FindWindow Lib "user32.dll" (lpClassName As String, lpWindowName As String) As IntPtr
    Public Declare Auto Function GetWindowThreadProcessId Lib "user32.dll" (hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
    Public Declare Auto Function GetWindowText Lib "user32.dll" (hWnd As IntPtr, lpString As String, nMaxCount As Integer) As Integer
    Public Declare Auto Function GetWindowRect Lib "user32.dll" (hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
    Public Declare Auto Function GetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer) As Integer
    Public Declare Auto Function SetWindowLong Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal nIndex As Integer, ByVal dwNewLong As Integer) As Integer
    Public Declare Auto Function SetWindowPos Lib "user32.dll" (hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
    Public Declare Auto Function ShowWindow Lib "user32.dll" (hWnd As IntPtr, nCmdShow As Integer) As Boolean
    Public Declare Auto Function IsWindow Lib "user32.dll" (ByVal hwnd As IntPtr) As Boolean

    'Causes a window or control to use a different set of visual style information than its class normally uses.
    Public Declare Auto Function SetWindowTheme Lib "uxtheMe.dll" (hwnd As IntPtr, appname As String, idlist As String) As Integer

    'This function retrieves the current color of the specified display element.
    Public Const COLOR_WINDOW As Integer = 5
    Public Const COLOR_WINDOWTEXT As Integer = 8
    Public Const COLOR_HIGHLIGHT As Integer = 13
    Public Const COLOR_3DFACE As Integer = 15
    Public Const COLOR_GRAYTEXT As Integer = 17
    Public Const COLOR_HOTLIGHT As Integer = 26
    Public Declare Auto Function GetSysColor Lib "user32.dll" (nIndex As Integer) As Integer

    'Keyboard Functions 'Synthesize a keystroke
    Public Const VK_CAPITAL As Integer = &H14
    Public Const VK_SCROLL As Integer = &H91
    Public Const VK_NUMLOCK As Integer = &H90
    Public Const VK_MEDIA_NEXT_TRACK As Integer = 176 '&B0
    Public Const VK_MEDIA_PREV_TRACK As Integer = 177 '&B1
    Public Const VK_MEDIA_STOP As Integer = 178 '&B2
    Public Const VK_MEDIA_PLAY_PAUSE As Integer = 179 '&B3
    Public Const KEYEVENTF_EXTENDEDKEY As Integer = &H1
    Public Const KEYEVENTF_KEYUP As Integer = &H2
    Public Declare Auto Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Integer, ByVal dwExtraInfo As Integer)

    'HotKeys
    Public Const WM_HOTKEY As Integer = 786 '&H312
    Public Const MOD_SHIFT As Integer = 4 '&H4
    Public Const MOD_CONTROL As Integer = 2 '&H2
    Public Const MOD_ALT As Integer = 1 '&H1
    Public Const MOD_WIN As Integer = 8 '&H8
    Public Declare Auto Function RegisterHotKey Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal id As Integer, ByVal fsModifiers As UInteger, ByVal vk As UInteger) As Boolean
    Public Declare Auto Function UnregisterHotKey Lib "user32.dll" (ByVal hWnd As IntPtr, ByVal id As Integer) As Boolean

    'ScreenSaver
    Public Enum EXECUTION_STATE As UInteger
        ES_AWAYMODE_REQUIRED = 64 '&H40
        ES_CONTINUOUS = 2147483648 '&H80000000
        ES_DISPLAY_REQUIRED = 2 '&H2
        ES_SYSTEM_REQUIRED = 1 '&H1
    End Enum
    Public Const SPI_GETSCREENSAVERRUNNING As Integer = 114
    Public Const SC_SCREENSAVE As UShort = 61760 '&HF140
    Public Declare Auto Function SystemParametersInfo Lib "user32.dll" (ByVal uiAction As UInteger, ByVal uiParam As UInteger, ByRef pvParam As Boolean, ByVal fWinIni As UInteger) As Boolean
    Public Declare Auto Function SetThreadExecutionState Lib "kernel32.dll" (ByVal esFlags As EXECUTION_STATE) As EXECUTION_STATE

    'AppCommand, For Catching Media Keys
    Public Const WM_APPCOMMAND As Integer = &H319
    Public Const APPCOMMAND_MEDIA_NEXTTRACK As Integer = 720896
    Public Const APPCOMMAND_MEDIA_PREVIOUSTRACK As Integer = 786432
    Public Const APPCOMMAND_MEDIA_STOP As Integer = 851968
    Public Const APPCOMMAND_MEDIA_PLAY_PAUSE As Integer = 917504

    'Windows Accent Color
    Public Const WM_DWMCOLORIZATIONCOLORCHANGED As Integer = &H320
    Public Structure DWMCOLORIZATIONPARAMS
        Public ColorizationColor As UInteger
        Public ColorizationAfterglow As UInteger
        Public ColorizationColorBalance As UInteger
        Public ColorizationAfterglowBalance As UInteger
        Public ColorizationBlurBalance As UInteger
        Public ColorizationGlassReflectionIntensity As UInteger
        Public ColorizationOpaqueBlend As UInteger
    End Structure
    <DllImport("dwmapi.dll", EntryPoint:="#127", PreserveSig:=False)>
    Public Shared Sub DwmGetColorizationParameters(<Out> ByRef parameters As DWMCOLORIZATIONPARAMS)
    End Sub

    'GetApplicationIcon, & Other File Info
    Public Const MAX_PATH As Integer = 260 '&H104
    Public Const SHGFI_ICON As Integer = 256 '&H100 'Retrieve the handle to the icon that represents the file and the index of the icon within the system image list.
    Public Const SHGFI_SMALLICON As Integer = 1 '&H1 '16x16, Modifies SHGFI_ICON
    Public Const SHGFI_LARGEICON As Integer = 0 '&H0 '32x32, Modifies SHGFI_ICON
    Public Const SHGFI_DISPLAYNAME As Integer = 512 '&H200 'Retrieve the display name for the file. The name is copied to the szDisplayName member of the structure specified in psfi.
    Public Const SHGFI_TYPENAME As Integer = 1024 '&H400 'Retrieve the string that describes the file's type. The string is copied to the szTypeName member of the structure specified in psfi.
    Public Structure SHFILEINFO
        Dim hIcon As IntPtr
        Dim iIcon As Integer
        Dim dwAttributes As Integer
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=MAX_PATH)> Dim szDisplayName As String
        <Runtime.InteropServices.MarshalAs(Runtime.InteropServices.UnmanagedType.ByValTStr, SizeConst:=80)> Dim szTypeName As String
    End Structure
    Public Declare Ansi Function SHGetFileInfo Lib "shell32.dll" (ByVal pszPath As String, ByVal dwFileAttributes As Integer, ByRef psfi As SHFILEINFO, ByVal cbFileInfo As Integer, ByVal uFlags As Integer) As IntPtr
    Public Declare Ansi Function DestroyIcon Lib "user32.dll" (ByVal hIcon As IntPtr) As Boolean

    'IdleTime, by getting the tick count of the last time the user provided input to the session.
    Public Structure LASTINPUTINFO
        Dim cbSize As UInteger
        Dim dwTime As UInteger
    End Structure
    Public Declare Auto Function GetLastInputInfo Lib "user32.dll" (ByRef plii As LASTINPUTINFO) As Boolean

    'Procedures

    ''' <summary>
    ''' To Remove A Borderless Window From Windows TaskSwitcher
    ''' </summary>
    ''' <param name="handle">Your Form's Handle</param>
    ''' <returns>True Is Success, Otherwise False</returns>
    Public Shared Function HideFormInTaskSwitcher(handle As IntPtr) As Boolean
        Try
            Dim exStyle As Integer = GetWindowLong(handle, GWL_EXSTYLE)
            Dim result As Integer = SetWindowLong(handle, GWL_EXSTYLE, exStyle Or WS_EX_TOOLWINDOW)
            Return result <> 0
        Catch
            Return False
        End Try
        'Try
        '    If SetWindowLong(handle, GWL_EXSTYLE, (GetWindowLong(handle, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW)) = 0 Then : Return False
        '    Else : Return True
        '    End If
        'Catch : Return False
        'End Try
    End Function

    ''' <summary>
    ''' Retrieves the icon associated with a file path, optionally returning the large or small version.
    ''' </summary>
    ''' <param name="filepath">The full path to the file.</param>
    ''' <param name="getLargeIcon">True to retrieve the large icon; False for small.</param>
    ''' <returns>The extracted icon, or Nothing if retrieval fails.</returns>
    Public Shared Function GetApplicationIcon(filepath As String, Optional getlargeicon As Boolean = False) As Icon

        filepath = filepath.Substring(0, Math.Min(filepath.Length, MAX_PATH))

        Dim shinfo As New SHFILEINFO()
        Dim flags As Integer = SHGFI_ICON Or If(getlargeicon, SHGFI_LARGEICON, SHGFI_SMALLICON)
        Dim status As IntPtr = SHGetFileInfo(filepath, 0, shinfo, Marshal.SizeOf(shinfo), flags)

        If status = IntPtr.Zero OrElse shinfo.hIcon = IntPtr.Zero Then Return Nothing

        Try
            Using tempIcon = Icon.FromHandle(shinfo.hIcon)
                Return DirectCast(tempIcon.Clone(), Icon)
            End Using
        Finally
            DestroyIcon(shinfo.hIcon)
        End Try

        'filepath = Microsoft.VisualBasic.Left(filepath, MAX_PATH)
        'Dim windowsfileinfo As New SHFILEINFO
        'Dim status As IntPtr = IntPtr.Zero
        'Dim nIcon As Icon
        'Try
        '    If getlargeicon Then : status = SHGetFileInfo(filepath, 0, windowsfileinfo, Runtime.InteropServices.Marshal.SizeOf(windowsfileinfo), SHGFI_ICON Or SHGFI_LARGEICON)
        '    Else : status = SHGetFileInfo(filepath, 0, windowsfileinfo, Runtime.InteropServices.Marshal.SizeOf(windowsfileinfo), SHGFI_ICON Or SHGFI_SMALLICON)
        '    End If
        '    If status = IntPtr.Zero Then : Throw New Exception
        '    Else
        '        nIcon = Icon.FromHandle(windowsfileinfo.hIcon)
        '        GetApplicationIcon = DirectCast(nIcon.Clone, Icon)
        '        nIcon.Dispose()
        '        If Not DestroyIcon(windowsfileinfo.hIcon) Then Throw New Exception
        '    End If
        'Catch : GetApplicationIcon = Nothing
        'Finally
        '    nIcon = Nothing
        '    windowsfileinfo = Nothing
        'End Try
    End Function

    ''' <summary>
    ''' Gets The Session Idle Time
    ''' </summary>
    ''' <returns>The number of Ticks since the last user input.</returns>
    Public Shared Function GetIdleTime() As UInteger
        Dim lastInput As New LASTINPUTINFO()
        lastInput.cbSize = CUInt(Marshal.SizeOf(lastInput))
        If Not GetLastInputInfo(lastInput) Then Return 0
        Return CUInt(Environment.TickCount) - lastInput.dwTime
        'Dim lastInput As New LASTINPUTINFO
        'lastInput.cbSize = CUInt(Runtime.InteropServices.Marshal.SizeOf(lastInput))
        'GetLastInputInfo(lastInput)
        'GetIdleTime = (CUInt(Environment.TickCount) - lastInput.dwTime)
        'lastInput = Nothing
    End Function

    ''' <summary>
    ''' Functions to extract the color elements of a windows system color. Each returns an Red, Green, or Blue element value.
    ''' </summary>
    ''' <param name="color">A Windows system color integer.</param>
    ''' <returns>A Short containing the extracted color element value.</returns>
    Public Shared Function GetRValue(ByVal color As Long) As Short
        GetRValue = CShort(color And &HFF&)
    End Function
    Public Shared Function GetGValue(ByVal color As Long) As Short
        GetGValue = CShort((color \ &H100&) And &HFF&)
    End Function
    Public Shared Function GetBValue(ByVal color As Long) As Short
        GetBValue = CShort((color \ &H10000) And &HFF&)
    End Function

    ''' <summary>
    ''' Gets the Windows system color of a certain display element.
    ''' </summary>
    ''' <param name="nIndex">The system color index (e.g., COLOR_WINDOW, COLOR_HIGHLIGHT).</param>
    ''' <returns>A .NET Color.</returns>
    Public Shared Function GetSystemColor(nIndex As Integer) As Color
        Dim colorvalue As Integer = GetSysColor(nIndex)
        Return System.Drawing.Color.FromArgb(255, GetRValue(colorvalue), GetGValue(colorvalue), GetBValue(colorvalue))
    End Function

End Class
