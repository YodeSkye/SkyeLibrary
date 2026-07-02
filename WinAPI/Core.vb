
Imports System.Runtime.InteropServices
Imports System.Text

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' ============================
        ' CORE CONSTANTS & MESSAGES
        ' ============================
        Public Const HWND_BROADCAST As Integer = 65535
        Public Const HTCLIENT As Integer = 1
        Public Const HTCAPTION As Integer = 2
        Public Const HTTRANSPARENT As Integer = -1
        Public Const MA_NOACTIVATE As Integer = 3
        Public Const ULW_ALPHA As Integer = &H2
        Public Const AC_SRC_OVER As Byte = &H0
        Public Const AC_SRC_ALPHA As Byte = &H1
        Public Const SC_MINIMIZE As Integer = &HF020
        Public Const SC_MAXIMIZE As UShort = 61488
        Public Const SC_MAXIMIZE_TBAR As UShort = 61490
        Public Const SC_RESTORE As UShort = 61728
        Public Const SC_RESTORE_TBAR As UShort = 61730
        Public Const SC_CLOSE As UShort = 61536
        Public Const WM_SYSCOMMAND As Integer = 274
        Public Const WM_ACTIVATE As UShort = &H6
        Public Const WM_ACTIVATEAPP As Integer = &H1C
        Public Const WM_NCPAINT As Integer = &H85
        Public Const WM_SETREDRAW As Integer = &HB
        Public Const WM_SETFOCUS As Integer = &H7
        Public Const WM_KILLFOCUS As Integer = &H8
        Public Const WM_PRINTCLIENT As Integer = &H318
        Public Const WM_RIGHTCLICKTASKBAR As Integer = 787
        Public Const WM_NCACTIVATE As Integer = &H86
        Public Const WM_SIZE As Integer = &H5
        Public Const WM_GET_CUSTOM_DATA As UInteger = &H8001
        Public Const WM_SETTINGCHANGE As Integer = &H1A
        Public Const WM_THEMECHANGED As Integer = &H31A
        Public Const WM_SYSCOLORCHANGE As Integer = &H15
        Public Const WM_DPICHANGED As Integer = &H2E0
        Public Const WM_PARENTNOTIFY As Integer = &H210
        Public Const WM_CANCELMODE As Integer = &H1F
        Public Const WM_NCHITTEST As Integer = &H84
        Public Const WM_CLOSE As Integer = &H10
        Public Const WM_DESTROY As Integer = &H2

        ' ============================
        ' WNDCLASSEX
        ' ============================
        <StructLayout(LayoutKind.Sequential, CharSet:=CharSet.Unicode)>
        Public Structure WNDCLASSEX
            Public cbSize As UInteger
            Public style As UInteger
            Public lpfnWndProc As IntPtr
            Public cbClsExtra As Integer
            Public cbWndExtra As Integer
            Public hInstance As IntPtr
            Public hIcon As IntPtr
            Public hCursor As IntPtr
            Public hbrBackground As IntPtr
            Public lpszMenuName As String
            Public lpszClassName As String
            Public hIconSm As IntPtr
        End Structure

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function RegisterClassEx(ByRef wc As WinAPI.WNDCLASSEX) As UShort
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function CreateWindowEx(dwExStyle As Integer,
                                              lpClassName As String,
                                              lpWindowName As String,
                                              dwStyle As Integer,
                                              X As Integer,
                                              Y As Integer,
                                              nWidth As Integer,
                                              nHeight As Integer,
                                              hWndParent As IntPtr,
                                              hMenu As IntPtr,
                                              hInstance As IntPtr,
                                              lpParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Public Shared Function DefWindowProc(hWnd As IntPtr, msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function GetClassName(hwnd As IntPtr, lpClassName As StringBuilder, nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SendMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function PostMessage(hWnd As IntPtr, Msg As UInteger, wParam As IntPtr, lParam As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function LockWorkStation() As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function IsWow64Process(hProcess As IntPtr, <Out> ByRef Wow64Process As Boolean) As Boolean
        End Function

        ' ============================
        ' RECT / POINT / SIZE
        ' ============================
        <StructLayout(LayoutKind.Sequential)>
        Public Structure RECT
            Public Left As Integer
            Public Top As Integer
            Public Right As Integer
            Public Bottom As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure POINT
            Public X As Integer
            Public Y As Integer
        End Structure

        <StructLayout(LayoutKind.Sequential)>
        Public Structure SIZE
            Public cx As Integer
            Public cy As Integer
        End Structure

        ' ============================
        ' ENUMWINDOWS
        ' ============================
        Public Delegate Function EnumWindowsProc(hWnd As IntPtr, lParam As IntPtr) As Boolean

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function EnumWindows(lpEnumFunc As EnumWindowsProc, lParam As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetForegroundWindow() As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetForegroundWindow(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetWindow(hWnd As IntPtr, uCmd As UInteger) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function FindWindow(lpClassName As String, lpWindowName As String) As IntPtr
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function FindWindowEx(parent As IntPtr, childAfter As IntPtr, lpszClass As String, lpszWindow As String) As IntPtr
        End Function

        <DllImport("user32.dll")>
        Public Shared Function IsWindowVisible(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetWindowThreadProcessId(hWnd As IntPtr, ByRef lpdwProcessId As UInteger) As UInteger
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function GetWindowText(hWnd As IntPtr, lpString As StringBuilder, nMaxCount As Integer) As Integer
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetWindowRect(hWnd As IntPtr, ByRef lpRect As RECT) As Boolean
        End Function

        ' ============================
        ' WINDOWLONG / WINDOWPOS
        ' ============================
        <DllImport("user32.dll", EntryPoint:="GetWindowLong", SetLastError:=True)>
        Public Shared Function GetWindowLong(hWnd As IntPtr, nIndex As Integer) As Integer
        End Function

        <DllImport("user32.dll", EntryPoint:="SetWindowLong", SetLastError:=True)>
        Public Shared Function SetWindowLong(hWnd As IntPtr, nIndex As Integer, dwNewLong As Integer) As Integer
        End Function

        <DllImport("user32.dll", EntryPoint:="GetWindowLongPtr", SetLastError:=True)>
        Public Shared Function GetWindowLongPtr(hWnd As IntPtr, nIndex As Integer) As IntPtr
        End Function

        <DllImport("user32.dll", EntryPoint:="SetWindowLongPtr", SetLastError:=True)>
        Public Shared Function SetWindowLongPtr(hWnd As IntPtr, nIndex As Integer, dwNewLong As IntPtr) As IntPtr
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetWindowPos(hWnd As IntPtr, hWndInsertAfter As IntPtr, X As Integer, Y As Integer, cx As Integer, cy As Integer, uFlags As UInteger) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function ShowWindow(hWnd As IntPtr, nCmdShow As Integer) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function IsWindow(hWnd As IntPtr) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function RedrawWindow(hWnd As IntPtr, lprcUpdate As IntPtr, hrgnUpdate As IntPtr, flags As UInteger) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function DestroyWindow(hWnd As IntPtr) As Boolean
        End Function

        ' ============================
        ' CLASS STYLES
        ' ============================
        Public Const CS_VREDRAW As Integer = &H1
        Public Const CS_HREDRAW As Integer = &H2
        Public Const CS_DBLCLKS As Integer = &H8
        Public Const CS_OWNDC As Integer = &H20
        Public Const CS_CLASSDC As Integer = &H40
        Public Const CS_PARENTDC As Integer = &H80
        Public Const CS_NOCLOSE As Integer = &H200
        Public Const CS_SAVEBITS As Integer = &H800
        Public Const CS_BYTEALIGNCLIENT As Integer = &H1000
        Public Const CS_BYTEALIGNWINDOW As Integer = &H2000
        Public Const CS_GLOBALCLASS As Integer = &H4000
        Public Const CS_IME As Integer = &H10000
        Public Const CS_DROPSHADOW As Integer = &H20000

        ' ============================
        ' EDIT CONTROL MARGINS
        ' ============================
        Public Const EM_SETMARGINS As Integer = &HD3
        Public Const EC_LEFTMARGIN As Integer = &H1
        Public Const EC_RIGHTMARGIN As Integer = &H2

        ' ============================
        ' SetWindowTheme
        ' ============================
        <DllImport("uxtheme.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function SetWindowTheme(hwnd As IntPtr, appname As String, idlist As String) As Integer
        End Function

        ' ============================
        ' COLORS
        ' ============================
        Public Const COLOR_WINDOW As Integer = 5
        Public Const COLOR_WINDOWTEXT As Integer = 8
        Public Const COLOR_HIGHLIGHT As Integer = 13
        Public Const COLOR_3DFACE As Integer = 15
        Public Const COLOR_GRAYTEXT As Integer = 17
        Public Const COLOR_HOTLIGHT As Integer = 26

        ' ============================
        ' SCREENSAVER / EXECUTION STATE
        ' ============================
        Public Enum EXECUTION_STATE As UInteger
            ES_AWAYMODE_REQUIRED = &H40
            ES_CONTINUOUS = &H80000000UI
            ES_DISPLAY_REQUIRED = &H2
            ES_SYSTEM_REQUIRED = &H1
        End Enum

        Public Const SPI_GETSCREENSAVERRUNNING As Integer = 114
        Public Const SC_SCREENSAVE As UShort = 61760

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SystemParametersInfo(uiAction As UInteger, uiParam As UInteger, ByRef pvParam As Boolean, fWinIni As UInteger) As Boolean
        End Function

        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function SystemParametersInfo(uiAction As UInteger, uiParam As UInteger, pvParam As String, fWinIni As UInteger) As Boolean
        End Function

        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function SetThreadExecutionState(esFlags As EXECUTION_STATE) As EXECUTION_STATE
        End Function

        ' ============================
        ' APPCOMMAND
        ' ============================
        Public Const WM_APPCOMMAND As Integer = &H319
        Public Const APPCOMMAND_MEDIA_NEXTTRACK As Integer = 720896
        Public Const APPCOMMAND_MEDIA_PREVIOUSTRACK As Integer = 786432
        Public Const APPCOMMAND_MEDIA_STOP As Integer = 851968
        Public Const APPCOMMAND_MEDIA_PLAY_PAUSE As Integer = 917504

        ' ============================
        ' DWM COLORIZATION
        ' ============================
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

    End Class

End Namespace
#Enable Warning CA1401
