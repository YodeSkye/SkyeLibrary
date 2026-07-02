
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        Public Const WM_HOTKEY As Integer = 786
        Public Const MOD_SHIFT As Integer = 4
        Public Const MOD_CONTROL As Integer = 2
        Public Const MOD_ALT As Integer = 1
        Public Const MOD_WIN As Integer = 8

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function RegisterHotKey(hWnd As IntPtr, id As Integer, fsModifiers As UInteger, vk As UInteger) As Boolean
        End Function

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function UnregisterHotKey(hWnd As IntPtr, id As Integer) As Boolean
        End Function

    End Class

End Namespace
#Enable Warning CA1401
