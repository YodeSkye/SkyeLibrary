Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        Public Const WM_MOUSEMOVE As Integer = &H200
        Public Const WM_SETCURSOR As Integer = &H20
        Public Const WM_LBUTTONDOWN As Integer = &H201
        Public Const WM_LBUTTONDBLCLK As Integer = &H203
        Public Const WM_LBUTTONUP As Integer = &H202
        Public Const WM_RBUTTONUP As Integer = &H205
        Public Const WM_RBUTTONDOWN As Integer = &H204
        Public Const WM_NCLBUTTONDOWN As Integer = &HA1
        Public Const WM_NCLBUTTONUP As Integer = &HA2
        Public Const WM_NCRBUTTONDOWN As Integer = &HA4
        Public Const WM_NCRBUTTONUP As Integer = &HA5
        Public Const WM_CONTEXTMENU As Integer = &H7B
        Public Const WM_MOUSEACTIVATE As Integer = &H21

        Public Const VK_CAPITAL As Integer = &H14
        Public Const VK_SCROLL As Integer = &H91
        Public Const VK_NUMLOCK As Integer = &H90
        Public Const VK_MEDIA_NEXT_TRACK As Integer = 176
        Public Const VK_MEDIA_PREV_TRACK As Integer = 177
        Public Const VK_MEDIA_STOP As Integer = 178
        Public Const VK_MEDIA_PLAY_PAUSE As Integer = 179

        Public Const KEYEVENTF_EXTENDEDKEY As Integer = &H1
        Public Const KEYEVENTF_KEYUP As Integer = &H2

        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Sub keybd_event(bVk As Byte, bScan As Byte, dwFlags As Integer, dwExtraInfo As IntPtr)
        End Sub

    End Class

End Namespace
#Enable Warning CA1401
