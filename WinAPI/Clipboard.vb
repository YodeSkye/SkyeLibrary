
Imports System.Runtime.InteropServices
Imports System.Text

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        Public Const WM_CHANGECBCHAIN As Integer = &H30D
        Public Const WM_DRAWCLIPBOARD As Integer = &H308
        Public Const WM_CLIPBOARDUPDATE As Integer = &H31D
        Public Const GMEM_MOVEABLE As UInteger = &H2
        Public Const CF_TEXT As UInteger = 1
        Public Const CF_BITMAP As UInteger = 2
        Public Const CF_OEMTEXT As UInteger = 7
        Public Const CF_UNICODETEXT As UInteger = 13
        Public Const CF_HDROP As UInteger = 15
        Public Const CF_DIB As UInteger = 8
        Public Const CF_DIBV5 As UInteger = 17
        Private Shared _CF_RTF As UInteger = 0
        Public Shared ReadOnly Property CF_RTF As UInteger
            Get
                If _CF_RTF = 0 Then
                    _CF_RTF = RegisterClipboardFormat("Rich Text Format")
                End If
                Return _CF_RTF
            End Get
        End Property
        Private Shared _CF_HTML As UInteger = 0
        Public Shared ReadOnly Property CF_HTML As UInteger
            Get
                If _CF_HTML = 0 Then
                    _CF_HTML = RegisterClipboardFormat("HTML Format")
                End If
                Return _CF_HTML
            End Get
        End Property


        ' API FUNCTIONS
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function AddClipboardFormatListener(hwnd As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetClipboardViewer(hWndNewViewer As IntPtr) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function ChangeClipboardChain(hWndRemove As IntPtr, hWndNewNext As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetClipboardSequenceNumber() As UInteger
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function OpenClipboard(hwnd As IntPtr) As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function CloseClipboard() As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function EmptyClipboard() As Boolean
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function EnumClipboardFormats(format As UInteger) As UInteger
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetClipboardData(uFormat As UInteger) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetClipboardData(uFormat As UInteger, hMem As IntPtr) As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalAlloc(uFlags As UInteger, dwBytes As UIntPtr) As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalLock(hMem As IntPtr) As IntPtr
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalUnlock(hMem As IntPtr) As Boolean
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalSize(hMem As IntPtr) As Integer
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function GetClipboardFormatName(format As UInteger, lpszFormatName As StringBuilder, cchMaxCount As Integer) As Integer
        End Function
        <DllImport("user32.dll", CharSet:=CharSet.Unicode, SetLastError:=True)>
        Public Shared Function RegisterClipboardFormat(lpszFormat As String) As UInteger
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalFlags(hMem As IntPtr) As Integer
        End Function
        <DllImport("kernel32.dll", SetLastError:=True)>
        Public Shared Function GlobalFree(hMem As IntPtr) As IntPtr
        End Function

    End Class

End Namespace
#Enable Warning CA1401
