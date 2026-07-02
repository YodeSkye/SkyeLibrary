
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' API FUNCTIONS
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function GetDC(hWnd As IntPtr) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function ReleaseDC(hWnd As IntPtr, hDC As IntPtr) As Integer
        End Function
        <DllImport("gdi32.dll", SetLastError:=True)>
        Public Shared Function CreateCompatibleDC(hdc As IntPtr) As IntPtr
        End Function
        <DllImport("gdi32.dll", SetLastError:=True)>
        Public Shared Function DeleteDC(hdc As IntPtr) As Boolean
        End Function
        <DllImport("gdi32.dll", SetLastError:=True)>
        Public Shared Function SelectObject(hdc As IntPtr, h As IntPtr) As IntPtr
        End Function
        <DllImport("gdi32.dll", SetLastError:=True)>
        Public Shared Function DeleteObject(hObject As IntPtr) As Boolean
        End Function
        <DllImport("gdi32.dll", SetLastError:=True)>
        Public Shared Function CreateRectRgn(nLeftRect As Integer, nTopRect As Integer, nRightRect As Integer, nBottomRect As Integer) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetWindowRgn(hWnd As IntPtr, hRgn As IntPtr, bRedraw As Boolean) As Integer
        End Function

    End Class

End Namespace
#Enable Warning CA1401
