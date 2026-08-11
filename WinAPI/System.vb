
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        <StructLayout(LayoutKind.Sequential)>
        Public Structure LASTINPUTINFO
            Public cbSize As UInteger
            Public dwTime As UInteger
        End Structure

        ' API FUNCTIONS
        <DllImport("user32.dll")>
        Public Shared Function GetLastInputInfo(ByRef plii As LASTINPUTINFO) As Boolean
        End Function
        <DllImport("user32.dll")>
        Public Shared Function GetSysColor(nIndex As Integer) As Integer
        End Function
        ' --- Functions ---
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function SetTimer(hWnd As IntPtr, nIDEvent As IntPtr, uElapse As UInteger, lpTimerFunc As IntPtr) As IntPtr
        End Function
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function KillTimer(hWnd As IntPtr, uIDEvent As IntPtr) As Boolean
        End Function

        ' METHODS
        Public Shared Function GetIdleTime() As UInteger
            Dim lastInput As New LASTINPUTINFO()
            lastInput.cbSize = CUInt(Marshal.SizeOf(lastInput))
            If Not GetLastInputInfo(lastInput) Then Return 0
            Return CUInt(Environment.TickCount) - lastInput.dwTime
        End Function
        Public Shared Function GetRValue(color As Long) As Short
            Return CShort(color And &HFF&)
        End Function
        Public Shared Function GetGValue(color As Long) As Short
            Return CShort((color \ &H100&) And &HFF&)
        End Function
        Public Shared Function GetBValue(color As Long) As Short
            Return CShort((color \ &H10000) And &HFF&)
        End Function
        Public Shared Function GetSystemColor(nIndex As Integer) As Color
            Dim colorvalue As Integer = GetSysColor(nIndex)
            Return Color.FromArgb(255, GetRValue(colorvalue), GetGValue(colorvalue), GetBValue(colorvalue))
        End Function

    End Class

End Namespace
#Enable Warning CA1401
