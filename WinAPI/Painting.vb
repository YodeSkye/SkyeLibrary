
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        Public Const WM_PAINT As Integer = &HF
        Public Const WM_ERASEBKGND As Integer = &H14
        <StructLayout(LayoutKind.Sequential)>
        Public Structure PAINTSTRUCT
            Public hdc As IntPtr
            Public fErase As Boolean
            Public rcPaint As Rectangle
            Public fRestore As Boolean
            Public fIncUpdate As Boolean
            <MarshalAs(UnmanagedType.ByValArray, SizeConst:=32)>
            Public rgbReserved As Byte()
        End Structure

        ' API FUNCTIONS
        <DllImport("user32.dll")>
        Public Shared Function BeginPaint(hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As IntPtr
        End Function
        <DllImport("user32.dll")>
        Public Shared Function EndPaint(hWnd As IntPtr, ByRef lpPaint As PAINTSTRUCT) As Boolean
        End Function

    End Class

End Namespace
#Enable Warning CA1401
