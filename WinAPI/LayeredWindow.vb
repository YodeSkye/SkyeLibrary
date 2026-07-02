
Imports System.Runtime.InteropServices

#Disable Warning CA1401
Namespace Skye

    Partial Public Class WinAPI

        ' DECLARATIONS
        Public Const WS_EX_LAYERED As Integer = &H80000
        <StructLayout(LayoutKind.Sequential, Pack:=1)>
        Public Structure BLENDFUNCTION
            Public BlendOp As Byte
            Public BlendFlags As Byte
            Public SourceConstantAlpha As Byte
            Public AlphaFormat As Byte
        End Structure

        ' API FUNCTIONS
        <DllImport("user32.dll", SetLastError:=True)>
        Public Shared Function UpdateLayeredWindow(
            hwnd As IntPtr,
            hdcDst As IntPtr,
            ByRef pptDst As POINT,
            ByRef psize As SIZE,
            hdcSrc As IntPtr,
            ByRef pptSrc As POINT,
            crKey As Integer,
            ByRef pblend As BLENDFUNCTION,
            dwFlags As Integer) As Boolean
        End Function

    End Class

End Namespace
#Enable Warning CA1401
